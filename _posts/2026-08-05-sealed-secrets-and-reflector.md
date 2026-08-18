---
title: GitOps 里 Secret 安全提交与跨 namespace 分发
date: 2026-08-05 11:00:00 +0800
categories: [Blogging, DevOps, Kubernetes]
tags: [kubernetes, gitops, secrets, sealed-secrets]
description: GitOps 要求所有部署状态都进 Git，但数据库密码、API Key 这类 Secret 明文进 Git 是绝对不能接受的。这篇文章讲清楚我们怎么用 Sealed Secrets 把加密后的 Secret 安全地提交到仓库、加密范围（scope）怎么选，以及同一份配置要在多个业务 namespace 里都有一份时怎么用 Reflector 自动分发，而不是手动复制。
---

[上一篇](../gitops-pipeline-design-decisions/) 写完 GitOps 流水线之后，有个问题一直没展开：GitOps 的核心规则是"仓库里写的就是集群里跑的"，那 Secret 也要进仓库吗？数据库密码、第三方 API Key，这些东西明文写进 Git，哪怕仓库是私有的，也是在给自己埋雷——权限管理出一次纰漏，或者仓库有一天要迁移、要开放给更多人协作，明文 Secret 就是定时炸弹。

但如果 Secret 不进仓库，GitOps"一切状态可追溯"的承诺就破了一个大洞——集群里到底跑着哪个密码，只有手动 `kubectl apply` 的人知道，Git 历史里看不出来。这篇文章讲的是怎么两头都要：Secret 可以安全地躺在 Git 里，同时集群状态依然完全可追溯。

## 思路：不加密 Secret 本身，加密"谁能读它"这件事

Sealed Secrets（Bitnami 出的一个 CRD + controller）的思路很直接：Secret 内容用非对称加密加密一遍，生成一个新的 K8s 资源叫 `SealedSecret`，这个资源可以放心提交到 Git——因为没有私钥，谁都解不开。集群里跑着一个 controller，手里握着唯一的私钥，看到 `SealedSecret` 被 apply 进来就自动解密，生成一个真正的 `Secret` 资源。公钥不需要保密，可以随便分发；私钥只在这一个集群里，从不出集群。

```bash
echo -n bar | kubectl create secret generic mysecret \
  --dry-run=client --from-file=foo=/dev/stdin -o yaml -n testing > mysecret.yaml
kubeseal -f mysecret.yaml -w seal-secret.yaml -o yaml
kubectl create -f seal-secret.yaml
```

流程上其实就是"正常方式生成一份 Secret 的 YAML，过一遍 kubeseal 再落盘"，多了一步，但换来的是这份产物可以安全提交。`kubeseal` 本身不需要连集群也能用——先 `kubeseal --fetch-cert` 把公钥导出成一个 `.pem` 文件，之后离线加密都靠这份公钥就够了，我们仓库里公钥文件就跟其它部署文件放在一起，这是唯一可以公开的那一半。

## 加密范围怎么选，决定了这份 Secret 能不能被"认出来"

这里有个容易被忽略但很重要的点：加密不只是加密内容，Secret 的**名字和所在的 namespace 默认也参与加密**。也就是说，同样的密码内容，加密成 `secret-a` 在 `testing` namespace 下产出的密文，和加密成 `secret-b` 在 `production` namespace 下产出的密文完全不同——你不能把一份 `SealedSecret` 复制粘贴改个名字或者换个 namespace 就直接用，必须重新加密。

这是默认的 `strict` scope，好处是最安全——密文被泄露出去也只能在指定的 名字+namespace 下被解密，攻击面最小。但代价是：一份配置要更新某个 key，或者要合并新的 key 进去，操作起来很啰嗦。所以像我们这种需要频繁往同一个 Secret 里加字段的场景（一个统一的应用配置 Secret，里面塞了十几个 key，各种数据库密码、第三方服务凭证），用的是 `namespace-wide` scope：

```yaml
apiVersion: bitnami.com/v1alpha1
kind: SealedSecret
metadata:
  name: application-global-secret
  namespace: production
  annotations:
    sealedsecrets.bitnami.com/namespace-wide: "true"
```

这个 scope 下，只要目标 namespace 对，Secret 叫什么名字不影响解密——换句话说，安全边界从"绑死这一个 Secret"放宽到了"绑死这一个 namespace"。选它是权衡过的：这份配置本来就是同一个 namespace 内部共享的，放宽到 namespace 级别不会引入新的越权风险，但换来的是可以随手 `kubeseal --merge-into` 往里面加字段，不用每次都重新生成整份密文。加密范围不是越严格越好，是要跟这份 Secret 实际的使用方式对上。

## 命令行不是唯一入口

`kubeseal` 命令行工具功能齐全，但要求本地装好客户端、还要理解 scope、merge 这些概念，不是所有需要改配置的人都愿意折腾这套。我们额外部署了一个 web UI（`sealed-secrets-ui`），暴露一个页面直接在浏览器里加密——本质上是把 `kubeseal --fetch-cert` 之后的加密逻辑包了一层界面，降低了"改一个配置项"这件事的门槛。工具选型这里的判断很简单：核心加密机制没必要重新发明，围绕它的使用体验值得投入。

## 同一份配置要出现在好几个 namespace，别手动复制

Sealed Secrets 解决的是"能不能安全提交"，但没解决另一个问题：如果同一份配置（不一定是 Secret，也可能是 ConfigMap）本来就该在好几个 namespace 里都存在一份，比如我们有个日志采集组件的配置，需要在 `production`、`production-tenant-a`、`production-tenant-b`、`production-tenant-c` 好几个业务 namespace 里都跑一份完全一样的配置——如果每个 namespace 手动维护一份，改一次要改好几处，跟最开始"几十份 Jenkinsfile"是同一类问题：同一份东西被复制了 N 份，改的时候必然会漏一个。

这里用的是 Reflector：在源资源上打几个 annotation，声明"这份东西允许被复制，允许复制到哪些 namespace"，Reflector 的 controller 会自动把它同步过去，源资源变了目标也跟着变，不需要任何人手动 `kubectl apply` 第二遍。

```yaml
metadata:
  annotations:
    reflector.v1.k8s.emberstack.com/reflection-allowed: "true"
    reflector.v1.k8s.emberstack.com/reflection-auto-enabled: "true"
    reflector.v1.k8s.emberstack.com/reflection-allowed-namespaces: "staging,production,production-tenant-a,production-tenant-b,production-tenant-c"
```

需要说明一下：Reflector 同时支持复制 ConfigMap 和 Secret，但我们目前只拿它复制过 ConfigMap（上面这个日志配置的例子），没有真的用它去分发 Secret。这不是因为技术上不行，纯粹是还没遇到"同一份 Secret 内容需要原样出现在多个 namespace"的真实场景——目前每套环境的 Secret 都是各自单独加密提交的。如果以后真出现这种需求，Reflector 和 Sealed Secrets 是可以配合起来用的：先把 Secret 解密到一个"源" namespace，再用 Reflector 分发出去，而不是对每个目标 namespace 都重新跑一遍加密流程。

## 小结

这套方案解决的是两个不同层次的问题：Sealed Secrets 回答"敏感配置能不能安全地进 Git"，靠的是非对称加密把"内容"和"解密权限"彻底分开，私钥只留在集群里；Reflector 回答"同一份配置要在多个 namespace 出现一份，怎么避免手动复制"，跟前面 GitOps 流水线里"改一处要改几十遍"是同一个思路的另一种体现——只要是"同一份东西需要存在多份拷贝"，答案基本都是让机器去同步，不要让人去复制粘贴。scope 怎么选、要不要上 Reflector，都不是无脑照抄文档的默认值，是要看这份配置实际怎么被使用、被更新，再决定该收紧还是该放宽。
