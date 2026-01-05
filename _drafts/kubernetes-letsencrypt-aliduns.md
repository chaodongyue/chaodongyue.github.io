---
title: 在Kubernetes内集成Let's Encrypt和阿里云DNS
# date: 2025-12-15 09:00:00 +0800
categories: [Blogging, DevOps, Kubernetes]
tags: [devops, kubernetes]
description: TODO
---

由于互联网对于安全的要求越来越严, 有很多中间件即使是部署在内网, 它也强制要求你使用HTTPS协议.

幸好有另外一种免费获取到SSL证书的方式, 就是使用Let's Encrypt.

下面介绍我在内网里使用 Kubernetes + Let's Encrypt + 阿里云DNS 部署可信HTTPS.

## 目标
现在目标可转变为在Kubernetes 部署 Let's Encrypt 并且能自动续期.

## 步骤拆解
1. 让证书到期能自动续期
2. 让Let's Encrypt能检验域名拥有权

### 证书自动续期
在 Kubernetes 里, 可以使用 [cert-manager](https://cert-manager.io/) 管理证书以及证书的自动续期.

由于cert-manager 内置的DNS供应商是不支持阿里做DNS的, 详细请看 [Supported DNS01 providers](https://cert-manager.io/docs/configuration/acme/dns01/#supported-dns01-providers).幸好cert-manager 的 Webhook 扩展了 cert-manager 的 DNS-01 验证所支持的 DNS 提供商, 已经有许多第三方实现, 详细列表和用法请参见 [Webhook](https://cert-manager.io/docs/configuration/acme/dns01/#webhook)

我使用的是 [AliDNS-Webhook](https://github.com/pragkent/alidns-webhook)

[cert-manager-alidns-webhook](https://github.com/DEVmachine-fr/cert-manager-alidns-webhook) 这Webhook 是一家公司的, 比较正规活跃, 但我把它部署在cert-manager 所在的namespace时无法启动Pod, 但在default 下却又可以.跟我的预期不一致就放弃了.

####  安装Webhook
第一步先去阿里云申请AK SK



第二步创建Secret
```yaml
apiVersion: v1
kind: Secret
metadata:
  name: alidns-secret
  namespace: cert-manager
data:
  access-key: YOUR_ACCESS_KEY
  secret-key: YOUR_SECRET_KEY
```

第三步
