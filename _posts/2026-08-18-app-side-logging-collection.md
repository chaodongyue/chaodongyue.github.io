---
title: 应用日志采集：数据的位置决定采集方式
date: 2026-08-18 10:00:00 +0800
categories: [Blogging, DevOps, Kubernetes]
tags: [kubernetes, devops, observability, opentelemetry, logging]
description: 日志采不采得全、采不采得对，先看数据落在哪、写成什么格式。这篇按数据位置拆开讲：stdout、PVC 文件、中间件自管文件、网络直推，各自该用什么采、坑在哪。
---

[上一篇]({% post_url 2026-08-14-observability-lgtm-and-otel-collector %})讲了集群侧的采集架构，解决的是"数据怎么搬"。这篇讲数据本身落在哪，因为搬运方式是被这一点定死的。

| 数据落在哪       | 用什么采                                       | 关键点                                         |
| ---------------- | ---------------------------------------------- | ---------------------------------------------- |
| stdout           | OTel Collector DaemonSet + filelog receiver    | Collector 自带 checkpoint 续传                 |
| PVC 上的文件     | Pod 内 sidecar 跑 fluent-bit tail              | position db 必须单独持久化，否则重启后重复采集 |
| 中间件自管的文件 | 视中间件而定，通常也是 sidecar 或专用 exporter | 路径、格式、轮转策略不受业务团队控制           |
| 应用直推网络     | 内嵌 OTLP log exporter，直连 Gateway           | 没有磁盘兜底，可靠性压在 SDK 的重试/缓冲       |

## 格式：一行 JSON，还是换行文本

不管数据落在 stdout 还是文件，底层采集基本都是按行读的，"一行=一条记录"是隐含前提。一行 JSON（`{"level":"error","traceId":"abc",...}`）天然满足这个前提，字段都在结构里，采集端直接 parser 解析。普通换行文本容易违反这个前提，最典型的是异常堆栈——一条异常天然是好几行，按行采会被拆成几十条独立记录，想合并回去只能靠采集端猜"这行是不是新记录的开头"，格式一有例外就拼错。

结论：堆栈之类的多行内容在应用层序列化进 JSON 的一个字段里，一次性解决，不留给采集端猜。

## stdout

OTel Collector 以 DaemonSet 部署，每节点一份；节点上这个 Collector 实例里配的 filelog receiver 是具体干活的组件，对宿主机上容器运行时落的 stdout/stderr 文件做 tail 读取（类似 `tail -f`），不是网络接收。自带 checkpoint，Collector 重启不丢不重。应用侧只管把内容打对：结构化 JSON + 字段命名跨服务统一（`traceId` 别一个服务叫 `trace_id`，一个塞进 `extra` 里，Grafana 关联查询就废了）。

## PVC 上的文件

OTel DaemonSet Collector 扫的是宿主机标准日志目录，够不到 Pod 自己挂的 PVC，得在 Pod 里加 sidecar 跑 fluent-bit 去 tail。

续传靠 fluent-bit 自己的 position db（一个 sqlite 文件，记录每个文件读到了第几个字节）。这个 db 如果只在 sidecar 的可写层里，sidecar 一重启——Pod 重建或者 fluent-bit 自己异常退出——position 就没了，重启后从头 tail，日志整段在 Loki 里重复。这类故障不报错、不掉数据，只是重复，很难当场发现。

把 position db 路径单独挂一块持久化存储（同一个 PVC 下的子路径，或单独一个小 PV），跟日志文件本身的生命周期绑在一起，重启后接着上次的 offset 读。

## 中间件自管的文件

MySQL 慢查询、Redis 之类，文件路径、命名、轮转策略是中间件自己的机制，业务团队改不了。能推远端的用中间件自带的转发能力，不能推的靠 sidecar 或专用采集配置去读对应路径。这块没有统一做法，按中间件适配，成本提前接受。

## 应用直推网络

内嵌 OTLP log exporter，生成后直接推给 Gateway，中间没有文件这一环。省掉了文件轮转和 position 续传，代价是可靠性全压在 SDK：Gateway 短暂不可用时，有没有重试和本地缓冲直接决定丢不丢数据，没有"数据还在磁盘上，恢复了还能补采"的兜底。

## 小结

四种位置对应四种完全不同的采集机制和风险点：stdout 续传是 Collector 的事；PVC 续传是自己的事，挂错 position db 的地方是最容易踩的坑；中间件文件只能适配、没有优化空间；网络直推没有磁盘兜底，可靠性看 SDK。先分清数据落在哪，再谈用什么采。
