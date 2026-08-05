---
title: 私有化 AI 模型性能与准确率实测
date: 2025-07-23 09:00:00 +0800
categories: [Blogging, AI]
tags: [ai, llm]
description: 基于 vLLM 部署 Qwen3、DeepSeek 等 Chat 与 Reranker 模型，记录 JMeter 并发压测、响应延迟及 Tool Calling 准确率测试结果。
---

# 测试环境

- 模型推理服务：vLLM
- 性能压测工具：JMeter
- GPU：NVIDIA L20
- JMeter 用户数：与表格中的并发数一致
- 循环次数：每个用户 100 次
- Ramp-up：0 秒，所有用户同时启动

除特别说明外，本文中的私有化模型均使用 vLLM 部署，性能数据均由 JMeter 压测得到。

# Reranker

| 模型                                  | 参数                                                                                          | documents总字数 | GPU | 并发数 | 平均响应时间(ms) | P95(ms) | P99(ms) | 错误率(%) |
| ------------------------------------- | --------------------------------------------------------------------------------------------- | --------------- | --- | ------ | ---------------- | ------- | ------- | --------- |
| Qwen/Qwen3-Reranker-4B                |                                                                                               | 3500            | 2   | 50     | 5775             | 8952    | 8958    | 2         |
| tomaarsen/Qwen3-Reranker-0.6B-seq-cls |                                                                                               | 3500            | 4   | 50     | 2047             | 2551    | 2599    | 0         |
| tomaarsen/Qwen3-Reranker-0.6B-seq-cls |                                                                                               | 3500            | 4   | 100    | 990              | 1623    | 2252    | 0.1       |
| tomaarsen/Qwen3-Reranker-0.6B-seq-cls |                                                                                               | 3500            | 4   | 150    | 1361             | 2376    | 2986    | 0.3       |
| tomaarsen/Qwen3-Reranker-0.6B-seq-cls |                                                                                               | 3500            | 4   | 200    | 1504             | 2263    | 2651    | 1.95      |
| tomaarsen/Qwen3-Reranker-4B-seq-cls   |                                                                                               | 3500            | 4   | 50     | 6979             | 8642    | 8657    | 0         |
| tomaarsen/Qwen3-Reranker-4B-seq-cls   | \--enable-prefix-caching                                                                      | 3500            | 4   | 50     | 6975             | 8620    | 8628    | 0         |
| tomaarsen/Qwen3-Reranker-4B-seq-cls   | \--enforce-eager                                                                              | 3500            | 4   | 50     | 7020             | 8697    | 8712    | 0         |
| tomaarsen/Qwen3-Reranker-4B-seq-cls   | \--tensor-parallel-size 1  --task score --data-parallel-size 4                                | 3500            | 4   | 50     | 2258             | 2657    | 2693    | 0         |
| tomaarsen/Qwen3-Reranker-4B-seq-cls   | \--tensor-parallel-size 1  --task score --data-parallel-size 4  --max-num-batched-tokens 8096 | 3500            | 4   | 50     | 2261             | 2651    | 2671    | 0         |
| tomaarsen/Qwen3-Reranker-0.6B-seq-cls | \--tensor-parallel-size 1  --task score --data-parallel-size 4                                | 3500            | 4   | 50     | 401              | 564     | 768     | 0         |
| tomaarsen/Qwen3-Reranker-8B-seq-cls   |                                                                                               | 3500            | 4   | 50     | 8190             | 10036   | 10042   | 39        |
| BAAI/bge-reranker-v2-m3               |                                                                                               | 3500            | 4   | 50     | 2144             | 2526    | 2547    | 0         |


# Chat

deepseek-v3 19996条数据, Tool Calling 命中率 86.89%

Qwen/Qwen3-235B-A22B-Instruct-2507-FP8  8卡部署, 单次请求要16秒, 太慢了

阿里云百炼 qwen3-32b no\_thinking 调用工具决策要6秒

百度千帆 deepseek-v3 决策要 7.93秒

Qwen/Qwen3-32B no\_think Temperature 0，测试 19996 条数据，Tool Calling 直接命中率为 91%；另有部分结果的 JSON 内容正确但符号缺失，因日志丢失无法准确统计，最后一次记录为 2%。

> `+` 后面的数值表示 JSON 内容正确，但因 JSON 符号缺失导致格式不正确。

Qwen/Qwen3-30B-A3B-Instruct-2507 Temperature 0，测试 2000 条数据，其中约一半为单问题，另一半为多问题。Tool Calling 命中率为 91.95% + 2.65% = 94.60%，其中 2.65% 的结果 JSON 内容正确，但因符号缺失导致格式不正确。

Qwen/Qwen3-30B-A3B-Instruct-2507 使用 2 个 GPU 核心部署时出现以下错误，将 `--max-model-len` 调整为 `102400` 后即可正常启动。

> To serve at least one request with the models's max seq len (262144), (12.00 GiB KV cache is needed, which is larger than the available KV cache memory (10.56 GiB)

## Tool Calling 成功率

只换 Qwen/Qwen3-235B-A22B-FP8, 其他什么都不换, 成功率只有 56%. 而且 `tool_call` 生成不出正确的 JSON

### 参数 Temperature 0.6

Qwen/Qwen3-235B-A22B-FP8 no\_think 测试 100 条数据，Tool Calling 命中率为 61% + 28% = 89%，其中 28% 的结果 JSON 内容正确，但因符号缺失导致格式不正确。Qwen/Qwen3-235B-A22B-FP8 thinking 测试 100 条数据，Tool Calling 命中率为 47% + 46% = 93%，其中 46% 的结果 JSON 内容正确，但因符号缺失导致格式不正确。

Qwen/Qwen3-32B no\_think 测试 100 条数据，Tool Calling 命中率为 90% + 9% = 99%，其中 9% 的结果 JSON 内容正确，但因符号缺失导致格式不正确。Qwen/Qwen3-32B thinking 测试 100 条数据，Tool Calling 命中率为 86% + 10% = 96%，其中 10% 的结果 JSON 内容正确，但因符号缺失导致格式不正确。

Qwen/Qwen3-30B-A3B-Instruct-2507 测试 2000 条数据，Tool Calling 命中率为 92.4% + 3.35% = 95.75%，其中 3.35% 的结果 JSON 内容正确，但因符号缺失导致格式不正确。

### 参数 Temperature 0

Qwen/Qwen3-235B-A22B-FP8 no\_think 测试 100 条数据，Tool Calling 命中率为 66% + 24% = 90%，其中 24% 的结果 JSON 内容正确，但因符号缺失导致格式不正确。Qwen/Qwen3-235B-A22B-FP8 thinking 测试 100 条数据，Tool Calling 命中率为 54% + 37% = 91%，其中 37% 的结果 JSON 内容正确，但因符号缺失导致格式不正确。

Qwen/Qwen3-32B no\_think 测试 100 条数据，Tool Calling 命中率为 94% + 2% = 96%，其中 2% 的结果 JSON 内容正确，但因符号缺失导致格式不正确。Qwen/Qwen3-32B thinking 测试 100 条数据，Tool Calling 命中率为 94% + 2% = 96%，其中 2% 的结果 JSON 内容正确，但因符号缺失导致格式不正确。

Qwen/Qwen3-30B-A3B-Instruct-2507 测试 2000 条数据，Tool Calling 命中率为 92.65% + 2.65% = 95.3%，其中 2.65% 的结果 JSON 内容正确，但因符号缺失导致格式不正确。

Qwen/Qwen3-32B 准确率比 Qwen/Qwen3-235B-A22B-FP8 高

结论: 使用Qwen/Qwen3-32B 非思考模式和推荐默认参数 Temperature 0
