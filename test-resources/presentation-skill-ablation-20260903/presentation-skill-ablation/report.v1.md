# Presentation Skill 双路线质量实验：研究报告

生成时间：2026-09-03T14:06:09.142Z

> 本报告只解释冻结的 12 个配对任务，不改变生产默认路线，也不把结构渲染证据描述成真实 PowerPoint 播放通过。

## 样本与硬门槛

- 任务数：4；两套 Skill；每个场景各一项 0→1 和 1→10。
- 双方都通过硬门槛的配对数：2。
- 已解析的盲评记录：8。

## 质量结果

- 盲评标签胜负（仅审计轨迹）：A 2，B 2，平局 4。
- 路线胜负（恢复随机映射后）：Shared 2，Kimi 2，平局 4。
- 有效配对分差（Shared−Kimi）均值：0；中位数：0。
- 精确符号检验：{"n":0,"positive":0,"negative":0,"pTwoSided":1}。
- 精确配对符号翻转检验：{"n":4,"method":"exact-sign-flip","combinations":16,"pTwoSided":1}。
- bootstrap 95% 区间：{"n":4,"low":0,"high":0,"median":0}。

## 按生命周期

| 生命周期 | 样本 | 均值分差 Shared−Kimi | 中位数分差 | bootstrap 95% |
| --- | ---: | ---: | ---: | --- |
| 0-to-1 | 0 | pending | pending | pending … pending |
| 1-to-10 | 4 | 0 | 0 | 0 … 0 |

## 按场景

| 场景 | 样本 | 均值分差 Shared−Kimi | 中位数分差 |
| --- | ---: | ---: | ---: |
| academic-research | 2 | 0 | 0 |
| analysis-decision | 2 | 0 | 0 |
| brand-creative | 0 | pending | pending |
| management-report | 0 | pending | pending |

## 按图片任务

| 分层 | 样本 | 均值分差 Shared−Kimi | 中位数分差 |
| --- | ---: | ---: | ---: |
| image | 0 | pending | pending |
| non-image | 4 | 0 | 0 |

## 效率（不并入质量分）

| 路线 | 作者数 | 平均 wall time (ms) | 平均输入 token | 平均工具调用 | 平均图片搜索 |
| --- | ---: | ---: | ---: | ---: | ---: |
| shared-what-kind-how | 4 | 479934.5 | pending | pending | pending |
| kimi-concise | 4 | 531392 | pending | pending | pending |

## 限制

- 冻结设计为 n=12 探索性配对样本；当前已完成作者配对 4/12，盲评记录 8/24
- No default Skill or production route is selected
- Structural/render evidence is not PowerPoint playback evidence
- Human calibration records are pending unless supplied separately

原始运行记录位于本次 runRoot；只将本汇总和冻结研究方法纳入版本控制。

