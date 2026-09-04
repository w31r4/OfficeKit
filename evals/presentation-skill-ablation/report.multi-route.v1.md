# Presentation Skill 4 路线质量实验：研究报告

生成时间：2026-09-03T06:47:35.950Z

> 当前基线 Skill、Shared What/What-kind/How、Kimi-style concise 与新增混合路线共用同一 PPJ、素材、渲染和复核能力；本报告只比较入口路由。

为保持冻结运行结果可追溯，机器字段仍沿用历史 ID `current-production`；它只表示“当时的现有 Skill 快照”，不表示已经存在生产标准或默认路线。

## 样本

- 任务数：12；路线：current-production、shared-what-kind-how、kimi-concise、hybrid-short-contract；全路线作者集合：6。
- 所有路线的作者侧硬门槛均通过的任务：4。
- 盲评认为所有路线均通过视觉/内容硬门槛的回合：6；实际纳入质量差分的回合：6。
- 有效盲评记录：12。

## 路线胜负（仅合格配对）

- current-production：1。
- shared-what-kind-how：1。
- kimi-concise：3。
- hybrid-short-contract：1。
- tie：0。

未过滤硬门槛的评审胜负仅作为诊断保留：current-production=2，shared-what-kind-how=4，kimi-concise=5，hybrid-short-contract=1，tie=0。
这些诊断胜负不进入质量结论。


## 两两质量分差（左路线 − 右路线）

| 比较 | n | 均值 | 中位数 | bootstrap 95% | p(sign) |
| --- | ---: | ---: | ---: | --- | ---: |
| current-production − shared-what-kind-how | 6 | -5.166666666666667 | -8 | -11.666666666666666 … 2.3333333333333335 | 0.6875 |
| current-production − kimi-concise | 6 | -8.666666666666666 | -13.5 | -16.833333333333332 … 1.3333333333333333 | 0.375 |
| current-production − hybrid-short-contract | 6 | -2.3333333333333335 | -4 | -6.833333333333333 … 2.6666666666666665 | 0.6875 |
| shared-what-kind-how − kimi-concise | 6 | -3.5 | -5 | -8.5 … 2.1666666666666665 | 0.6875 |
| shared-what-kind-how − hybrid-short-contract | 6 | 2.8333333333333335 | 1.5 | -4 … 10.666666666666666 | 1 |
| kimi-concise − hybrid-short-contract | 6 | 6.333333333333333 | 8 | -1 … 12.666666666666666 | 0.6875 |

## 分层结果

以下只在作者侧和盲评侧所有路线均通过硬门槛的回合中统计。

| 分层 | 比较 | n | 均值 | 中位数 |
| --- | --- | ---: | ---: | ---: |
| lifecycle=0-to-1 | current-production − shared-what-kind-how | 6 | -5.166666666666667 | -8 |
| lifecycle=0-to-1 | current-production − kimi-concise | 6 | -8.666666666666666 | -13.5 |
| lifecycle=0-to-1 | current-production − hybrid-short-contract | 6 | -2.3333333333333335 | -4 |
| lifecycle=0-to-1 | shared-what-kind-how − kimi-concise | 6 | -3.5 | -5 |
| lifecycle=0-to-1 | shared-what-kind-how − hybrid-short-contract | 6 | 2.8333333333333335 | 1.5 |
| lifecycle=0-to-1 | kimi-concise − hybrid-short-contract | 6 | 6.333333333333333 | 8 |
| lifecycle=1-to-10 | current-production − shared-what-kind-how | 0 | pending | pending |
| lifecycle=1-to-10 | current-production − kimi-concise | 0 | pending | pending |
| lifecycle=1-to-10 | current-production − hybrid-short-contract | 0 | pending | pending |
| lifecycle=1-to-10 | shared-what-kind-how − kimi-concise | 0 | pending | pending |
| lifecycle=1-to-10 | shared-what-kind-how − hybrid-short-contract | 0 | pending | pending |
| lifecycle=1-to-10 | kimi-concise − hybrid-short-contract | 0 | pending | pending |
| scenario=academic-research | current-production − shared-what-kind-how | 2 | 6 | 6 |
| scenario=academic-research | current-production − kimi-concise | 2 | 6 | 6 |
| scenario=academic-research | current-production − hybrid-short-contract | 2 | 1.5 | 1.5 |
| scenario=academic-research | shared-what-kind-how − kimi-concise | 2 | 0 | 0 |
| scenario=academic-research | shared-what-kind-how − hybrid-short-contract | 2 | -4.5 | -4.5 |
| scenario=academic-research | kimi-concise − hybrid-short-contract | 2 | -4.5 | -4.5 |
| scenario=analysis-decision | current-production − shared-what-kind-how | 0 | pending | pending |
| scenario=analysis-decision | current-production − kimi-concise | 0 | pending | pending |
| scenario=analysis-decision | current-production − hybrid-short-contract | 0 | pending | pending |
| scenario=analysis-decision | shared-what-kind-how − kimi-concise | 0 | pending | pending |
| scenario=analysis-decision | shared-what-kind-how − hybrid-short-contract | 0 | pending | pending |
| scenario=analysis-decision | kimi-concise − hybrid-short-contract | 0 | pending | pending |
| scenario=brand-creative | current-production − shared-what-kind-how | 0 | pending | pending |
| scenario=brand-creative | current-production − kimi-concise | 0 | pending | pending |
| scenario=brand-creative | current-production − hybrid-short-contract | 0 | pending | pending |
| scenario=brand-creative | shared-what-kind-how − kimi-concise | 0 | pending | pending |
| scenario=brand-creative | shared-what-kind-how − hybrid-short-contract | 0 | pending | pending |
| scenario=brand-creative | kimi-concise − hybrid-short-contract | 0 | pending | pending |
| scenario=education-training | current-production − shared-what-kind-how | 2 | -13.5 | -13.5 |
| scenario=education-training | current-production − kimi-concise | 2 | -13.5 | -13.5 |
| scenario=education-training | current-production − hybrid-short-contract | 2 | -0.5 | -0.5 |
| scenario=education-training | shared-what-kind-how − kimi-concise | 2 | 0 | 0 |
| scenario=education-training | shared-what-kind-how − hybrid-short-contract | 2 | 13 | 13 |
| scenario=education-training | kimi-concise − hybrid-short-contract | 2 | 13 | 13 |
| scenario=management-report | current-production − shared-what-kind-how | 2 | -8 | -8 |
| scenario=management-report | current-production − kimi-concise | 2 | -18.5 | -18.5 |
| scenario=management-report | current-production − hybrid-short-contract | 2 | -8 | -8 |
| scenario=management-report | shared-what-kind-how − kimi-concise | 2 | -10.5 | -10.5 |
| scenario=management-report | shared-what-kind-how − hybrid-short-contract | 2 | 0 | 0 |
| scenario=management-report | kimi-concise − hybrid-short-contract | 2 | 10.5 | 10.5 |
| scenario=technical-engineering | current-production − shared-what-kind-how | 0 | pending | pending |
| scenario=technical-engineering | current-production − kimi-concise | 0 | pending | pending |
| scenario=technical-engineering | current-production − hybrid-short-contract | 0 | pending | pending |
| scenario=technical-engineering | shared-what-kind-how − kimi-concise | 0 | pending | pending |
| scenario=technical-engineering | shared-what-kind-how − hybrid-short-contract | 0 | pending | pending |
| scenario=technical-engineering | kimi-concise − hybrid-short-contract | 0 | pending | pending |
| asset=image | current-production − shared-what-kind-how | 2 | -13.5 | -13.5 |
| asset=image | current-production − kimi-concise | 2 | -13.5 | -13.5 |
| asset=image | current-production − hybrid-short-contract | 2 | -0.5 | -0.5 |
| asset=image | shared-what-kind-how − kimi-concise | 2 | 0 | 0 |
| asset=image | shared-what-kind-how − hybrid-short-contract | 2 | 13 | 13 |
| asset=image | kimi-concise − hybrid-short-contract | 2 | 13 | 13 |
| asset=non-image | current-production − shared-what-kind-how | 4 | -1 | -2 |
| asset=non-image | current-production − kimi-concise | 4 | -6.25 | -8 |
| asset=non-image | current-production − hybrid-short-contract | 4 | -3.25 | -4 |
| asset=non-image | shared-what-kind-how − kimi-concise | 4 | -5.25 | -6 |
| asset=non-image | shared-what-kind-how − hybrid-short-contract | 4 | -2.25 | -2.5 |
| asset=non-image | kimi-concise − hybrid-short-contract | 4 | 3 | 0 |

## 效率

| 路线 | 作者数 | 平均 wall time (ms) | 平均输入 token | 平均工具调用 |
| --- | ---: | ---: | ---: | ---: |
| current-production | 12 | 1032112.5 | 7564695.125 | 150.66666666666666 |
| shared-what-kind-how | 12 | 1077783.4166666667 | 6552249.8 | 146.83333333333334 |
| kimi-concise | 12 | 1029989.25 | 7127431.428571428 | 134.33333333333334 |
| hybrid-short-contract | 6 | 1006425.8333333334 | 5443158 | 142.5 |

作者超时：18 / 48。
作者侧硬门槛失败：technical-engineering-01/hybrid-short-contract (failed)；brand-creative-01/shared-what-kind-how (failed)；analysis-decision-10/hybrid-short-contract (missing)；management-report-10/kimi-concise (failed)；management-report-10/hybrid-short-contract (missing)；technical-engineering-10/hybrid-short-contract (missing)；academic-research-10/shared-what-kind-how (failed)；academic-research-10/hybrid-short-contract (missing)；education-training-10/hybrid-short-contract (missing)；brand-creative-10/shared-what-kind-how (failed)；brand-creative-10/hybrid-short-contract (missing)。
盲评侧视觉/内容硬门槛失败：analysis-decision-01/r1/hybrid-short-contract (failed)；analysis-decision-01/r2/hybrid-short-contract (failed)；technical-engineering-01/r1/hybrid-short-contract (failed)；technical-engineering-01/r2/hybrid-short-contract (failed)；brand-creative-01/r1/shared-what-kind-how (failed)；brand-creative-01/r2/shared-what-kind-how (failed)。

## 结论（探索性）

- 这是一轮新增混合路线的探索性四方比较；质量差分同时要求作者侧和盲评侧硬门槛通过，不能用未通过门槛的外观分数抵消工程失败。
- 路线胜负（合格回合）为：current-production=1、shared-what-kind-how=1、kimi-concise=3、hybrid-short-contract=1、tie=0。
- 每一对路线的均值、bootstrap 区间和配对检验见上表；样本仍是冻结的 12 个任务，不能外推到所有 PPT 场景。
- 新混合路线的价值判断应同时看质量、硬门槛、编辑保真和成本；如果只在 0→1 获胜而 1→10 退化，不应直接替换当前基线。


## 限制

- 4 方盲评只把作者侧和盲评侧硬门槛都通过的回合纳入 pairwise 质量分差
- Structural/render evidence is not PowerPoint playback evidence
- Human calibration records are pending unless supplied separately
