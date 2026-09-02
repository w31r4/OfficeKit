# Presentation Skill 三路线质量实验：研究报告

生成时间：2026-09-02T20:53:07.248Z

> 当前生产 Skill 是冻结控制组；Shared What/What-kind/How 与 Kimi-style concise 是实验组。

## 样本

- 任务数：12；三方作者集合：12。
- 三方均通过硬门槛的任务：8。
- 有效盲评记录：24。

## 路线胜负（仅合格配对）

- current-production：3。
- shared-what-kind-how：3。
- kimi-concise：7。
- tie：3。

未过滤硬门槛的评审胜负仅作为诊断保留：current-production=8，shared-what-kind-how=3，kimi-concise=10，tie=3。
这些诊断胜负不进入质量结论。


## 两两质量分差（左路线 − 右路线）

| 比较 | n | 均值 | 中位数 | bootstrap 95% | p(sign) |
| --- | ---: | ---: | ---: | --- | ---: |
| current-production − shared-what-kind-how | 16 | -1 | 0 | -6 … 4.3125 | 1 |
| current-production − kimi-concise | 16 | -3.1875 | -2 | -7.875 … 1.6875 | 0.5810546875 |
| shared-what-kind-how − kimi-concise | 16 | -2.1875 | -0.5 | -8.75 … 4.0625 | 0.5810546875 |

## 分层结果

以下只在三方均通过硬门槛的回合中统计。

| 分层 | 比较 | n | 均值 | 中位数 |
| --- | --- | ---: | ---: | ---: |
| lifecycle=0-to-1 | current-production − shared-what-kind-how | 10 | -5.6 | -7 |
| lifecycle=0-to-1 | current-production − kimi-concise | 10 | -4.5 | -6.5 |
| lifecycle=0-to-1 | shared-what-kind-how − kimi-concise | 10 | 1.1 | -0.5 |
| lifecycle=1-to-10 | current-production − shared-what-kind-how | 6 | 6.666666666666667 | 3.5 |
| lifecycle=1-to-10 | current-production − kimi-concise | 6 | -1 | 0 |
| lifecycle=1-to-10 | shared-what-kind-how − kimi-concise | 6 | -7.666666666666667 | -0.5 |
| scenario=academic-research | current-production − shared-what-kind-how | 2 | 5 | 5 |
| scenario=academic-research | current-production − kimi-concise | 2 | 9.5 | 9.5 |
| scenario=academic-research | shared-what-kind-how − kimi-concise | 2 | 4.5 | 4.5 |
| scenario=analysis-decision | current-production − shared-what-kind-how | 4 | -6.75 | -4.5 |
| scenario=analysis-decision | current-production − kimi-concise | 4 | -4.25 | -2 |
| scenario=analysis-decision | shared-what-kind-how − kimi-concise | 4 | 2.5 | 0 |
| scenario=brand-creative | current-production − shared-what-kind-how | 0 | pending | pending |
| scenario=brand-creative | current-production − kimi-concise | 0 | pending | pending |
| scenario=brand-creative | shared-what-kind-how − kimi-concise | 0 | pending | pending |
| scenario=education-training | current-production − shared-what-kind-how | 4 | 7.75 | 7.5 |
| scenario=education-training | current-production − kimi-concise | 4 | -11.25 | -9.5 |
| scenario=education-training | shared-what-kind-how − kimi-concise | 4 | -19 | -17 |
| scenario=management-report | current-production − shared-what-kind-how | 2 | -8 | -8 |
| scenario=management-report | current-production − kimi-concise | 2 | -14.5 | -14.5 |
| scenario=management-report | shared-what-kind-how − kimi-concise | 2 | -6.5 | -6.5 |
| scenario=technical-engineering | current-production − shared-what-kind-how | 4 | -3.5 | -2.5 |
| scenario=technical-engineering | current-production − kimi-concise | 4 | 5.25 | 6.5 |
| scenario=technical-engineering | shared-what-kind-how − kimi-concise | 4 | 8.75 | 6.5 |
| asset=image | current-production − shared-what-kind-how | 4 | -5.75 | -7 |
| asset=image | current-production − kimi-concise | 4 | -4.5 | -3 |
| asset=image | shared-what-kind-how − kimi-concise | 4 | 1.25 | 1 |
| asset=non-image | current-production − shared-what-kind-how | 12 | 0.5833333333333334 | 0 |
| asset=non-image | current-production − kimi-concise | 12 | -2.75 | -2 |
| asset=non-image | shared-what-kind-how − kimi-concise | 12 | -3.3333333333333335 | -0.5 |

## 效率

| 路线 | 作者数 | 平均 wall time (ms) | 平均输入 token | 平均工具调用 |
| --- | ---: | ---: | ---: | ---: |
| current-production | 12 | 1032112.5 | 7564695.125 | 150.66666666666666 |
| shared-what-kind-how | 12 | 1077783.4166666667 | 6552249.8 | 146.83333333333334 |
| kimi-concise | 12 | 1029989.25 | 7127431.428571428 | 134.33333333333334 |

作者超时：16 / 36。
硬门槛失败：brand-creative-01/shared-what-kind-how (failed)；management-report-10/kimi-concise (failed)；academic-research-10/shared-what-kind-how (failed)；brand-creative-10/shared-what-kind-how (failed)。

## 结论（探索性）

- 没有足够证据宣布单一路线胜出：三组总体配对区间均跨过 0，exact paired permutation 的双侧 p 值也未达到稳定差异标准。
- 合格回合的路线胜负为 current-production=3、shared-what-kind-how=3、kimi-concise=7、tie=3；这比包含失败产物的诊断胜负更适合作为结论依据。
- 0→1 中，两条实验路线相对当前生产路线的分差点估计为负（current−shared=-5.6，current−kimi=-4.5），说明新入口在从零创作上有潜力；shared 与 kimi 的差异很小（+1.1）。
- 1→10 中，shared 相对 current 落后（+6.67），kimi 与 current 接近（-1.0），更像是编辑/源绑定可靠性和任务完成度的差异，而不是单纯视觉偏好。
- brand-creative 的两项任务没有三方同时过门槛，因此本轮不能对品牌场景下结论；四个硬门槛失败必须作为工程问题处理。
- 暂不改变生产默认路由：可把 kimi-concise 作为后续 0→1 优化候选，把 current-production 保留为稳定控制/编辑回退，把 shared 路线先修复失败的源绑定和超时问题后再复测。


## 限制

- 三方盲评只把三方作者都通过硬门槛的回合纳入 pairwise 质量分差
- Structural/render evidence is not PowerPoint playback evidence
- Human calibration records are pending unless supplied separately
