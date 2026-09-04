# Kimi concise v2 复测报告

这是一次真正重新执行的 v2 复测，不使用旧 Kimi 分数替代。控制臂 `shared-what-kind-how` 是 main 快照，另一臂是 Kimi concise v2。

## 运行冻结

- main 快照：`9b86939ae513055836ce470af6bcec2db1e3c6e0`。
- main Skill SHA-256：`8e595264e51bd54d897e3bd9dafafac6362a220dcec75f8ccb7a2706bc2c3b13`。
- Kimi concise v2 SHA-256：`df1a4e998991d91a259456ff670d7323a7a369c3e8922206e68848dad86bc31e`。
- 模型：`gpt-5.6-luna`，reasoning `max`；8 个作者会话、8 个匿名盲评会话。
- 四个此前低分的 1→10 场景；两臂使用同一份预先投影 PPJ，避免把导入开销混入比较。

## 结论

**Kimi concise v2 在这四个旧低分场景中没有拿到高分，也没有证明优于 main。**按 85 分作为高分线：严格合格配对中 Kimi v2 为 `0/4`，main 也是 `0/4`；两者在 4 个合格盲评回合中全部同分。

这轮主要暴露了测试输入和交付可靠性问题：两个场景的目标表面不存在而双方合理拒绝修改，一个场景 Kimi v2 的 PPJ check 超时失败，一个场景 main 没有输出 PPTX。不能据此声称 v2 审美已提升或已退化。

## 逐场景

| 场景 | main 严格门槛 | Kimi v2 严格门槛 | main 盲评中位数 | Kimi v2 盲评中位数 | 说明 |
| --- | --- | --- | ---: | ---: | --- |
| analysis-decision-10 | passed | passed | 59.5 | 59.5 | 指定页没有图表或 endpoint，双方 no-op，像素相同 |
| management-report-10 | passed | failed | 71 | 20 | Kimi v2 check 报 `ppj.schema.pattern` 有界时间预算错误 |
| academic-research-10 | passed | passed | 76.5 | 76.5 | 指定页没有 note/result/table/chart surface，双方 no-op，像素相同 |
| brand-creative-10 | failed | passed | 20 | 71 | main 没有最终 PPTX；Kimi v2 有完整输出 |

分数来自两轮匿名 Luna Max 盲评的固定七维量表；失败产物的分数只作为诊断，不进入严格质量结论。

## 分数与胜负

- 严格合格回合：4（analysis 2 + academic 2）。
- 严格合格质量分：main 中位数 `68.5`，Kimi v2 中位数 `68.5`；均值均为 `68.0`。
- 严格合格胜负：main `0`、Kimi v2 `0`、平局 `4`。
- 含失败产物的诊断分数：两臂中位数均为 `67.5`；只用于解释盲评，没有进入质量结论。
- Kimi v2 达到 85 分：严格 `0/4`，诊断 `0/8`；最高回合为学术场景 `80`。
- 作者平均 wall time：main `479934.5 ms`；Kimi v2 `531392 ms`。token 未成功从 Codex 事件稳定采集，不能伪造成本结论。

## 失败原因

1. `analysis-decision-10` 和 `academic-research-10` 的 brief 与指定页不匹配，页上没有要编辑的 chart、endpoint、note、result、table surface。两套 Skill 都遵守了不发明事实、不把编辑挪到别页，所以 PPJ 与 PPTX 保持 byte-identical。
2. `management-report-10` 的 Kimi v2 产物 build、render、reimport 证据存在，但 `ppj check` 报 `ppj.schema.pattern: String pattern validation exceeded its bounded time budget`，这是硬门槛失败。
3. `brand-creative-10` 的 main agent 只留下 PPJ，没有最终 PPTX，因而不能作为交付物与 Kimi v2 公平比较。

## 证据

- 原始运行目录：[focused-low-v2-20260903](/Users/zfang/workspace/officekit-main-skill-eval-20260903/evals/presentation-skill-ablation/runs/focused-low-v2-20260903)。
- 盲评目录：`blind-review/round-1`、`blind-review/round-2`。
- 作者记录：`authors/<case>/<arm>/evidence/author-run.json`。
- 这次运行没有修改仓库源码或 main 快照。

## 下一步

如果要判断 Kimi concise v2 是否真的能把旧低分场景拉高，下一轮必须先修正 case：给出实际存在的目标页、具体替换值和可编辑 surface；然后保持这版 v2 不变再复跑。当前数据只支持：**v2 的结构更短，但短结构本身尚未带来可测的质量提升；工程硬门槛和任务契约比路由名称更先决定结果。**
