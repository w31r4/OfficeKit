## Why

OfficeKit 已有丰富的 PPJ、图片、图表、图层和 source-bound 能力，但 Skill 的信息组织仍缺少可验证的选择：一套是分层的 What / What-kind / How，另一套是短路由、场景指南和任务级 style brief。需要在相同运行时与输入上比较两种指导链路，区分“技能组织/注意力”与“原语能力”的影响，而不是凭单张 PPT 的印象决定默认路线。

## What Changes

- 增加一次性 presentation Skill 双路线实验，覆盖六个场景、0→1 创作和 1→10 连续编辑。
- 在实验专用目录提供 Shared What / What-kind / How 与 Kimi-style concise 两套 clean-room Skill overlay；两者共用 PPJ、原语、图片搜索、素材权利、渲染和复核不变量。
- 固化 12 个任务、复杂源稿、能力覆盖表、评分量表、模型运行参数和哈希；记录 24 个作者会话、24 个匿名盲评会话以及人类校准。
- 加入结构 oracle、视觉盲评、效率 ledger 和配对统计，分别报告质量、功能、展示、成本及未验证边界。
- 保留已有遮挡、稀疏数据、证据边界、渲染返工、PPJ/source-bound/opaque 等 Skill 原则；本 change 不选择默认路线，不修改 PPJ、编译器或公开 API。

## Capabilities

### New Capabilities

- `presentation-skill-ablation`: 冻结两套 Presentation Skill overlay、实验任务、运行证据、盲评和配对分析的研究契约。

### Modified Capabilities

无。实验使用现有 Presentations Skill 与 PPJ 能力，不改变其生产行为。

## Impact

- 新增 `evals/presentation-skill-ablation/` 的任务、两套 Skill、一次性 runner、证据 ledger 与报告。
- 新增 OpenSpec 设计、契约和任务清单。
- 实验运行会生成 PPTX、PPJ、渲染图和图片搜索日志；原始结果放在 worktree 外或被 Git 忽略，不进入 npm 包。
- 使用固定的 `gpt-5.6-luna`、`model_reasoning_effort=max` 和现有 OfficeKit PPJ CLI；不引入新的运行时依赖或 Office wire 版本。
