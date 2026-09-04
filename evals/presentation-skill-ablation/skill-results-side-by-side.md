# Presentation Skill 路线并列观察

生成于 2026-09-03。本文是一次冻结实验的可视化记录，不代表任何路线已经定型。图片目录中的 `current-production` 只是历史运行 ID，不代表生产状态。

## 先给结论

在目前**可比且通过作者侧与盲评侧硬门槛的 0→1 回合**中，`kimi-concise` 是最好的视觉基线：合格路线胜负为 Kimi 3、Current 1、Shared 1、Hybrid 1；加权绝对质量均值约为 Kimi 93.7、Shared 90.2、Hybrid 87.3、Current 85.0（满分 100）。

这足以支持一个有限判断：Kimi 式短路由加局部 `style brief`，更容易让模型快速形成明确的视觉意图，也更容易在不同场景间换构图，因此目前看起来最有可塑性。它**不能**证明 Kimi 在所有场景、尤其是 1→10 编辑上都更好；本轮 Hybrid 尚未完成 1→10，且所有配对置信区间都跨 0，样本只有 6 个可比 0→1 回合。

实际建议是：把 Kimi concise 当作 0→1 的视觉入口实验候选，把 Current 的证据边界、source-bound 保真、遮挡检查和渲染返工保留为所有路线的硬门槛；不要把目前的分数直接当成路线定型依据。

## 四条路线各自做了什么

| 路线 | 入口形态 | 强项 | 当前短板 |
| --- | --- | --- | --- |
| Current baseline | 当前较长 Skill 与分散 reference | 安全边界、PPJ/source-bound、复核纪律 | 上下文重，模型更容易落到保守、熟悉的构图 |
| Shared What / What-kind / How | 先写沟通任务，再读场景与 how 能力 | 分析页、技术页的关系表达较完整 | 规则层较多；品牌页出现硬门槛失败 |
| Kimi concise | 短主入口 → 场景指南 → 局部 style brief → Compose | 视觉层级、字体、图片和构图决策快，品牌/教育/管理页领先 | 1→10 保真还没有在本轮完成；短路由本身不会自动解决错误数据拓扑 |
| Hybrid short contract | Kimi 短路由 + 一句话 What/What-kind + Current 硬门槛 | 目标是兼顾创造力和安全性 | 当前只有 0→1；分析页把缺失中间观测连成 0，说明规则没有落到构图决策 |

“可塑性”在这里不是“风格更花”，而是模型能否在不依赖固定模板的情况下，根据页面职责重新组合图片、字体、几何、图表和负空间。Kimi 路线目前最像一个轻量的创作约束：约束足够少，能留下构图自由度；又有局部 style brief，避免每页从零解释审美。

## 0→1 单页并列

图片均为同一批冻结任务的 `slide-001.png`。四列只改变入口 Skill，PPJ、OfficeKit、素材规则、渲染和复核能力相同。

### analysis-decision-01

| Current baseline | Shared What / What-kind / How | Kimi concise | Hybrid |
| --- | --- | --- | --- |
| ![analysis current](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/analysis-decision-01/current-production/outputs/previews/slide-001.png) | ![analysis shared](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/analysis-decision-01/shared-what-kind-how/outputs/previews/slide-001.png) | ![analysis kimi](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/analysis-decision-01/kimi-concise/outputs/previews/slide-001.png) | ![analysis hybrid](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/hybrid-20260903/authors/analysis-decision-01/hybrid-short-contract/outputs/previews/slide-001.png) |

这里最值得注意的不是谁“更好看”，而是缺失数据的拓扑：Hybrid 的折线仍把 H2–H7 当成 0 连起来，盲评因此两轮都判为硬门槛失败。正确的可视化应使用 endpoint comparison 或明确断线，不能让短路由的优雅外观掩盖事实错误。

### management-report-01

| Current baseline | Shared What / What-kind / How | Kimi concise | Hybrid |
| --- | --- | --- | --- |
| ![management current](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/management-report-01/current-production/outputs/previews/slide-001.png) | ![management shared](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/management-report-01/shared-what-kind-how/outputs/previews/slide-001.png) | ![management kimi](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/management-report-01/kimi-concise/outputs/previews/slide-001.png) | ![management hybrid](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/hybrid-20260903/authors/management-report-01/hybrid-short-contract/outputs/previews/slide-001.png) |

管理汇报中，Kimi 的标题—证据—结论节奏更紧；Current 更像可靠的报告页，但画布占用和视觉焦点较保守。Shared 的 What 层帮助沟通任务，却没有始终转化为更强的视觉主次。

### technical-engineering-01

| Current baseline | Shared What / What-kind / How | Kimi concise | Hybrid |
| --- | --- | --- | --- |
| ![technical current](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/technical-engineering-01/current-production/outputs/previews/slide-001.png) | ![technical shared](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/technical-engineering-01/shared-what-kind-how/outputs/previews/slide-001.png) | ![technical kimi](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/technical-engineering-01/kimi-concise/outputs/previews/slide-001.png) | ![technical hybrid](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/hybrid-20260903/authors/technical-engineering-01/hybrid-short-contract/outputs/previews/slide-001.png) |

技术页需要关系、因果和节点证据。Shared 在这类页面的意图说明更完整；但 Hybrid 本轮出现硬门槛失败，说明“短入口 + 安全规则”仍需要一个把关系拓扑落实到 PPJ 元素的中间检查。

### academic-research-01

| Current baseline | Shared What / What-kind / How | Kimi concise | Hybrid |
| --- | --- | --- | --- |
| ![academic current](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/academic-research-01/current-production/outputs/previews/slide-001.png) | ![academic shared](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/academic-research-01/shared-what-kind-how/outputs/previews/slide-001.png) | ![academic kimi](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/academic-research-01/kimi-concise/outputs/previews/slide-001.png) | ![academic hybrid](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/hybrid-20260903/authors/academic-research-01/hybrid-short-contract/outputs/previews/slide-001.png) |

学术页的差异主要来自证据密度与阅读顺序。Kimi 更容易先做出一个主结论，再把方法或限制放到次级层；Current 的证据边界更稳，但若没有额外设计意图，页面会留下较多未承担职责的空间。

### education-training-01

| Current baseline | Shared What / What-kind / How | Kimi concise | Hybrid |
| --- | --- | --- | --- |
| ![education current](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/education-training-01/current-production/outputs/previews/slide-001.png) | ![education shared](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/education-training-01/shared-what-kind-how/outputs/previews/slide-001.png) | ![education kimi](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/education-training-01/kimi-concise/outputs/previews/slide-001.png) | ![education hybrid](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/hybrid-20260903/authors/education-training-01/hybrid-short-contract/outputs/previews/slide-001.png) |

教学页很依赖清晰的节奏与一个可记忆的视觉锚点。Kimi 的局部 style brief 通常能更快形成这个锚点；Shared 的结构解释更全，但容易把注意力分散在“应该包含什么”而不是“先让读者看到什么”。

### brand-creative-01

| Current baseline | Shared What / What-kind / How | Kimi concise | Hybrid |
| --- | --- | --- | --- |
| ![brand current](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/brand-creative-01/current-production/outputs/previews/slide-001.png) | ![brand shared](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/brand-creative-01/shared-what-kind-how/outputs/previews/slide-001.png) | ![brand kimi](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/tri-route-20260903/authors/brand-creative-01/kimi-concise/outputs/previews/slide-001.png) | ![brand hybrid](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/runs/hybrid-20260903/authors/brand-creative-01/hybrid-short-contract/outputs/previews/slide-001.png) |

品牌创意页最能看出“模板限制思路”与“短路由留出空间”的区别。Kimi 更容易采用全幅图像、强字级差、少量几何和明确的焦点；Shared 本轮盲评侧硬门槛失败，说明层级规则多不等于有设计感。Hybrid 的品牌页是本轮相对成功的混合样例：有氛围图、有可编辑的展示字和波形，但它不能抵消其他场景的事实与完成度问题。

## 分数和硬门槛怎么读

### 质量差分

只统计作者侧与盲评侧都通过的 6 个 0→1 回合；所有差分的 bootstrap 95% 区间都跨 0：

| 比较 | 平均差（左−右） | 中位数 | 解释 |
| --- | ---: | ---: | --- |
| Current − Shared | −5.17 | −8 | Shared 的视觉/沟通表现暂时占优，但不稳定 |
| Current − Kimi | −8.67 | −13.5 | Kimi 是当前最强视觉基线 |
| Current − Hybrid | −2.33 | −4 | Hybrid 可能改善 Current，但证据不足 |
| Shared − Kimi | −3.50 | −5 | Kimi 的构图与完成度暂时领先 Shared |
| Kimi − Hybrid | +6.33 | +8 | Hybrid 还没有达到 Kimi 的稳定程度 |

效率上，Kimi 的平均工具调用约 134 次，低于 Shared 约 147 次和 Current 约 151 次；这与“短路由更快进入创作”一致，但本轮 Hybrid 只有 6 个作者结果，不能直接做完整路线的成本比较。

### 不能被均分掩盖的失败

- Hybrid `analysis-decision-01` 两轮都把未知中间点画成连续的 0 趋势；这是事实拓扑错误，不是审美扣分。
- Hybrid `technical-engineering-01` 作者侧硬门槛失败；其余 Hybrid 1→10 结果大量缺失，因此尚未证明编辑可塑性。
- Shared `brand-creative-01` 两轮盲评硬门槛失败。
- 本轮没有 PowerPoint 播放证据；结论只覆盖 PPJ、结构、渲染和盲评观察。

## 四个图片字段的作用

这次新增的 `visualProfile` 是素材能力层，不是某个 Skill 的审美模板：

```json
{
  "alphaPresent": true,
  "subjectBounds": { "x": 0.12, "y": 0.08, "width": 0.76, "height": 0.84 },
  "edgeQuality": "soft",
  "shadowMode": "baked"
}
```

- `alphaPresent`：文件是否有已知透明通道；未知就保留 `null`。
- `subjectBounds`：在有限 alpha 检查中得到的非透明主体范围，不冒充语义识别。
- `edgeQuality`：`clean`、`soft`、`fringe` 或 `unknown`，用于判断扣图边缘是否需要谨慎。
- `shadowMode`：`none`、`baked`、`separate` 或 `unknown`；图片自带的阴影不能再叠一层原生阴影。

它们能让 Kimi 式视觉创作更稳定：先判断图片是不是适合“物品置于新背景”或“主体跨越底图”，再决定 crop、mask、蒙层和阴影，而不是把任意图片硬塞进一个精致布局。字段只描述可证明的图像事实，不证明法律许可、主体身份或审美质量。

## 下一步建议

1. 继续保留 Kimi concise 的短入口和局部 style brief，作为 0→1 的优先实验候选。
2. 将 Current 的证据与交付硬门槛抽成所有路线共享的不可绕过层，并增加“缺失数据不得连线”的 PPJ/Review 检查，而不是只写在长文档里。
3. 补完 Hybrid 的 1→10 配对实验，再判断它是否能把 Kimi 的视觉优势带到局部编辑。
4. 继续用 `visualProfile` 约束背景替换和扣图；无法证明 alpha、主体范围或边缘质量时，选择原生图形或明确标记人工复核。
5. 在路线切换前再做一轮小规模人类校准；当前差异方向很清楚，但统计证据还不足以声称普适胜负。

完整统计见 [report.multi-route.v1.md](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/report.multi-route.v1.md)，原始结构化结果见 [summary.multi-route.v1.json](/Users/zfang/workspace/officekit-presentation-skill-ablation/evals/presentation-skill-ablation/evidence/summary.multi-route.v1.json)。
