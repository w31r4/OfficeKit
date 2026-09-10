# PPJ × Kimi PPTD × 完整 PowerPoint 能力缺口待实现文档

> 状态：实现中（增量台账；未宣称完整 parity）
>
> 审计时间：2026-09-03
>
> 当前代码基线：`main@9b86939a`（本地 `main` 与 `origin/main` 一致）
>
> 范围：PPJ v1、PPTX NativeAOT codec、source-bound `nativeRef`、Presentation
> Skills。本文不包含 Windows PowerPoint 验收，也不把完整发布或 npm 发布作为本轮前置条件。

这份文档把两种不同的差距分开记录：

1. **相较 Kimi PPTD 的差距**：只比较 Kimi 的公开 `pptd.md` 语言能表达、能编译的语义。
2. **相较完整 PowerPoint 的差距**：比较 PowerPoint/OOXML 的完整对象、关系、宿主行为和编辑能力。

“有这个字段”不等于“能力已完成”。PPJ 的完成标准是：

```text
PPJ 表达
  → authored 编译
  → PPTX 导出
  → PPJ/原生投影导入
  → 受控编辑（如适用）
  → 二次导入恢复
  → 包范围、非目标内容和渲染证据
```

对第三方 PPTX，无法证明完整所有权的内容必须保留为 source-bound/opaque，或者在导出前明确失败；不能为了让编辑成功而拍平成图片或重写整个包。

## 1. 当前基线

### 1.1 语言与工程规模

| 项目 | 当前状态 | 说明 |
| --- | --- | --- |
| PPJ typed page elements | 已有 16 类 | `text`、`shape`、`line`、`icon`、`image`、`chart`、`table`、`connector`、`group`、`media`、`placeholder`、`smartArt`、`ole`、`opaque`、`component`、`slot` |
| PPJ 文档字段 | 920 个 root/definition fields | 由 schema/维护脚本重新计算；字段数不是完成度，其中很多是边界、投影或 authored-only 语义 |
| source-edit leaves | 122 个闭合叶子类型 | 当前闭合词汇包含图表类别/数值、系列样式、SmartArt 文本、几何/文本/图片、表格样式等边界；每个叶子仍须绑定 source hash、所有权和依赖关系 |
| Kimi PPTD 元素 | 7 类 | `text`、`shape`、`line`、`image`、`icon`、`table`、`chart` |
| Kimi PPTD 图表 | 13 类 series | `bar`、`line`、`area`、`scatter`、`bubble`、`candlestick`、`pie`、`radar`、`waterfall`、`heatmap`、`treemap`、`sunburst`、`sankey` |
| Kimi 动画 DSL | 未发现 | 当前检查的 `pptd.md` 没有 animation、transition、Morph、timing、trigger 字段 |
| PPJ 动画 | 部分完成 | 已有 transition、Morph、入口/退出/强调、点击顺序、段落构建、图表构建、delay/stagger |
| Presentation Template Library | 已完成基础库 | 39 个 schema v3 风格包，含 `SKILL.md`、预览图、校准图和可选 PPJ/PPTX 参考 |
| SmartArt | 部分完成（8 类布局 profile） | 支持 authored SmartArt、8 类布局、自定义定义资产、受控 source-bound 投影和显式 detach；OfficeKit 自有 `picture` 布局的已存在节点图片可按内容哈希投影为 PPJ asset，并在已证明的单一缓存图片图上投影/回写 `nodes[].image` 的 `fit`（stretch/tile）、crop、opacity 后受控替换 |
| Windows PowerPoint | 证据单列，不计入 gap | 当前证据是结构化导入/导出、Office Open XML 校验、LibreOffice/Keynote 或模型渲染；不声明 Windows 播放通过，也不因此扣减 PPJ 完成度 |

参考实现与覆盖记录：[PPJ language reference](../skills/presentations/skills/presentations/references/ppj.md)、[coverage](coverage.md)。Kimi 基准为相邻工作区的 `office-artifact-tool/kimi/skills/presentations/reference/pptd.md`，本次读取版本为 1,886 行，所在仓库为 `4bc98349b74c`，文件 SHA-256 为 `1711bbed8b4e62e47bd94ba8489a62e4e7d06fb6ebf24fb8e29ee73baf8ea0a3`。文件 hash 只用于固定本次对照快照，不代表 Kimi 的永久版本；相邻仓库的其他工作树改动不计入本对照。

### 1.2 状态标记

- **已完成**：PPJ 语义、编译器和必要的回读/恢复证据已存在；若只完成 authored，会明确写成“已完成（authored）”。
- **部分完成**：有确定的有限 profile，超出 profile 的内容会 source-preserved 或 fail closed。
- **只读/保留**：能识别、预览或保留原始内容，但没有安全的语义编辑路径。
- **未建模**：没有公共 PPJ 语义或没有可复用的编译路径。
- **宿主未验收**：结构可能已经正确，但尚未在目标 Office 宿主中验证实际播放或交互行为；这是独立证据状态，不是 PPJ 待实现 gap。

每个待办同时记录三条进度：

- **语言进度**：PPJ 是否能表达。
- **编解码进度**：是否能 authored 编译、导入、二次恢复，或 source-bound 编辑。
- **宿主进度（不计入 gap）**：是否有真实 PowerPoint 行为证据；本阶段默认不执行 Windows 验收，缺少该证据不降低 PPJ/codec 的语义完成度。

### 1.3 三维进度总表

下表是扫读时的权威状态。条目正文可以解释边界，但不得把下面的“部分”理解成整项能力已完成。

| ID | 差距类型 | 语言进度 | 编解码/回读进度 | 宿主证据（独立记录，不计入 PPJ gap/完成度） |
| --- | --- | --- | --- | --- |
| K-01 | Kimi 直接缺口 | 有界 profile 完成：`line` + typed path command + `points/viewBox/curve` | authored 编译→去嵌入 PPJ→再投影，以及 literal source-bound 单字段修改、changed-part 和二次投影回归均有证据；Kimi 点列已在 authored 边界降为 line/quadratic/cubic；复杂 geometry 仍按 F-04 管理 | 未验收 |
| K-02 | Kimi 直接缺口 | 有界 profile 完成：`chart.style.frame` | frame 的 solid/gradient/image fill、line/shadow 的 ChartSpace 读写、capability、frame-only source-bound 单字段 changed-part 与二次投影均有证据；image 素材替换已闭合 ChartPart `.rels`、新增/删除 media 与二次投影，共享/外部关系和完整 effect graph 仍按 F-07 管理 | 未验收 |
| K-03 | Kimi 直接缺口 | 已有 `dataset`/`encoding`/series-level `encode`/`dataFilter`/`seriesDefaults` schema | 宽表和 Kimi 风格 series encode 已归一化为 canonical categories/series；close/flow/isTotal、数值字符串、heatmap 矩阵、candlestick/sankey/waterfall 通道和 0/1 primary/secondary combo axis 映射有 parser/validator/ authored 证据；本轮新增本地工作簿公式引用的受控 source-bound ChartPart 编辑，并把安全识别出的点暴露为 opaque `nativeRef.leaves[].kind=chartDataCategory/chartDataValue/chartDataXValue/chartDataYValue/chartDataBubbleSize`，可由 PPJ 触发 ChartPart cache + embedded worksheet cell 的双 footprint 写回和二次投影；native point profile 已覆盖 bar/line/area/pie/doughnut/radar 六类单 family category plot、bounded column/line/area combo 的每条 `c:val/c:numRef` plot，以及 scatter/bubble 的 X/Y/size 数值通道，并由圆形/雷达/散点/气泡/组合 fixture 验证 cache-to-cell 绑定；本轮又为 category `c:cat/c:strRef` 增加直接 inline/string worksheet cell 的双 footprint 叶子与二次投影；直接 `c:strLit`/`c:numLit` literal cache 也已纳入 opaque native leaf 的 ChartPart-only 编辑，并由保留 `externalData` 关系的 fixture 证明不会误写 workbook；13 类 series 各自已有去嵌入 PPJ 的 native projection 证据（特殊族诚实降为 vector/group，waterfall 在内部降为 column）；bounded column/line/area 混合已有 authored/native 证据；source-bound common plot 标量现可对 legend、stacking、gapWidth、overlap、bar/column varyColors、轴/网格可见性、line smooth/varyColors 以及 bubble/circular 整数字段做 kind-checked `setChartPlot` 写回并二次投影；普通 x/y/secondary 轴的 `tickLabelInterval`、`min`、`max`、`majorUnit`、`position` (`bottom`/`top` for horizontal axes, `left`/`right` for vertical axes)、`visible`、`reverse`、`tickLabelsVisible`、`tickLabelPosition` (`nextTo`/`high`/`low`/`none`)、`numberFormat`、`axisLine`、`gridLine` 也可在 authored 和 source-bound 路径解析声明的 `size`/`boolean`/`string` token，并保持 ChartPart-only footprint；本轮又将现有 ChartML 直读/写能力接成独立 `setChartSeriesStyle`，line/scatter/radar 的 marker（含 symbol、size、RGB/alpha fill、marker stroke）以及非 scatter series 的 direct stroke 可在 source-bound PPJ 中增改/删除，只改目标 ChartPart 并由二次投影证明；更广泛跨 family 组合、共享/外链 workbook、全量 footprint 和多级/数组轴仍缺 | 未验收 |
| K-04 | 共同产品缺口 | 已有 `compositing` 受限语义（opacity/blend/clip/isolation） | shape/image/line/icon/placeholder 的 normal opacity authored lowering 已闭合；connector 现在也支持 `compositing.opacity`，通过单一 native line-alpha owner 乘入并规范投影为 `stroke.opacity`；本轮 `compositing.opacity` 和 image/line 根 opacity 也可引用 `opacity` kind 的 design token；本轮又闭合了 image 上单个、非 inverse、preset 或 bounded custom `compositing.clipStack` 到既有 picture mask owner 的 authored→DrawingML→去嵌入 PPJ→`image.mask` 回投影，preset adjustment 与 custom path command 都保留；多级/反向/超出 custom codec/非图片/与 `image.mask` 冲突的 clip、非 normal blend、isolation 明确 fail closed，native effect closure 尚缺 | 未验收 |
| K-05 | Kimi 直接表达力差距 | 已有 typed `tokens`、`stylePrecedence`、`predicates` schema | C# 校验、只读 grammar evaluator、grammar color token fallback 及 tint/shade authored lowering 已接入；本轮新增 `grammarTokenRef` 的 authored lowering：文字 `size`/`bold`/`italic`/`font`/`fontFamily`、图片及 image paint `fit`/`opacity`、形状透明度、solid fill opacity、image/line 根 opacity、stroke width、shadow opacity，以及 chart `titleTextStyle` 的 `fontSize`/`fontFamily`/`bold`/`italic`/`color` 会按声明 kind 解析，并由 authored→PPTX→再投影测试证明；同一 image-paint token 还可在 source-bound 图片元素和形状图片填充编辑中解析并回投影；source-bound chart common plot 标量、chart 标题/图例/数据标签/坐标轴文字样式（含 `fill/underline/alignment`）和固定拓扑表格 cell textStyle 均有 kind-checked resolver、ChartPart/SlidePart-only patch 和回投影证据；普通 x/y/secondary 轴的数值边界、方向、标签、`numberFormat` 和轴/网格线字段，以及 chart/radar data-label 的 `numberFormat`，同样接受 kind-checked `size`/`boolean`/`string` token，`setChartAxis`/`setChartLabels` 只写目标 ChartPart；`stylePrecedence` 首个来源现在也覆盖形状、图表和表格的有界字段（表头行数、交错行），并可对 chart `titleTextStyle`/`legendTextStyle`/`dataLabels.textStyle` 的 `fontSize`/`fontFamily`/`fontFamilyEastAsia`/`bold`/`italic`/`color` 按显式嵌套规则做有限浅合并；source-bound 实心 shape fill/stroke、有限 gradient stop、shape/image/chart-frame/table-cell shadow/border color、图片边框宽度和实心幻灯片背景的 RGB 颜色现在也可引用声明为 `color`/`size` kind 的 grammar token（保留有限 tint/shade/alpha 解析；未声明 token 的 stroke/shadow 仍按标准 DrawingML theme token 处理），并以 SlidePart/ChartPart-only 与二次投影回归证明；错误 kind 在颜色、stroke width、image border width、image fit 和 chart axes/number-format 位点均 fail closed；本轮又新增 `design.styles.image`/`image.styleRef`，对图片 paint 的 fit/crop/focus/opacity/border/shadow 按声明 precedence 做 authored/source-bound 生效值解析；完整样式写回、theme/master cascade 与更广泛 PPTX source-bound style closure 仍缺 | 未验收 |
| K-06 | Kimi 规范外的 PPJ 扩展 | 已有组件 `imagePolicy`（role/fit/mask/尺寸/rights） | slot 约束、asset metadata、authored 图片槽位编译/投影、schema-v3 `imageSlots` 示例绑定、metadata-only replacement plan、显式 elementId 驱动的纯 PPJ replacement transaction，以及 `applyTemplateImageReplacementToPptx` 的 source/asset hash→NativeAOT compile→changed-parts/output-hash→二次 projection 事务边界已有窄证据；真实 imported/source-bound fixture 的单 owner changed-part/二次 projection 联测已完成（确定性去嵌入 PPJ fixture）；真正跨导入焦点语义保留、共享/歧义关系治理和 source-owned accessibility 仍缺 | 未验收 |
| K-07 | PPJ 相对完整 PPT 的扩展 | 已有 timing graph sugar | 有限 graph 归一化到动画数组，compact `animations[]` 与 `timing.nodes[]` 都保留规范化 trigger 字段，`timeline` 继续以 `start` 为准；repeat/autoReverse/easing 和 trigger→start 归一化已闭合；完整 trigger closure、motion path、媒体 timing 尚缺 | 结构/Keynote 有证据；Windows 未验收（不计入 gap） |
| K-08 | 共同产品缺口 | 有界 review + authored `grid`/`flow`/weighted stack repeat | review 已加入越界、z-order、旋转矩形/阴影/箭头 visual bounds 和确定性文本溢出估算；组件 grid/flow repeat、anchor 和 horizontal/vertical `layout.weights` weighted stack 已能编译/投影，真实宿主测量、跨对象约束、solver 和 apply 操作尚缺 | 未验收 |
| F-01 | 完整 PowerPoint 差距 | 部分：source/provenance/PPJ 有 | 有界 closure edit 有，任意关系图未完成 | 未验收 |
| F-02 | 完整 PowerPoint 差距 | 部分：多 master + 有限 layouts | authored 已支持多个 master/layout，并新增 direct master/layout background、owner-local direct placeholder frame/text 的 hash-bound source-bound 回写；layout placeholder 若省略 direct `xfrm` 但可按所属 master 的同 type/index 唯一匹配，则 PPJ 投影其 effective frame，编辑后只在 layout owner 添加 direct `xfrm`；slide placeholder 若省略 direct `xfrm`，且能沿 slide→layout→master 以唯一 type/index 找到完整 frame，则 PPJ 也投影 effective frame，首次 `setFrame` 只在 slide owner 物化 direct `xfrm`；常见 picture/chart/table/date/footer/slide-number placeholder 现在可投影并在有 bounded capability 时编辑 frame/text，未知 `other` 仍只读；直接嵌入 master/layout 图片背景的 crop/opacity 与单 owner 资源替换（各自 `.rels`、新媒体写入和旧媒体清理）也有 source-bound 回写证据；去嵌入 PPJ 后恢复图数量；rotation/flip 的可选属性存在、显式零/false 与删除现在均有窄闭环，复杂 imported inheritance 未完成 | 未验收 |
| F-03 | 完整 PowerPoint 差距 | 有界：rich text/typed field/line-break/文本容器 bodyPr 子集 | authored/source-bound typed field（固定值与 bounded automatic `slidenum`/`author`/`datetime`/`datetimeFigureOut`/`datetime1`/`datetime2`/`datetime3`/`datetime4`/`datetime5`/`datetime6`/`datetime7`/`datetime8`/`datetime9`/`datetime10`/`datetime11`/`datetime12`/`datetime13`/`uaqdatetime1`/`uaqdatetime2`/`uaqdatetime3`/`uaqdatetime4`/`uaqdatetime5`/`uaqdatetime6`/`uaqdatetime7`）与有序 `run.break` 子集有；automatic `slidenum` 的缓存会按所属 slide 刷新，`author`、`datetime`、`datetimeFigureOut`、`uaqdatetime1`、`uaqdatetime2`、`uaqdatetime3`、`uaqdatetime4`、`uaqdatetime5`、`uaqdatetime6` 与 `uaqdatetime7` 保留缓存文本交给宿主刷新，`datetimeFigureOut` 请求宿主使用 MM/DD/YYYY 格式，`datetime1` 请求宿主使用 MM/DD/YYYY 格式，`datetime2` 请求宿主使用 Day, Month DD, YYYY 格式，`datetime3` 请求宿主使用 DD Month YYYY 格式，`datetime4` 请求宿主使用 Month DD, YYYY 格式，`datetime5` 请求宿主使用 DD-Mon-YY 格式，`datetime6` 请求宿主使用 Month YY 格式，`datetime7` 请求宿主使用 Mon-YY 格式，`datetime8` 请求宿主使用 MM/DD/YYYY hh:mm AM/PM 格式，`datetime9` 请求宿主使用 MM/DD/YYYY hh:mm:ss AM/PM 格式，`datetime10` 请求宿主使用 hh:mm 格式，`datetime11` 请求宿主使用 hh:mm:ss 格式，`datetime12` 请求宿主使用 hh:mm AM/PM 格式，`datetime13` 请求宿主使用 hh:mm:ss AM/PM 格式，`uaqdatetime1` 请求宿主使用 DD/MM/YYYY（Umm al‑Qura 日历），`uaqdatetime2` 请求宿主使用 Day, DD Month, YYYY（Umm al‑Qura 日历），`uaqdatetime3` 请求宿主使用 DD Month, YYYY（Umm al‑Qura 日历），`uaqdatetime4` 请求宿主使用 DD/MM/YY（Umm al‑Qura 日历），`uaqdatetime5` 请求宿主使用 YYYY-DD-MM（Umm al‑Qura 日历），`uaqdatetime6` 请求宿主使用 DD-Month-YY（Umm al‑Qura 日历），`uaqdatetime7` 请求宿主使用 DD Month YYYY（Umm al‑Qura 日历），静态 field display 的 source-bound 单字段编辑仍只改目标 SlidePart 并可二次投影恢复；line-break 只保持固定 inline topology；普通文本框、带文本形状和占位符现在可通过独立 `setTextBodyStyle` 回写直接 bodyPr 的 vertical alignment、wrap、inset、columns、column gap/direction、vertical text、rotation、horizontal/vertical overflow、upright、anchorCenter、有限 auto-fit 及 canonical `normalAutoFit` 百分比，并有 SlidePart-only/二次投影回归；普通 run、默认 run、表格和图表文本新增复杂脚本字体 `fontFamilyComplexScript`→`a:cs` 的 authored/imported/source-bound 单叶闭环；形式化 text owner 也增加 `text.language` 的 authored precedence 到直接 `a:rPr/@lang`；source-free authored text glow、inner shadow、reflection 与 soft edge 已分别写入直接 `a:effectLst`，并有 run/default-run XML、嵌入恢复和非法值拒绝证据；直接 rich-text run 的 source-bound glow 与 inner shadow 也已按严格拓扑投影，分别提供颜色、几何和显式 alpha native leaf 的 SlidePart-only 回写证据；paragraph `defaultText` 的 glow-radius/color/theme-color/opacity、shadow-blur/distance/direction/alignment/color/theme-color/opacity/rotate-with-shape、inner-shadow-blur/distance/direction/color/alpha、reflection-blur-distance-start-opacity-end-opacity-direction/soft-edge radius strict native leaves 也已有 focused source-bound 回写证据；未覆盖的其它 automatic field、完整 field/WordArt、继承、显式删除和复杂 bodyPr 未完成 | 未验收 |
| F-04 | 完整 PowerPoint 差距 | 部分：preset/custom geometry/connector/group transform | literal custom-path profile 有 authored/source-bound 编辑和单 SlidePart changed-part/二次投影证据；普通 group 的外层 `off/ext`、显式 `rot/flipH/flipV` 和局部 `chOff/chExt` 现在都可作为独立 native leaf 做 source-bound 单字段回写，PPJ 以 `childFrame` 保留子坐标矩形；对严格 image-fill shape 的 partial/formula custom geometry，独立 literal `val N` 调整 sibling 也可作为 native leaf 回写；完整 guide/handle 多路径拓扑、子空间联动和自动 descendant rescale 仍未完成 | 未验收 |
| F-05 | 完整 PowerPoint 差距 | 部分：图片/fill/crop/mask/effect 子集 | authored/source-bound 有界 profile 有；recognized picture 的 preset-mask identity（含默认 `rect` 的规范化叶子、完整 preset adjustments 的 `image.mask.preset`/`image.mask.adjustments` 变更），以及完整 `a:avLst/a:gd fmla="val N"` 的逐槽 `imageMaskAdjustment` native leaf 回写；当 mask 因 partial/formula adjustment 无法建模为 typed image 时，recognized preset + 简单直接 guide list 仍可为独立 literal `val N` sibling 颁发同一 native leaf，保留其余公式和 opaque 拓扑；border/shadow 通过 `setImageEffects`，普通形状/线条的 outer shadow 通过 `setShapeEffects`，picture outer shadow 的 source-bound `imageShadowRotateWithShape`/`imageShadowBlurRadiusEmu`/`imageShadowDistanceEmu`/`imageShadowDirectionDegrees`/`imageShadowAlignment`/`imageShadowOpacityThousandthPercent` 也可在严格直接 owner 内独立 token 回写，source-free authored 的 shape/image/line/picture `glow`、`innerShadow`、`reflection` 和 `softEdge` 已分别以有限 color/radius/opacity、color/blur/distance/angle/opacity、blur/startOpacity/endOpacity/distance/angle 或 radius 写入直接 `a:effectLst` owner，并保持 glow→inner shadow→outer shadow→reflection→soft edge 的独立顺序；source-bound 普通 shape/line/picture 的 direct `reflection` 也已按 full-span owner 投影，提供 `shapeReflection*`/`imageReflection*` 五类 native leaf、SlidePart-only token splice、outer-shadow 保留和复杂/非 full-span fail-closed 回归（`PpjSourceBoundShapeImageReflectionEditsOwnersAndReprojects`）；source-bound 普通 shape/line/picture 的 direct `softEdge` 也已按 direct、outer-shadow→soft-edge 或 full-span reflection→soft-edge owner 投影，提供 `shapeSoftEdgeRadiusEmu`/`imageSoftEdgeRadiusEmu` native leaf、SlidePart-only `rad` token splice、前置效果保留和复杂图 fail-closed 回归（`PpjSourceBoundShapeImageSoftEdgeEditsOwnersAndReprojects`）；source-bound 实心幻灯片背景的直接 RGB/`color` grammar token/opacity，以及直接嵌入图片背景的 bounded crop/opacity 和单 owner 资源替换通过既有 `p:bg`/`a:blipFill` 写回，slide、master、layout 的单 owner 关系/媒体闭包均有最小二次投影证据；OfficeKit 自有 picture SmartArt 的单一缓存图片也可投影/回写 `fit`（stretch/tile）、crop、opacity，但复杂遮罩、effect、共享/外链关系仍 source-owned；malformed/extension/child-bearing/unknown geometry 和非 literal 目标 guide 仍 source-owned | 未验收 |
| F-06 | 完整 PowerPoint 差距 | 有界：矩形表格、六个原生表格属性标志与 direct cell style 子集 | authored/source-bound 表格样式闭环已有；direct RGB/gradient/no-fill cell fill、单一直接嵌入 image fill（含关系/媒体闭包）、四边框、固定段落/run 拓扑的多 run 文本替换、跨段落样式一致多 run 文本样式，以及固定拓扑、纯文本、每 run 直接样式可表示的样式不一致多 run 文本 body 均已有 source-bound changed-part/二次投影证据；混合 run 通过结构化 `text.paragraphs[].runs[].style` 投影和 `setTableCellStyle/table.cell.textStyle` 回写，保持 run 数、段落数和未建模 XML；固定拓扑 mixed-run cell 还可在 `text.style` 暴露 bounded direct `a:bodyPr`（vertical alignment、wrap、四边 inset、columns/column gap/direction、vertical text、有限 auto-fit）并 source-bound 回写，只改所属 SlidePart；固定拓扑 cell textStyle 仍可解析 `size`/`bold`/`italic`/`fontFamily`/`fontFamilyEastAsia`/RGB(alpha) grammar token 并只改所属 SlidePart；同一 SlidePart 内被多个 cell 复用的 image relationship 现在采用 copy-on-write，替换目标新增关系并保留其他 cell 的旧媒体；固定拓扑单段落的嵌入式 picture bullet 现在也能以 PPJ `bullet: { type: "picture", asset }` 投影，保留并 source-bound 编辑其文字、直接 bullet font/color/size 样式，资产 ID 在 PPJ 与 native picture-bullet 命名空间之间按 hash 映射；现有 `table.style.headerRows`、`table.style.bandedRows`、`table.style.bandedColumns`、`table.style.firstColumnEmphasis` 与 `table.style.lastColumnEmphasis` 另有独立 `tableHeaderRows`/`tableBandedRows`/`tableBandedColumns`/`tableFirstColumnEmphasis`/`tableLastColumnEmphasis` native leaves，可分别只 token-splice 直接 `a:tblPr/@firstRow`/`@bandRow`/`@bandCol`/`@firstCol`/`@lastCol`；段落/列表/字段/高级效果、未建模 bodyPr/继承/reflow、cell inheritance、跨 owner/外部关系仍缺 | 未验收 |
| F-07 | 完整 PowerPoint 差距 | 部分：16 类 PPJ chart | native/vector 有界 profile 有；本轮新增安全本地公式引用的 ChartPart 读写，以及 opaque native chart 数据点的 bar/line/area/pie/doughnut/radar 和 bounded column/line/area combo ChartPart cache + embedded worksheet 双 footprint source-bound 写回；圆形/雷达/组合识别回归已补齐；直接 `c:strLit`/`c:numLit` literal cache 现在可作为 native leaf 做 ChartPart-only 增改，外部 `externalData` 不会被误写入 literal footprint；现有 ChartML 直读/写还通过独立 `setChartSeriesStyle` 闭合 line/scatter/radar marker（含 marker fill/stroke）与非 scatter series direct stroke 的增改删，保持 ChartPart-only footprint 和二次投影；柱/面积/组合/普通散点图的 `c:varyColors` 现在也进入 presence-aware `style.varyColors`，保留显式 true/false；本轮又补齐普通/combo ChartPart 的 `style.plotAreaLine`，以 `c:plotArea/c:spPr/a:ln` 承载直接 RGB、宽度和预置虚线，并与已有 plot-area fill 共存；又接入普通/combo ChartPart 的 `styleIndex`，以直接 `c:style/@val` 承载 1–48 的内建样式索引，source-bound 只改目标 ChartPart；完整 ChartML/workbook closure、外链/共享关系仍未完成 | 未验收 |
| F-08 | 完整 PowerPoint 差距 | 有界：8 类 SmartArt 布局 + picture 节点素材/缓存图片 paint 替换 | authored/source-bound 有界闭环有；OfficeKit 自有 picture 缓存图的节点 asset 替换，以及单一嵌入 blip 的 `fit`（stretch/tile）、crop、opacity 投影和 source-bound 回写，均具备关系/媒体闭包和二次投影证据；任意 DiagramML 未完成 | 未验收 |
| F-09 | 完整 PowerPoint 差距 | 部分：transition/animation/Morph 子集 | timing graph 有界闭环有；新增 source-bound 有限图编辑只改所属 SlidePart，并验证 repeat/autoReverse/easing/delay/duration 的二次投影；完整 timing 未完成 | 结构/Keynote 有证据；Windows 未验收（不计入 gap） |
| F-10 | 完整 PowerPoint 差距 | 基础/clone：media metadata | payload 编辑、播放控制未完成 | 未验收 |
| F-11 | 完整 PowerPoint 差距 | 少数 typed profile | OLE/3D/Ink/Custom XML/宏未完成 | 未验收 |
| F-12 | 完整 PowerPoint 差距 | 有界：notes/comments/sections/custom shows（legacy 与 modern root/direct-reply） | source-free modern author/person/anchor 已由 PPJ 确定性生成；master/handout、复杂 thread/action topology 未完成 | 未验收 |
| F-13 | 完整 PowerPoint 差距 | 部分：文本 hyperlink + typed click/hover shape action | URI、内部 slide、custom show、有限 action verb 已 authored/投影；安全形状的 click/hover source-bound 目标替换/移除已闭合并保留关系 changed-part 证据；trigger/声音/宏仍缺 | 未验收 |
| F-14 | 完整 PowerPoint 差距 | 部分：accessibility metadata + explicit reading order | authored/投影显式 reading order 和 PPJ machine review 已有；shape、image、chart、table、connector、group 的 canonical `accessibility` 现在颁发独立 `setAccessibility` capability，可在 source-bound PPJ 中增改/清除 title、description、decorative，并只改所属 SlidePart、保留图片残余扩展、二次投影恢复；安全 source-bound reading-order 通过 shape-tree z-order 回写并二次投影验证；group direct-child readingOrder 也已在有界 profile 中回写为本地 shape-tree 顺序；Checker 等价、SmartArt 内部/表格宿主语义仍缺 | 未验收 |
| F-15 | 完整 PowerPoint 差距 | 部分：theme color/font/style + bounded tint/shade/alpha/channel/discrete/hue transforms | grammar color token 的 tint/shade、authored `design.theme.accentColors` 六角色色板、`design.theme.accentTransforms` 六角色的直接 `a:tint`/`a:shade`/`a:lumMod`/`a:lumOff`/`a:alphaMod`/`a:alphaOff`/`a:satMod`/`a:satOff`/`a:redMod`/`a:redOff`/`a:greenMod`/`a:greenOff`/`a:blueMod`/`a:blueOff`/`a:hueMod`/`a:hueOff`/`a:gray`/`a:comp`/`a:inv`/`a:gamma`/`a:invGamma`、`design.theme.colorRoles` 六个 dark/light/hyperlink 角色（含有限 RGBA alpha）和 `design.theme.fontScheme.major/minor/majorEastAsia/minorEastAsia/majorComplexScript/minorComplexScript` 已有 bounded lowering；复杂脚本字体 direct `a:cs` owner 已有 authored/imported/source-bound 单叶闭环；完整 transform/effect/font scheme 未完成 | 未验收 |
| F-16 | 共同产品缺口 | 部分：只读 review + authored grid/flow/anchor/weighted stack repeat | bounds/z-order/保守 visual bounds/文本估算和有限 grid/flow/anchor/weighted stack repeat 已有；真实测量、跨对象约束、通用 solver/apply 未完成 | 未验收 |
| F-17 | 完整 PowerPoint 差距 | 未建模/只读为主 | 文档安全/签名/发布设置未完成 | 未验收 |
| F-18 | 宿主证据记录（不计入 PPJ gap） | 不属于 PPJ 字段 | 结构和模型渲染已有，Windows lane 未启动 | 明确未验收；仅记录证据状态 |

表中 K/F 行最后一列的“未验收”只表示 Windows PowerPoint 宿主证据尚未启动；它不参与语言进度、编解码进度、gap 数量或完成度分母。只有前两列描述的 PPJ 语义与编解码边界才进入本计划。

读表结论：相较 Kimi，最直接的结构性差距集中在 3 个原语（独立曲线、图表容器 frame、通用 dataset/encode），另有 1 个设计语法表达力差距；动画属于 PPJ 已领先、但还没有完整 PowerPoint 闭环的扩展项。相较完整 PowerPoint，17 个语义主项中没有任何一项可以标成“完整 parity”：14 项已有有界 profile，F-11 只有少数安全 profile，F-16 尚未交付布局求解，F-17 主要未建模/只读。F-18 只记录宿主证据，不属于 PPJ gap 或完成度分母。

### 1.4 K/F 主从关系

K 项用于说明 Kimi 语言对照；F 项用于说明完整 PowerPoint。实现时只建立一份主任务，避免重复统计：

| Kimi 对照项 | 完整 PPT 主项 | 关系 |
| --- | --- | --- |
| K-01 | F-04 | K-01 是独立 freeCurve 的语言差距，F-04 还包含完整 geometry/group/connector topology |
| K-02 | F-07 | K-02 是 chart container frame，F-07 还包含完整 ChartML |
| K-03 | F-07 | K-03 是 dataset/encode，F-07 还包含 workbook/formula/extension/3D |
| K-04 | F-05 | K-04 是 compositing 设计，F-05 还包含 PowerPoint effect graph |
| K-05 | F-15 | K-05 是可执行 design grammar，F-15 还包含 theme XML 和字体效果继承 |
| K-06 | F-05 / F-15 | image slot 属于模板系统；具体图片效果仍归 F-05，主题绑定归 F-15 |
| K-07 | F-09 | K-07 是 PPJ 扩展，F-09 是完整 timing/Morph/trigger 差距 |
| K-08 | F-16 | 两者都属于布局/审查能力，不是 Kimi `pptd.md` 已有的原语 |

K-04、K-08、F-16 是共同产品缺口；K-07 不是 Kimi parity 缺口。跟踪总量时以 F 项为完整能力主记录，K 项只记录对照和优先级。

### 1.5 证据索引与证据等级

状态表和本索引是本文件的状态权威；条目正文负责解释边界，不能单凭一个 schema 字段、一个 codec primitive 或一张渲染图把状态升级为“已完成”。证据等级只描述当前证明强度，不等于能力百分比：

| 等级 | 含义 |
| --- | --- |
| S0 | 只有 schema、文档、设计或对照规范；尚无可复现实现闭环 |
| S1 | 有 authored 编译、包结构或 NativeAOT/模型检查 |
| S2 | 有 authored 编译 → 导入 → 二次导入，或有受控 projection/reader 回读 |
| S3 | 有 source-bound 受控编辑、closure/所有权、changed-part/residual 证明和二次导入 |
| S4 | 有目标宿主打开、编辑、播放或交互证据；只作独立宿主证据记录，不进入本表 PPJ gap/完成度分母；Windows PowerPoint 未验收时不计入进度 |

| 范围 | 主证据入口 | 当前最高等级 | 证据边界 |
| --- | --- | --- | --- |
| PPJ schema 与元素 | [`src/ppj/ppj-v1.schema.json`](../src/ppj/ppj-v1.schema.json)、[PPJ reference](../skills/presentations/skills/presentations/references/ppj.md) | S2 | 字段和 authored profile 不自动证明第三方导入可编辑 |
| 总体覆盖 | [`docs/coverage.md`](coverage.md) | S2 | 是覆盖目录，不替代每个 primitive 的 round-trip 证据 |
| 图片、背景、paint、mask | [`openspec/changes/ppj-image-paint/`](../openspec/changes/ppj-image-paint/)、[`openspec/changes/ppj-custom-image-masks/`](../openspec/changes/ppj-custom-image-masks/) | S2–S3 | 只覆盖声明的 effect/closure profile |
| 图表 | [`openspec/changes/ppj-analytical-chart-primitives/`](../openspec/changes/ppj-analytical-chart-primitives/)、[`evals/presentation-six-sample-import/`](../evals/presentation-six-sample-import/) | S2–S3 | native/vector/opaque 分层，非完整 ChartML |
| 动画与 Morph | [`openspec/changes/presentation-motion-compiler/`](../openspec/changes/presentation-motion-compiler/) | S2 | 有结构/部分模型或 Keynote 观察；不是 Windows 播放证明 |
| SmartArt | [`openspec/changes/ppj-native-smartart-engine/`](../openspec/changes/ppj-native-smartart-engine/)、[`test/ppj-smartart-copy-on-write.mjs`](../test/ppj-smartart-copy-on-write.mjs)、`PptxCodecTests.SourceBoundPictureSmartArtCanReplaceAnExistingNodeAsset` | S2–S3 | 只覆盖 authored、已授权 source-bound text/graph/frame 和 OfficeKit 自有 picture 节点 asset profile |
| 模板库 | [`openspec/changes/presentation-template-skills/`](../openspec/changes/presentation-template-skills/)、[`docs/template-library-provenance.md`](template-library-provenance.md) | S1–S2 | 预览图是 guidance/evidence，不是可编辑页面本体 |
| Authoring / PPJ 编译 | [`openspec/changes/presentation-program-json/`](../openspec/changes/presentation-program-json/)、[`evals/presentation-program-json/`](../evals/presentation-program-json/)、[`evals/pptx-generation/`](../evals/pptx-generation/) | S1–S2 | 不能替代 source-bound 依赖闭包证明 |
| 本轮 PPJ gap profiles | `native/OfficeKit/tests/OfficeKit.Codec.Tests/PptxCodecTests.cs::PpjGapProfilesCompileAndReproject`、`PpjRadarSpokeGrammarTokensAuthorAndReproject`、`PpjKimiSmoothLineAuthorsMultiSegmentBezier`、`PpjKimiArbitraryMultiSegmentBezierStaysTypedPath`、`PpjStylePrecedenceAuthorsStyleRefBeforeInlineForShapeChartAndTable`、`PpjImageOpacityGrammarTokenAuthorsAndReprojects`、`PpjConnectorStrokeOpacityGrammarTokenAuthorsAndReprojects`、`PpjChartFrameImageFillAuthorsAndReprojects`、`PpjDatasetEncodingCoversKimiChannels`、`PpjDatasetEncodingAuthorsAllKimiSeriesFamilies`、`PpjFormulaChartReferenceProjectsAndEditsOnlyChartPart`、`PpjSourceBoundNativeChartDataLeafEditsCacheAndWorkbookAndReprojects`、`NativeChartDataLeavesCoverCircularAndRadarCategoryPlots`、`NativeChartDataLeavesCoverCategoricalComboPlots`、`NativeChartDataLeavesCoverScatterAndBubbleNumericChannels`、`SourceBoundLiteralCustomGeometryPathEditChangesOnlySlideAndReprojects`、`SourceBoundPartialCustomGeometryLiteralSiblingEditsAndReprojects`、`PpjTextFieldAuthorsAndProjectsAsTypedRun`、`PpjSourceBoundTextBodyStyleEditsTextShapeAndReprojects`、`PpjSourceBoundTableBandedRowsEditsLeafAndReprojects`、`PpjSourceBoundTableBandedColumnsEditsLeafAndReprojects`、`PpjSourceBoundTableHeaderRowsEditsLeafAndReprojects`、`PpjSourceBoundTableFirstColumnEmphasisEditsLeafAndReprojects`、`PpjSourceBoundTableLastColumnEmphasisEditsLeafAndReprojects`、`PpjReadingOrderAuthorsAndProjectsAsExplicitPermutation`、`PpjGroupReadingOrderAuthorsAndReordersLocalShapeTree`、`PpjSourceBoundAccessibilityEditsAndReprojects`、`PpjComponentImagePolicyAuthorsAndRejectsUnsafeReplacement`、`PpjComponentImageCropBindingAuthorsAndReprojects`、`PpjComponentGridRepeatAuthorsAndReprojectsDeterministicFrames`、`PpjSourceBoundTableCellStylesEditOnlySlideAndReproject`、`PpjSourceBoundTableCellMixedRunStylesEditOnlySlideAndReproject`、`PpjSourceBoundTableCellPictureBulletPreservesAssetAndEditsText`、`PpjSourceBoundTableCellFieldPreservesIdentityAndEditsCachedText`、`PpjSourceBoundTableCellLineBreakPreservesInlineAndEditsText`、`PpjSourceBoundTableCellImageFillReplacesRelationshipAndReprojects`、`PpjSourceBoundTableCellSharedImageFillPreservesOtherReference`、`PpjSourceBoundImageMaskPresetAndCustomIdentityChangesAndReprojects`；`src/ppj/review.mjs`/`test/review.mjs`；`docs/coverage.md` 的 PPJ gap-profile evidence | S3（公式引用、native chart data、literal custom geometry、partial custom-geometry literal sibling、static field display、generic text-container bodyPr style、table style `tableHeaderRows`/`tableBandedRows`/`tableBandedColumns`/`tableFirstColumnEmphasis`/`tableLastColumnEmphasis`、table field cached-text edit、table line-break topology/edit、shape click/hover action、page/group reading-order 窄路径、source-bound accessibility 六类 owner、chart-frame image relationship replacement、same-owner shared table-cell image copy-on-write、picture preset/custom-mask identity 与完整调整预设切换、mixed-run table text body）/S2（其余 authored/二次投影）+ review S1 | 覆盖 Kimi points line（含 5–128 点 smooth 多段 cubic lowering，且任意外部多段 cubic 保持 typed path）、chart frame（含 solid/gradient/image frame、显式 frame marker、image crop/opacity 和单 owner image replacement 的 ChartPart `.rels`/media closure）、dataset/encoding（13 类单 family authored、bounded column/line/area combo、0/1 secondary combo axis）、受限本地 `strRef/numRef` 公式引用、opaque native chart numeric leaves 的 ChartPart/cache + embedded worksheet 双 footprint（bar/line/area/pie/doughnut/radar category value、bounded column/line/area combo category value，以及 scatter/bubble X/Y/size 通道）、literal custom geometry path、partial custom-geometry literal sibling、静态 typed field display、普通文本框/带文本形状/占位符 direct bodyPr style、固定拓扑表格字段缓存文本与 `run.break`、grammar declarations/color tint/shade、有限 stylePrecedence（形状、图表和表格有界字段的首个来源）、image/line/stroke/shadow opacity token、imagePolicy 与重复组件 typed image.crop、timing sugar（repeat/autoReverse/easing/trigger）、shape click/hover action（含 source-bound URL 替换/删除和关系清理）、explicit page/group reading order（page 的 shape-tree z-order 与 group 的本地 child z-order 回写）、source-bound accessibility 六类 owner、normal compositing、component grid/flow/anchor、固定表格几何与 cell style/image replacement（含同一 SlidePart 多引用关系 copy-on-write）、recognized picture preset/custom-mask identity（含完整 adjustments）以及 chart/radar axis、data-label grammar token 的 authored/source-bound 回投影；只读 layout/accessibility report（含旋转/阴影 visual bounds 与文本估算）；更广泛混合 series/workbook、native effects、solver、media/macro action、SmartArt reading order 与宿主行为仍未证明 |
动画新增 `PpjSourceBoundTimingGraphEditsOnlySlideAndReprojects`：去嵌入 PPJ 后编辑有限 timing graph 的 duration/delay/repeat/autoReverse/easing，changed-part 仅为目标 SlidePart，Open XML 校验通过并可二次投影恢复。该结构化测试不等同于 Windows PowerPoint 播放验收。

点级 data-label 的三个新增窄回归 `PpjChartPointDataLabelFillAuthorAndEditSourceChart`、`PpjChartPointDataLabelLineAuthorAndEditSourceChart` 和 `PpjChartPointDataLabelTextAuthorAndEditSourceChart` 分别覆盖直接 fill、line 和单 run literal text 的 authored/source-bound ChartPart-only 回写与二次投影。

图表轴位置窄回归 `PpjChartAxisPositionsAuthorAndEditSourceChart` 覆盖 `c:axPos` 的水平/垂直位置 authoring、去嵌入投影、默认位置 source-bound 回写和 ChartPart-only 二次投影。

`test/review.mjs` 新增 `line` 箭头头型 review smoke：按端点方向和 stroke width 计算 `none|triangle|stealth|diamond|oval|open` 的确定性保守 visual bounds，并同时验证箭头越界、遮挡检测标记为 `arrow-visual-bounds` 及稳定人工建议；这是只读几何代理，不是 Windows/PowerPoint 精确轮廓验收。

本轮新增 `PpjConnectorCompositingOpacityAuthorsAndReprojects`：connector 的
`compositing.opacity` grammar token 与既有 `stroke.opacity` 相乘，去嵌入 PPJ
后以 `stroke.opacity` 规范投影并保留 `connector` 类型。

| Source-bound 六样本 | [`evals/presentation-six-sample-import/`](../evals/presentation-six-sample-import/)、相关 capability/codec 测试 | S3 | 样本是边界证据，不代表任意第三方包 |
| Authoring compiler 质量 | [`evals/presentation-authoring-compiler/`](../evals/presentation-authoring-compiler/) | S1–S2 | 评估结果需与本文件的字段状态分别记录 |
| Kimi PPTD 对照 | `../office-artifact-tool/kimi/skills/presentations/reference/pptd.md`；仓库 `4bc98349b74c`；SHA-256 `1711bbed8b4e62e47bd94ba8489a62e4e7d06fb6ebf24fb8e29ee73baf8ea0a3` | S0 | 规范对照快照，不是 Kimi 的实现或宿主验收 |

上表最后一行的 hash 用于固定本次对照快照。宿主证据单独记录在 F-18，不将 LibreOffice、Poppler、Keynote 或模型渲染提升为 Windows PowerPoint 证据。

本轮补充的背景证据：`PptxCodecTests::SourceBoundImageBackgroundCropAndOpacityEditOnlySlideAndReprojects` 覆盖直接嵌入图片背景的 crop、`fit`/`opacity` grammar token 投影、仅所属 `SlidePart` 的 source-bound 回写、Open XML 校验和二次投影恢复；`PptxCodecTests::SourceBoundImageBackgroundReplacementClosesRelationshipsAndReprojects` 覆盖单 owner slide 背景替换及关系/媒体闭包；`PpjPresentationTests::PpjSourceBoundMasterAndLayoutImageBackgroundCropOpacityEditAndReprojects` 覆盖 master/layout 各自 owner-local 的同类回写；新增 `PpjPresentationTests::PpjSourceBoundMasterAndLayoutImageBackgroundReplacementClosesRelationshipsAndReprojects` 证明 master/layout 也可各自替换图片并清理旧关系/媒体，再经 Open XML 校验和二次投影恢复。该证据不扩展到外部/跨 owner 共享图片关系或 effect/clip 图。

本轮新增 SmartArt 图片证据：`PptxCodecTests::SourceBoundPictureSmartArtCanReplaceAnExistingNodeAsset` 以 OfficeKit 自有 `picture` 布局 SmartArt 为边界，先从四个标准 diagram parts、cached drawing 和其 media 闭包投影节点 asset，再替换一个已有节点的图片。编译只允许该 SmartArt 的 SlidePart、四个 diagram parts、cached drawing、缓存 drawing 的嵌套 `.rels` 和媒体闭包变化；新增关系源路径按实际 `DiagramPersistLayoutPart` 登记，避免把嵌套图片关系错误归到 slide。Open XML 校验、源保护和二次投影均通过；添加/删除节点 asset、第三方/共享/外部关系和未验证缓存仍 fail-closed。

本轮 SmartArt paint 增量补充：上述 `SourceBoundPictureSmartArtCanReplaceAnExistingNodeAsset` 还覆盖 canonical `nodes[].image`（`fit: stretch|tile`、crop、opacity）的投影和同一 source-bound 事务中的回写；复杂 blipFill、添加/删除图片和共享/外链关系仍 fail-closed。

## 2. 相较 Kimi PPTD 的差距

Kimi PPTD 是一个偏 authoring 的 YAML 多文件 DSL，不是完整 PPTX 的读写规范，也没有为任意第三方 PPTX 提供 source-bound 编辑保证。因此这里的“追平”只表示 PPJ 能表达 Kimi 公开规范中的语义，不表示两边拥有相同的内部文件拓扑。

### K-01 独立自由曲线原语 `linePath` / `freeCurve`

**优先级：P0；状态：有界 profile 完成，完整 PowerPoint geometry 仍归 F-04。**

Kimi 有独立的 `line` 元素，包含 `viewBox`、任意点列、sharp/round/smooth 曲线、起止箭头、旋转、翻转、透明度、边框和阴影。PPJ 原有的 `connector` 是端点驱动的 `straight`、`elbow`、`curved`，端点可以绑定元素或绝对坐标；本轮已补上独立 `line` 的 literal path 和有限 Kimi 点列 lowering，但这仍不是完整 DrawingML geometry 的替代品。

**当前进度：** `line` 已进入 PPJ schema/wire，支持结构化 `moveTo/lineTo/quadraticTo/cubicTo/arcTo` path command、stroke、箭头、透明度和阴影的 authored profile；NativeAOT 已有 literal-path reader/writer，导入后可投影为独立 `line`。本轮又接入 Kimi 风格的 `points` 字符串、二元 `viewBox` 和 `curve: sharp|round|smooth`：sharp/round 降为折线，smooth 在 2/3/4 点时分别降为线段/二次/三次 Bézier，5–128 点把首尾作为端点、其余点作为同一个高阶 Bézier 的控制点，并按等参数区间确定性降为多段 cubic Bézier；越界、非法数值和超过当前有界控制点数的图形 fail closed。完整 authored → PPTX → 去嵌入 PPJ → 再投影，以及 source-bound 单字段曲线命令修改仍有证据：只改变目标 `ppt/slides/slide*.xml`，保留其余包内容，并由二次投影恢复新命令值。高层 `points` 是 authored 语法，导入无私有 PPJ 时会在可稳定反解的有界路径上保留 compact points，否则诚实正规化为 typed path；复杂 geometry 的全量语义仍归 F-04，不把本 profile 夸大为完整 DrawingML geometry。

**仍缺：**

- 本轮新增窄闭环：`lineStartArrowWidth`、`lineStartArrowLength`、`lineEndArrowWidth`、`lineEndArrowLength` 分别绑定 shape/connector 直接 `a:headEnd/@w|@len` 与 `a:tailEnd/@w|@len`；只有已有 `sm|med|lg` token、单一 direct line profile 才发 native leaf，编辑只替换对应 endpoint 属性，缺失、unknown、extension、none 或歧义保持 source-owned。
- Kimi 点列的高层 `curve` 标记现在有四条可证明窄映射：literal `moveTo + lineTo` 路径按 direct line join 回读为 `points/viewBox/curve: sharp|round`，单段 quadratic/cubic 回读为 `points/viewBox/curve: smooth`，而 authored 的 5–128 点高阶 Bézier lowering 会按同一确定性分段与控制点公式回投影为 compact `points/viewBox/curve: smooth`（仅在最多 48 个控制点且量化后的 native path 可稳定反解时保留）；任意不符合该公式的多段 Bézier、arc/reference 路径仍正规化为 typed path。source-bound 的 compact points/curve 修改会重写同一条已授权 path，并在没有显式 stroke join 时把 `round`/`sharp` 映射到 native round/miter join；不支持的曲率图仍 fail closed。
- 完整 DrawingML guide/formula/handle/connection-site 图；
- 超过 48 个控制点的 compact points 回读，以及量化误差导致无法稳定反解的外部高阶 Bézier；这类路径仍回读为 typed path；
- 多路径、闭合填充路径以及复杂 group transform 的完整语义；
- 与 F-04 合并后的任意第三方 geometry/source-bound 编辑。

**待实现：**

1. 保持独立 `line`/`freeCurve` 元素；不改变现有 `connector` 的端点拓扑语义。
2. 已将 Kimi `points/viewBox/curve` 作为 authored 语法降为结构化 path command，而不是把 SVG 字符串直接暴露为执行入口；对无引用的纯直线 native path，snapshot-free projection 会保留紧凑 `points/viewBox/curve: sharp|round` 拼法；对单段 quadratic/cubic 和 5–128 点、且控制点符合 PPJ 高阶 Bézier 分段公式的多段 cubic，在最多 48 个点且有界反解稳定时也会保留 `curve: smooth` 及点列。compact `points/viewBox/curve` 的 source-bound 修改只重写已授权 path，复杂路径继续使用 typed path 并保持 fail closed。
3. authored 路径写入可编辑 DrawingML；仅对无 guide、无 handle、无未知扩展的 literal native path 提供 source-bound 投影。
4. 复杂路径保持 opaque，不自动转成 connector；只对 literal、无未知子树的路径颁发 source-bound 能力。

**验收：**

- authored：Kimi `points` 的 sharp/round 折线、smooth 的 2/3/4 点线段/二次/三次曲线及 5–128 点高阶 Bézier 多段 cubic，加上显式 path 的圆弧、箭头、旋转/翻转、透明度和阴影编译与二次导入；不符合高阶 Bézier 分段公式的多段 cubic 必须保留 typed path，不得误识别成 compact sugar；
- source-bound：一个 literal path 的单字段修改只改变目标 slide shape tree，其他 part 和关系保持不变，二次投影恢复新命令；
- 任何越界/非法点列、当前边界外的 smooth 控制点、formula-backed path、未知扩展、guide/handle 图都明确拒绝；
- 不使用 PNG 或整页图片替代曲线语义。

### K-02 图表容器级 `chartFrame`

**优先级：P0；状态：有界 profile 完成，完整 ChartML effect 仍归 F-07。**

Kimi 的 `Chart` 直接拥有 `fill`、`border`、`shadow`，这些属性作用于整个图表矩形容器。PPJ 当前图表有 `chartAreaFill`、`plotAreaFill`、标题、图例、轴和数据标签样式，但没有与 Kimi 一一对应的图表整体边框和阴影字段。

**当前进度：** `chart.style.frame` 已进入 schema/wire；ChartSpace 直接 `c:spPr` 的 solid/gradient fill、line、outer shadow 已有受限 reader/writer，projection 会颁发独立 `setChartFrame`。Presentation 路径现在允许 frame 装饰与 fill 共存，同时保持 XLSX 旧 fill-only 严格边界；没有 frame 的旧程序仍走 `chartAreaFill` 兼容路径。本轮用无图例/无数据标签的最小图表完成 gradient frame authored→再投影，以及 frame-only source-bound 单字段线宽修改：只改变目标 `ppt/slides/charts/chart*.xml`，二次投影恢复新宽度，且请求未带入自由曲线等其他修改。

**仍缺：**

- 共享或外部 image relationship 的跨 owner 治理，以及更完整的整体 fill/border/shadow effect graph（direct ChartSpace 的 solid/gradient/image 和单 owner 素材替换已在当前有界 profile 内支持）；
- 容器效果与 plot area、series paint 的边界；
- 导入后只改容器属性、不改图表数据/轴/series 的更广泛 source-bound operation；当前已证明的 changed-part 是目标 ChartPart，而非外围 shape。

**待实现：**

1. 在 `chart.style` 下增加明确的 `frame`（或等价命名），区分 `frame`、`chartArea`、`plotArea` 和 series paint。
2. authored 输出使用图表 graphic frame 的直接 shape properties；不能把外围 shape 当作同一个图表对象。
3. source-bound 只接受唯一、无未知扩展的 graphic-frame profile；为 `setChartFrame` 颁发独立 capability。

**验收：**

- solid/gradient/image frame fill、四边线、阴影的 authored round-trip；image frame 已证明可经 ChartPart relationship 写入并二次投影，source-bound 还覆盖单 owner 素材替换：更新 ChartPart `.rels`、新增目标媒体、清理不再引用的旧关系/媒体，并以二次投影恢复 crop、stretch/tile、opacity；
- 图表数据、标题、图例、轴和 series XML 在 frame-only 编辑中保持不变；focused test 证明只改目标 ChartPart；
- 不支持的 native effect graph 保留原文并撤回 capability；
- 图表整体旋转/翻转仍按 PowerPoint 限制处理，不伪造全局 opacity。

### K-03 通用图表数据集与 `encode`

**优先级：P0；状态：部分完成。**

Kimi 的图表顶层使用一个 `ChartData { cols, rows }`，每个 series 用 `encode` 指定 x/y/category/value/open/high/low/close/size/source/target/flow，还支持 `dataFilter`、`seriesDefaults` 和数组形式的多坐标轴。PPJ 当前以 `categories + series[].values` 为主，并通过 `xValues`、`bubbleSizes`、`parents`、`sources`、`targets` 等专用数组表达语义。

PPJ 的固定数组对常见图表更安全，也更容易做 source-bound 依赖证明；但它不能直接复用一个宽表数据集，也不能表达 Kimi 的通用字段映射和每个 series 的筛选器。

**当前进度：** `chart.data.dataset`、顶层 `encoding`、series-level `encode`、`dataFilter` 和按类型的 `seriesDefaults` 已进入 schema/parser；编译器将宽表和 Kimi 风格的每个 series 映射正规化为 canonical `categories + series`。当前已经覆盖 `x/y/category/value`、`open/high/low/close`、`size`、`source/target/flow`、`parent`、`level`、`isTotal`，支持 null/missing、有限字面值筛选、数值字符串转换和 heatmap 的 x×y 矩阵展开；series 的 `xAxisIndex/yAxisIndex` 已支持有限 0/1 映射到 primary/secondary combo 轴组，并贯穿 canonical model 与 projector 回投影为可再次消费的双索引；Kimi 风格的一项/两项 `xAxis/yAxis` 数组会正规化为相应轴对，超过两级或在非 combo 中使用 secondary 会 fail closed。对无公式、无稀疏点覆盖、无高级 series 样式的原生分类图，以及同样满足条件的 scatter/bubble 数值图，projection 现在在保留旧 `categories + series`（数值图保留 `xValues + values (+ bubbleSizes)`）字段的同时增加可复用的 canonical `dataset + encoding` 视图；复杂公式、点级拓扑或高级样式仍只输出旧字段，避免 canonicalize 时丢失语义。`PpjDatasetEncodingCoversKimiChannels` 覆盖代表性通道、secondary-axis combo 和双轴数组；`PpjDatasetEncodingAuthorsAllKimiSeriesFamilies` 已用参数化测试完成 Kimi 13 类 series 各自的 PPJ 校验、authored 编译和去嵌入 PPJ 后的 native projection：bar/line/area/scatter/bubble/pie/radar 保持 native ChartPart，waterfall 在内部以四个 series 降为 column，candlestick/heatmap/treemap/sunburst/sankey 诚实降为可编辑 group/vector（candlestick 同时覆盖 line overlay，hierarchy `levels` 进入 dataset-series schema）。`seriesDefaults` 现在对嵌套对象执行递归合并：series 可局部覆盖 `marker`、`dataLabels` 及其 `textStyle` 等子字段，同时保留未覆盖的默认值，并把合并结果继承到 canonical series；`PpjDatasetEncodingAuthorsAllKimiSeriesFamilies` 已覆盖 nested data-label text-style 的继承/覆盖。ChartML `strRef/numRef` 的安全本地公式引用仍由 `PpjFormulaChartReferenceProjectsAndEditsOnlyChartPart` 覆盖：投影保留公式字段，公式变更只改所属 ChartPart，缓存不随意单独改写；对无法进入高层 chart 的 native chart，若唯一内部工作簿、单 worksheet、bar/line/area/pie/doughnut/radar 的 `c:val/c:numRef`，或 bounded column/line/area combo 每条 plot 的 `c:val/c:numRef`，或 scatter/bubble 的 `c:xVal/c:numRef`、`c:yVal/c:numRef`、`c:bubbleSize/c:numRef`，缓存与数字单元格逐点一致，则由 `nativeRef.leaves[]` 按 `chartDataValue`、`chartDataXValue`、`chartDataYValue`、`chartDataBubbleSize` 暴露独立叶子，PPJ source-bound 编辑同时改 ChartPart cache 和对应 worksheet cell，`PpjSourceBoundNativeChartDataLeafEditsCacheAndWorkbookAndReprojects` 已证明 changed-part/二次投影，`NativeChartDataLeavesCoverCircularAndRadarCategoryPlots`、`NativeChartDataLeavesCoverCategoricalComboPlots` 与 `NativeChartDataLeavesCoverScatterAndBubbleNumericChannels` 补足圆形/雷达、组合及数值图三通道识别回归。已有 bounded categorical `column/line/area` 混合图和 candlestick overlay；现有 ChartML reader/writer 还通过 `setChartSeriesStyle` 开放 line/scatter/radar marker 的 symbol、size、直接 RGB/alpha fill、marker stroke，以及非 scatter series direct stroke 的增改删；该路径只写目标 ChartPart，`PpjGapProfilesCompileAndReproject` 已证明 combo line series 的样式修改和删除会在二次投影恢复。仍未完成的是更广泛的跨 family 组合、轴数组对象原始形状/超过两级坐标系的无损恢复、共享/外链 workbook 以及完整 ChartPart/嵌入 workbook closure。

**本轮点级数据标签填充增量（2026-09-07）：** Kimi 的 `dataLabelOverrides[].fill` 现在补齐为 PPJ `series[].dataLabels.points[].fill`；普通 native/combo ChartPart 的稀疏 `c:dLbl` 以直接 `c:spPr` 承载 bounded none、solid RGB 或 literal gradient paint，authored 与 source-bound 回投影恢复点级填充。source-bound 只改目标 ChartPart；point label line、图片/主题/效果图和复杂 label graph 继续 fail-closed。`PpjChartPointDataLabelFillAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮点级数据标签轮廓增量（2026-09-07）：** Kimi 的 `dataLabelOverrides[].line` 现在补齐为 PPJ `series[].dataLabels.points[].line`；普通 native/combo ChartPart 的稀疏 `c:dLbl` 以直接 `c:spPr/a:ln` 承载 bounded RGB、宽度和预置虚线，并可与点级 fill 共存。source-bound 只改目标 ChartPart；自定义 line/effect graph、leader-line 几何和 vector fallback 继续 fail-closed。`PpjChartPointDataLabelLineAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮点级数据标签文本增量（2026-09-07）：** Kimi 的 `dataLabelOverrides[].text` 现在补齐为 PPJ `series[].dataLabels.points[].text`；普通 native/combo ChartPart 的稀疏 `c:dLbl` 以单段单 run 的 literal `c:tx/c:rich` 承载文本，并可与点级 fill/line 共存。source-bound 只改目标 ChartPart；公式引用、多段/多 run rich text、布局/效果图和 vector fallback 继续 fail-closed。`PpjChartPointDataLabelTextAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**仍缺：**

- source-bound 时可定位、可同步修改的 ChartPart/cache/workbook 二维 footprint（安全数字点通道已有独立双 footprint 叶子；series marker/stroke 则是 ChartPart-only 样式闭包）；
- 除 bounded `column/line/area` 与 candlestick overlay 外的全 13 类 Kimi series 跨 family 混合约束和各自完整样式/坐标规则；单 family authored 编译与 PPJ 再投影已有参数化证据；
- 超过 primary/secondary 两级的 axis index、Kimi 轴数组的完整原始形状和任意混合坐标系；当前只保证一项/两项数组映射到 PPJ combo 的 primary/secondary 轴组，并在 series 上回投影 0/1 索引，不保留超过两级或数组与独立 secondary 字段的原始写法；
- 固定数组与表格编码之间的无损恢复规则（安全原生分类图和简单 scatter/bubble 现在附带 canonical dataset，但仍不承诺保留输入宽表的列名/行对象形状；带公式、点级拓扑或高级样式的图继续使用旧字段）；
- ChartML 公式引用与 embedded workbook 的完整同步闭包（当前高层公式仍只保留/编辑本地引用字符串，不求值；opaque native chart 只对已证明的单 worksheet 数值缓存开放同步，单 family category、bounded column/line/area combo 的 `c:val` 以及 scatter/bubble 的 X/Y/size 已按独立通道开放）；外链工作簿、错误引用、共享缓存和 irregular ChartML 的安全边界。

**待实现：**

1. `chart.dataset`/`chart.encoding` 和 series-level `encode` 已作为兼容层落地，旧的 `categories + series` 继续有效；安全原生分类图及简单 scatter/bubble 的 projection 会同时提供 canonical dataset 视图，复杂图仍保守留在旧字段。
2. 将 Kimi 的 encode 归一化为 PPJ 内部 canonical series，不让编译器在运行时解释任意表达式；numeric channel 只接受有限数字或可解析的 invariant 字符串。
3. 对 `dataFilter` 只接受有限的列、字面值和比较形式；不引入 JS、XPath 或网络查询；二维 heatmap 缺行显式保留为 null。
   直接 `c:strLit`/`c:numLit` literal cache 现在也可在 opaque native chart 中作为独立叶子编辑；该路径只改所属 ChartPart，不把保留的 `externalData` 关系误当成 workbook 依赖。
4. 当前已闭合两条不同的 workbook/ChartPart 窄路径，另加一条 literal ChartPart-only 路径：高层 chart 的本地公式变更只改 ChartPart；opaque native chart 的安全 category/value 数点会通过 `nativeRef.leaves[].kind=chartDataCategory/chartDataValue/chartDataXValue/chartDataYValue/chartDataBubbleSize` 按通道同步改 ChartPart cache 与唯一 embedded worksheet cell。后者仍不求值：`chartDataCategory` 只接受单 worksheet 的 `c:cat/c:strRef` 与直接 inline/string cell；数值点只接受单 worksheet、单 family category 或 bounded column/line/area combo 的 `c:val/c:numRef`，以及 scatter/bubble 的 `c:xVal/c:numRef`、`c:yVal/c:numRef`、`c:bubbleSize/c:numRef`，每条缓存与对应单元格都必须逐点一致；各通道独立证明，共享字符串、外链、公式、错误引用和不规则拓扑继续 fail closed。

**验收：**

- Kimi 13 类 series 的单表表达、null/missing、单 family authored 闭环、bounded `column/line/area` 混合、candlestick overlay 和有限 0/1 axis-array index；当前已闭合 13 类单 family 的编译/再投影，其他跨 family 混合约束和完整样式/坐标规则仍待补齐；
- 旧 PPJ 程序导入后可以恢复为旧字段或 canonical dataset；
- authored 编译后二次导入恢复 encode/dataFilter 语义；
- source-bound 的无公式数据只改声明的数据 footprint；高层本地公式引用仍只允许改 ChartPart 且保留原 cache；opaque native chart 的 `chartDataCategory`/`chartDataValue`/`chartDataXValue`/`chartDataYValue`/`chartDataBubbleSize` 窄路径才允许同时改 ChartPart 对应 cache 与 embedded worksheet cell，并由 PPJ native leaf 测试证明 category `c:strRef`（直接 inline/string cell）、六类单 family category plot、bounded column/line/area combo 以及 scatter/bubble 的 X/Y/size 三通道双 footprint；其他 workbook 拓扑仍不开放，不能把这条窄路径冒充通用 workbook 同步；
- 不把第三方任意 shape 网格反推成 chart dataset。

### K-04 图层合成：统一透明度、混合模式、裁剪栈和隔离

**优先级：P0；状态：部分完成，且部分属于 Kimi 与 PPJ 的共同缺口。**

**本轮增量：** 对 source-bound 普通非 placeholder、无可见文字且所有已建模填充/渐变/图片填充/描边/阴影 alpha 一致的 shape，新增 `compositing.opacity` 复合 owner；精确 `setOpacity/compositing.opacity` 回写把绝对 alpha 同步到各 native paint owner，并只改目标 SlidePart。混合 alpha、文字、line、custom/opaque image-fill topology 继续不暴露该 capability。

PPJ 已有图片 `fit/crop/opacity/mask`、fill/stroke/shadow alpha、shape opacity、渐变、image fill 和 custom-path image mask。Kimi 有 `crop`、`fit`、`cropShape`、`opacity`，但当前 `pptd.md` 没有 `blendMode`、`compositeOperation` 或 `isolation` 字段。

**当前进度：** 已增加受限 `compositing` 声明和诊断；shape/image/line/icon/placeholder 的 normal opacity 可 authored 编译，connector 的 `compositing.opacity` 也已通过既有 native line-alpha owner authored lowering，并在去嵌入 PPJ 后规范投影为 `stroke.opacity`；`compositing.opacity` 可使用声明为 `opacity` kind 的 grammar token。页面背景的 source-bound solid fill 现在接受直接 RGB 或声明为 `color` kind 的 grammar token（含既有 tint/shade/alpha 解析），也接受直接 opacity 或同类 grammar token，只改所属 `SlidePart` 的 `p:bg` 并通过二次投影恢复。本轮又把 image 上单个、非 inverse、preset 或 bounded custom `compositing.clipStack` 绑定到既有 picture mask owner，包含 preset adjustments 和 custom path commands；去掉嵌入 PPJ 后分别由原生 mask 规范投影为 `image.mask`。多级/反向/超出 custom codec/非图片/与 `image.mask` 冲突的 clip、非 normal blend 和 isolation 仍明确 fail closed。native/source-bound effect closure、固定合成顺序和 lossy render-as-image 分级仍未完成。

**仍缺：**

本轮后续增量（2026-09-05）：`lineOpacityThousandthPercent` 为 shape/connector 的 direct `a:ln/a:solidFill/{a:srgbClr|a:schemeClr}/a:alpha/@val` 提供独立 native leaf；只有单一 direct outline、单一 RGB/theme solid paint 和已有 `0..100000` alpha token 才发 leaf，source-bound 编辑只替换 alpha token，并验证 SlidePart-only changed-part 与二次投影；缺失、compound、继承、扩展或歧义 line graph 保持 source-owned。

- `normal/multiply/screen/overlay` 等可声明的混合模式；
- group/page/layer 级的透明度合成顺序；
- 多级 clip/mask 栈，以及反向、超出 custom codec、非图片或与已有 `image.mask` 组合的 clip 语义；
- mask 坐标系、fit、crop、透明度和阴影的固定渲染顺序；
- 对不支持 native blend 的情况进行显式能力分级，而不是隐式拍平。

**待实现：**

1. 定义 `compositing` 语义对象，至少区分 element opacity、paint opacity、clip/mask 和 blend mode。
2. 将可由 DrawingML 原生表达的组合落成可编辑对象；无法由 PowerPoint 原生表达的组合必须输出 structured unsupported 或显式 lossy `renderAsImage` 结果，不能偷偷替换。
3. 对 group/page 合成采用固定 z-order 和 alpha 规则，避免导入后重新排序。
4. source-bound 只对完整、唯一、无扩展的 effect/clip closure 发 capability；其余保留原 XML。

**验收：**

- authored：图片、形状、文字、group 的 crop/mask/opacity 顺序可重现；
- native 支持的组合完成二次导入和编辑；
- native 不支持的 blend 明确报告原因，不生成“看起来对但不可编辑”的假语义；
- 未改变 source-bound 非目标 effects、关系和媒体字节。

### K-05 结构化设计 token、样式优先级和继承

**优先级：P1；状态：部分完成。**

**本轮补充：** 普通 x/y/secondary 轴标题、轴以及 chart/radar data-label 的 `numberFormat`，plot/series/point data-label 的可见性和位置，以及 radar spoke label 的字号/字体/粗斜体，现在可使用声明为 `string`/`boolean`/`size` 的 grammar token；authored 与 source-bound 路径都会做 kind 检查，后者仅修改目标 ChartPart 并通过二次投影恢复。

**本轮图表文字增量（2026-09-07）：** Kimi `ChartTextStyleConfig.underline` 已接入 PPJ `chartTextStyle.underline`，并映射为直接 DrawingML `a:rPr/@u` 或 `a:defRPr/@u`；公开的 `single`/`double` 会分别规范化为原生 `sng`/`dbl`，其余值限定在受控 underline token 集合内。标题、图例、数据标签以及同一 chart text-style codec 覆盖的坐标轴文字 owner 已有 authored/source-bound 编解码路径；focused round-trip 已验证标题/图例/数据标签的 ChartPart-only 改写和二次投影。`uFill`、`uLn` 等带效果的下划线图以及未知 rich-text 拓扑继续保留为 source-owned 并 fail closed。

**本轮图表文字对齐增量（2026-09-07）：** Kimi `ChartTextStyleConfig.alignment` 已接入 PPJ `chartTextStyle.alignment`，并映射为直接 DrawingML `a:pPr/@algn`；公开的 `left`/`center`/`right`/`justify` 会分别规范化为原生 `l`/`ctr`/`r`/`just`。标题、图例、数据标签、轴标题和刻度标签的段落 owner 已有 authored/source-bound 编解码路径；focused round-trip 已验证标题/图例/数据标签的 ChartPart-only 改写和二次投影。未知 rich-text 段落拓扑和 vector fallback 继续保留为 source-owned 并 fail closed。

**本轮图表文字填充增量（2026-09-07）：** Kimi `ChartTextStyleConfig.fill` 已接入 PPJ `chartTextStyle.fill`；普通 native/combo ChartPart 的标题、图例、数据标签、轴标题和刻度标签文字 owner 支持 bounded `none`、solid RGB 和 literal gradient paint，落到直接 DrawingML `a:noFill`/`a:solidFill`/`a:gradFill`。solid 继续规范化为既有 `color` 表示，none/gradient 保留 `fill` 形状；authored 与 source-bound 回投影均只改目标 ChartPart。图片、主题/效果图和未知 rich-text 拓扑继续 fail-closed。`PpjChartTextFillAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**图片背景补充：** 直接嵌入的 slide image paint 也走同一 `grammarTokenRef` 解析路径；source-bound crop 回写同时可将 `fit`/`opacity` 解析为声明的 `string`/`opacity` kind，并保持 `SlidePart`/关系媒体闭包及二次投影恢复。当前证明仍限于直接嵌入、单 owner 图片，不涵盖外部/共享关系或 effect/clip 图。

**本轮图片样式增量：** `design.styles.image` 提供可复用的有界图片 paint style，图片元素可用 `styleRef`，并以 inline `style` 做字段级覆盖。`image.fit`、`crop`、`focus`、`opacity`、`border`、`shadow` 均按声明的 `stylePrecedence` 首个来源解析；authored 编译和 source-bound 图片编辑复用既有 fit/crop/opacity/effects capability，不改变图片关系或媒体所有权。source-bound 只比较生效值，若仅替换 style 引用但生效值不变则保持物理 no-op；复杂 mask、主题级 cascade 和跨对象样式仍 source-owned。

PPJ 已有 theme、fonts、named styles 和 `designGrammar`。但 `surfaceHierarchy`、`typographyRhythm`、`imageRules`、`chartRules` 等 grammar 字段目前是受长度约束的字符串列表；它们是指导和审阅依据，不是可以独立执行的约束 AST。Kimi `pptd.md` 对 text/table/chart 的 style priority 和默认值链写得更明确。

**当前进度：** `designGrammar.tokens`、`stylePrecedence`、`predicates` 已有 typed schema 和 C# 语义校验（类型、有限值、唯一性、比较操作）。`reviewPpjArtifact` 有确定性的只读 grammar evaluator：按声明顺序计算 inline/styleRef/theme/master/default 的首个命中，并将命中的 `{token}` 解析为实际值、保留 token/kind 证据，再逐元素评估有限比较谓词；predicate 失败只报告 warning，不把审美判断伪装成编译错误。本轮又把 `grammarTokenRef` 接入 authored compiler 的受限字段：文字 `size`/`bold`/`italic`/`font`/`fontFamily`、图片元素和 image paint `fit`/`opacity`、形状透明度、solid fill opacity、image/line 根 opacity、stroke width、shadow opacity、表格布尔样式标志，以及 chart `titleTextStyle` 的 `fontSize`/`fontFamily`/`bold`/`italic`/`color` 会先校验 token kind，再写入 NativeAOT 原生样式；图表常用的 `legend`/`stacking`/`bubbleSizeMode` 枚举、`gapWidth`/切片角度/孔径/气泡缩放整数、轴/标签可见性和 line `smooth`/`varyColors` 也支持 `string`/`size`/`boolean` token，并在解析后重新检查有限词汇和整数边界；普通 x/y/secondary 轴的 `title`、`tickLabelInterval`、`min`、`max`、`majorUnit`、`visible`、`reverse`、`tickLabelsVisible`、`axisLine`、`gridLine`，plot/series/point data-label 的显示标志、位置、`numberFormat`，以及 radar spoke label 的字号、字体和粗斜体现在也按相同规则解析，true grid-line 使用 DrawingML 的默认可见规范化。并将 `stylePrecedence` 的首个命中接入文字 run 的 `size`/`bold`/`italic`/`font`/`fontFamily`、形状 `fill/stroke/shadow/opacity`、图表 `legend/chartAreaFill/plotAreaFill/frame/titleTextStyle` 和表格 `headerRows/bandedRows` 的 authored 解析。图表文字样式保持旧的整对象来源选择；只有显式声明 `chart.*TextStyle.<field>` 嵌套规则时，才按字段执行有限浅合并，以便在不改变旧程序结果的前提下表达 Kimi 风格的部分覆盖。测试 `PpjStylePrecedenceAuthorsStyleRefBeforeInlineForShapeChartAndTable`、`PpjGapProfilesCompileAndReproject`、`PpjRadarSpokeGrammarTokensAuthorAndReproject` 和 `PpjImageOpacityGrammarTokenAuthorsAndReprojects` 覆盖 authored→PPTX→去嵌入 PPJ→再投影的具体值恢复，验证 styleRef 优先于 inline、嵌套字段可从 inline 补齐、图片元素与 shape/text/chart/table image paint 的 fit/opacity、图表 token 的枚举/整数/布尔字段、表格布尔 token 生效，并验证颜色、stroke width 和 chart axis/label 位点引用错误 kind 会 fail closed；其中 image-paint 的 fit/opacity、solid fill opacity、chart/table direct solid fill opacity，以及 shape/line/connector stroke width/opacity token 还通过 source-bound 受控编辑与二次投影；本轮又把 source-bound 图表标题/图例/数据标签/坐标轴文字样式、雷达 spoke label 文字样式、图表 frame 的 gradient stop 颜色，以及数据标签显示标志/位置和轴标题的解析接入同一 kind-checked resolver，并由 `PpjGapProfilesCompileAndReproject` 与 `PpjRadarSpokeGrammarTokensAuthorAndReproject` 证明目标 ChartPart-only 修改和二次投影恢复。样式写回、完整 master/theme 继承和更广泛 source-bound style closure 仍未完成。

**本轮正式宿主增量：** 对六个文字标量 `text.size`/`bold`/`italic`/`font`/`fontFamily`/`fontFamilyEastAsia`，新增 authored-only 的有限宿主链：`run` → `paragraph` → `element` → `styleRef` → `layout` → `master` → `theme` → `default`。`textStyle`、段落 `style.defaultText`、named style、layout/master 直连 `style` 都可提供对应值；编译器按首个命中值写入 native text run，`reviewPpjArtifact` 按每个 rich-text run 报告相同 winner/source，文本容器的 `style` 仍只负责 body/layout 属性。源绑定 PPJ 若携带这些新 owner 或 formal source 会明确 fail closed。回归 `PpjFormalTextStylePrecedenceAuthorsAndReprojectsEffectiveRunSizes` 覆盖 run、paragraph、element、named style、layout、master、default 以及 body-style fallback。

**Theme owner 边界：** `design.theme.textStyle` 现在是这八个 formal 文字标量的直接 authored fallback owner；theme 命中、缺字段回落到 default token 和 source-bound 拒绝该字段都有最小 round-trip/review 证据。它不等于可编辑的完整 `theme1.xml` font/effect scheme，完整 theme/master cascade 仍是残差。

**语言 token 增量：** `text.language` 现在也可引用声明为 `string` 的 grammar token；解析后仍必须通过有界 BCP-47 校验，再写入直接 `a:rPr/@lang`。错误 token kind 或非法 tag fail closed，不做 locale 推断、字体回退或 source-bound theme 写回。

**仍缺：**

- chart/table/image 的完整跨对象 cascade、原生 master/theme 样式图和更广泛 source-bound style closure；本轮只闭合了六个文字标量的 authored formal owner 链，且必须由声明的 `stylePrecedence` 规则启用；
- 更多可验证的 design token 类型和字段（当前 authored profile 已扩展到图表常用枚举/整数/布尔字段，但仍不是所有规则都用 prose 之外的可执行 AST）；
- chart/table/image 风格的完整同一继承机制（图片 paint style 的 fit/crop/focus/opacity/border/shadow 已有有界 `design.styles.image` profile，主题/master cascade、mask/effect 与跨 owner 关系仍缺）；
- “规则违反”与“编译失败”的边界，以及把解析结果应用到 native style 的写回路径。

**待实现：**

1. 已增加有限的 `design.tokens`、`stylePrecedence` 和 typed rule predicates，并在 review 中实现只读求值；`grammarTokenRef` 的文字/形状/image/line stroke/shadow/chart 标题基本属性 authored lowering 已落地，原有字符串 grammar 保持兼容。图表标题、图例和数据标签文字样式支持显式嵌套 precedence 的有限字段级浅合并；未声明嵌套规则时仍保持整对象的既有来源优先级。
   图片还可通过 `design.styles.image` + `image.styleRef` 复用 fit/crop/focus/opacity/border/shadow；source-bound 只对生效值触发现有局部 capability，避免把纯引用变化误写成图片关系变化。
2. 当前 evaluator 和 authored compiler 对六个文字标量已采用 formal precedence 的“首个声明来源优先”规则，并将最终值降低到 native run；下一步才把同等完整继承扩展到 chart/table/image 以及原生 master/theme 样式图。
3. 将能验证的规则转成 diagnostics；审美判断和叙事判断继续保留为 warning，不伪装成形式证明。

**验收：**

- 同一 PPJ 在不同 review/编译运行中 token 与 precedence 解析稳定；
- 直接字段覆盖继承字段的规则有机器测试；
- style token 变更不会意外改动 source-bound 原生对象；
- 无法解析的主题/effect graph 仍保留并标记边界。

### K-06 模板图片示例与可复用图片槽位

**优先级：P1；状态：部分完成（基础模板库与单 owner source-bound 事务已交付，跨导入语义治理仍缺）。**

当前 Presentation Template Library 已有 39 个风格包、预览图和校准图；模板主要作为风格指导和视觉校准，明确不包含可执行页面代码或固定版式骨架。

**当前进度：** component `slotDefinition.imagePolicy` 已支持 role、fit/mask 白名单、最小尺寸和 rights，并在 PPJ 校验中检查 asset metadata；组件参数现在还可以用 typed `crop` 值绑定到模板内的 `image.crop`，或用 typed `focus` 值绑定 `fit: "cover"` 的 `image.focus`，由 authored 编译推导非对称 crop，去嵌入再投影会恢复可执行裁剪边界。模板搜索的 schema-v3 sidecar 现在还支持可选的 `imageSlots`：每个槽位绑定一个已声明的 calibration example，并明确 role、allowedFit、allowedMask、最小像素尺寸和 rights；Evidence Ledger 模板已有两个通过 hash/路径交叉校验的槽位，搜索结果会返回这些边界及已校验示例的绝对路径/哈希。Template Creator 的包装 spec 也接受绝对 `examplePath` 并在生成 sidecar 时改写为包内路径；选定槽位和内容寻址图片元数据可生成 `office-kit/template-image-slot/v1` replacement plan，检查 fit/mask/尺寸/rights，并默认保留既有 crop/focus/accessibility。本轮新增 `applyTemplateImageReplacement`：调用方必须提供唯一 `elementId`，函数深拷贝 PPJ、校验 plan/policy 和内容寻址 asset 声明，追加新 asset 或验证已有 asset，按显式 override 应用 fit 与 `none|rect|roundRect|ellipse` mask，并保持未覆盖的 crop/focus/border/shadow/accessibility；source-bound 图片会先检查 `replaceImage`、`setImageFit`、`setImageMask` 及调整/路径所需字段，不满足就 fail closed。新增 `applyTemplateImageReplacementToPptx` 作为显式 native 事务：校验 exact source/asset bytes，动态调用 PPJ NativeAOT compiler，要求 output hash 与 changed-part receipt，再把输出送回 PPJ projector；source-free fixture 已在本机 darwin-arm64 NativeAOT 上完成真实 compile→project→目标元素语义确认；随后又在去嵌入 PPJ 的 imported/source-bound fixture 上完成同一 plan→NativeAOT compile→changed-parts/output-hash→二次 projection，证明单 owner 替换保持 source-bound、目标 SlidePart/关系/媒体闭包可审计且目标元素指向新媒体；模板搜索保持 metadata-only，native 依赖不在模块加载时初始化。事务有输入不变、no-op、能力缺失、custom-mask、asset hash 拒绝和 compile/project 编排回归，避免按 role 或像素猜测槽位。

**仍缺：**

- 让全部模板逐步声明图片槽位，并把槽位角色与示例图的语义绑定扩展到 `hero/background/avatar/chart-source` 等更多场景（当前只有 Evidence Ledger 作为窄证据）；
- 在真实 imported/source-bound PPTX 上继续扩大 replacement plan 的关系治理边界；当前已有确定性去嵌入 PPJ fixture 的 changed-part/二次投影联测，但仍不直接按角色扫描，也不自动处理共享、歧义或外部关系；
- 图片槽位的裁剪、mask、焦点和版权/替换约束（当前已完成显式 crop/focus binding 的 authored lowering，不做自动 saliency 推导）；
- 由示例图反推出可验证风格 token 的路径；
- 同一模板在新内容密度下的安全变体。

**待实现：**

1. 已在 schema-v3 metadata 中增加 typed `imageSlots`、适用角色、可用 fit/mask、最小尺寸、rights 和已声明示例路径；Template Creator 会做绝对示例路径到包内路径的安全改写；PPJ component 仍负责 authored 时的 crop/focus/capability 约束，搜索模块提供显式 replacement plan，`applyTemplateImageReplacement` 将该 plan 绑定到唯一 PPJ image owner，`applyTemplateImageReplacementToPptx` 再以 exact source/asset hash 调用 NativeAOT 编译并二次投影，返回可审计的 changed-parts/output-hash 收据。
2. 预览/校准图继续作为 evidence，不把图片像素当成可编辑模板本体。
3. 生成 PPJ 时把槽位解析为正式 asset 引用和显式 frame，不复制外部源 deck 的隐含关系。
4. 记录模板指导、实际选择和最终 PPJ 的差异，避免把“参考图存在”误报为“模板可复用”。

**验收：**

- 搜索结果能返回角色和适用边界，而不是只返回一张 preview（Evidence Ledger 的两个槽位已有机器回归）；
- 选定槽位后可生成并应用 replacement plan 替换 asset，保留 crop/mask/fit 和 accessibility/rights；当前已验证 plan 的尺寸、rights、fit/mask 拒绝路径，以及纯事务的输入不变、no-op、asset 声明和 source-bound capability 约束；`applyTemplateImageReplacementToPptx` 已把 plan 接到 NativeAOT compile/project 收据并覆盖 hash/编排拒绝，source-free fixture 和确定性 imported/source-bound fixture 均已完成真实 NativeAOT compile→changed-part→二次投影语义确认；共享、歧义、外部关系以及 source-owned accessibility 仍需保持 fail closed；
- 验收时必须通过 no-op、二次导入、素材 hash 和渲染检查；验收前不得把该槽位标为已完成；
- 模板仍不携带可执行页面脚本或未声明的源 PPTX。

### K-07 高级动画与渐进披露

**优先级：P1；状态：PPJ 已领先于 Kimi，但相对完整 PowerPoint 仍部分完成。**

PPJ 已有页面 transition、Morph、入口/退出/强调效果、`withPrevious`/`afterPrevious`/`onClick`、delay/stagger、段落构建和图表构建。Kimi `pptd.md` 没有对应 DSL，因此这不是 Kimi parity 的缺口，而是 PPJ 的继续拓展项。

**当前进度：** 页面 `timing/timingGraph` 已作为 typed sugar 接入 parser/validator，并正规化到既有动画数组；compact `animations[]` 与 `timing.nodes[]` 都保留规范化 `trigger` 字段，`timeline` 继续以 `start` 为准。当前有限 profile 支持 linear/ease-in/ease-out/ease-in-out、repeat 1–8、autoReverse，以及与 `start` 一致的 `onClick`/`afterPrevious`/`withPrevious` trigger。NativeAOT 会把这些字段写入 `p:cTn`，再投影回 PPJ；trigger 不一致、未知 timing 或超出预算会 fail closed。新增 `PpjSourceBoundTimingGraphEditsOnlySlideAndReprojects` 已在去嵌入 PPJ 后修改 duration/delay/repeat/autoReverse/easing，并证明只改目标 `ppt/slides/slide1.xml`、Open XML 有效且二次投影恢复；该测试是有限 timing graph 的 source-bound 证据，不是宿主播放证明。

**仍缺：**

- motion path、keyframe 和更丰富的 easing/keyframe 曲线；
- 形状触发器、书签/超链接触发器和条件序列；
- media timing、音效、trim、播放状态；
- 更完整的 chart build、SmartArt build 和跨对象依赖；
- “渐进披露状态机”而不仅是有序动画数组。

**待实现：**

1. 扩展为 typed timing graph，保留当前简化数组作为 sugar。
2. 在已有有限 profile 之上补 trigger closure、sequence/parallel group、motion path 和媒体 timing；repeat/autoReverse/easing 已有 bounded wire/profile，不把它们误报为完整宿主动画。
3. 对未知 imported timing 继续 opaque-preserved；不能把不完整 timing graph 猜成简单 fade。
4. 将 playback evidence 与结构 round-trip 分开记录；本阶段不启用 Windows 验收。

**验收：**

- authored timing graph 可稳定导出并二次恢复；
- 删除或修改一个动画只影响拥有该 timing closure 的 part；
- 触发器、媒体和未知扩展没有 capability 时 fail closed；
- review 报告清楚区分结构存在、宿主识别和真实播放。

### K-08 自动排版、遮挡检测和文本溢出

**优先级：P1；状态：部分完成（只读 review 已交付）；Kimi PPTD 也没有完整对应 DSL。**

PPJ 已有 frame、master/layout、placeholder、component repeat/when、text wrap/AutoFit 和 review warning。导入画布调整也明确不自动 scale、reflow、crop 或移动元素。Kimi 主要是每个元素的 `bounds`，没有 constraint/flow/grid/occlusion solver。

**当前进度：** `reviewPpjArtifact` 已输出 bounds、越界、z-order 和遮挡报告，并抑制父子容器重复告警；本轮加入旋转矩形的轴对齐可见包围盒、direct/style/chart-frame shadow 的距离/方向/blur 保守扩张、line/connector 的有限箭头头型轮廓代理，以及基于字体大小、Unicode 字符宽度、边距和换行的确定性文本溢出估算。箭头代理按端点方向、stroke width 和 `none|triangle|stealth|diamond|oval|open` 头型生成确定性的保守包围盒，并在报告中标记 `arrow-visual-bounds`，不声称等价于 PowerPoint 的精确轮廓。组件 repeat 另有 authored `grid` 和有限 `flow`（显式或稳定近方形列数换行）lowering，支持 `anchor: start|center|end` 把矩阵放置在实例剩余空间中，也支持 horizontal/vertical `layout.weights` 按 item-count-matched 正权重分配 stack slot；结果始终是普通 PPJ frame。文本报告明确标记 `deterministic-character-metric`，并把有 AutoFit 的对象与无 AutoFit 的 warning 分开；review 仍是只读，不会移动或重排对象，也不把估算当成 PowerPoint 实际排版。

**仍缺：**

- stack 的完整约束、跨对象约束和 solver 语义；当前只覆盖 authored component repeat 的有限 horizontal/vertical `layout.weights` 槽位分配，以及 flow/anchor 的有限矩阵放置；
- 文本实际测量、换行、AutoFit 后的碰撞检测；
- 图片 mask 和真实宿主文本测量；阴影/旋转/有限箭头头型已纳入保守可见边界，文本目前只有确定性估算，真实 mask 轮廓仍未纳入；
- z-order 遮挡、不可见内容和安全边距诊断；
- 自动执行修复；review 现在只给确定性建议，不会替 source-bound 对象移动或重排。

**待实现：**

1. 只读的 `layout-review-v1` 已完成首个可见边界 profile：计算几何交集、边界越界、阴影/旋转/有限箭头头型包围盒、遮挡关系和确定性文本溢出估算，并为越界/重叠给出稳定的人工修复建议；下一步补真实字体/宿主测量和 mask 轮廓。
2. 已在 authored component `grid` 之外增加有限 `flow`（省略列数时使用稳定近方形列数）、`anchor`（剩余空间的 start/center/end 对齐）和 `layout.weights` weighted stack（horizontal/vertical 方向、正权重与 item 数严格匹配、gap 保留）；所有结果序列化为普通 PPJ frame。完整 stack 约束、跨对象约束和 solver 仍保持未实现。
3. 对 source-bound 页面默认只给出报告和建议，不自动移动或重排原对象。
4. 任何自动修复都必须成为显式操作，并保留原始 frame、变更清单和可回滚证据。

**验收：**

- 同一输入得到相同的 layout report 和修复建议；
- 报告包含对象 ID、原始 frame、可见边界、遮挡对象、检测方式、严重级别和建议动作；阴影/旋转边界使用确定性的保守计算；
- authored 修复后无预定义碰撞和文本溢出；
- imported source 不因 review 被静默改写。

## 3. 相较完整 PowerPoint 的待实现内容

完整 PowerPoint 不是一个固定的单一 DSL：它包括 PresentationML、DrawingML、ChartML、DiagramML、媒体关系、嵌入 Office 文件、主题继承、扩展命名空间、宿主计算和 UI 行为。下面的清单以“有界可编辑 PPJ”作为目标，不承诺重造 PowerPoint 的全部实现。

### F-01 OPC 包、关系和源保真

**优先级：持续基础能力；状态：部分完成。**

**当前进度：** 无操作时可以保持源字节；已有 source hash、part/relationship inventory、typed edit plan、非目标 part residual 和若干图片、表格、图表、SmartArt、OLE 的 source-bound 操作。

**仍缺：** 任意关系拓扑编辑、共享关系重连、外部链接重写、任意扩展 part 修改、宏/自定义 XML/未知 content type 的通用语义编辑。

**待实现与验收：** 每种新操作都必须声明 closure、所有权和 changed-part footprint；建立 canonical OPC hash；二次导入后稳定恢复；对 shared/external/ambiguous closure 继续保留或失败，不做全包重写。

### F-02 Presentation、Slide、Master、Layout 和 Theme

**优先级：P0；状态：部分完成。**

**当前进度：** authored 支持多个 canonical master、`blank/title/titleOnly/obj` layouts、页面布局引用、直接背景、标题/正文/副标题占位符；每个 layout 绑定到所属 master，去嵌入 PPJ 后仍可恢复多个 master/layout/page 图。source-bound PPJ 现在为可安全编辑的 master/layout direct background 颁发 hash-bound `setBackground` capability，可增改/清除对应 `p:bg`，只改变所属 `slideMaster`/`slideLayout` part，并通过二次投影恢复；对拥有可识别直接 `a:xfrm` 和固定文本拓扑的 owner-local master/layout placeholder，PPJ 进一步颁发占位符级 `setFrame`/`replaceText` capability，可只改对应 owner part 的坐标、有限 rotation/flip 或文字并二次投影恢复；rotation/flip 的可选属性现在按“属性存在性”而非默认值处理，显式 `0`/`false` 会保留为 native leaf，删除属性会清除对应 `a:xfrm` 属性，`PpjSourceBoundMasterAndLayoutPlaceholderEditsAndReprojects` 已验证 master-only changed-part、Open XML 校验和二次投影恢复；对 slide 上拥有完整 owner-local `a:xfrm` 的 placeholder，PPJ 现在额外开放完整 direct `a:xfrm` 的 `setFrame`（x/y/width/height/有限 rotation/flip），direct/effective frame 投影同样保留可选属性存在性，`PpjSourceBoundDirectSlidePlaceholderFrameEditOnlyChangesSlidePart` 已证明只改 slide part、Open XML 校验和二次投影恢复；对 slide placeholder 没有 direct `a:xfrm`、但可沿 slide→layout→master 以唯一同 type/index 找到完整 frame 的情况，投影器现在提供 effective frame，`setFrame` 编辑首次只在 slide part 物化 owner-local `a:xfrm`，并可一起写入有限 rotation/flip，`PpjInheritedSlidePlaceholderProjectsEffectiveFrameAndMaterializesOnSlideEdit` 已证明 changed-part、Open XML 校验和二次投影恢复；对 layout placeholder 没有 direct `a:xfrm`、但所属 master 存在唯一同 type/index direct frame 的情况，投影器现在提供该 effective frame，`setFrame` 编辑仍由 layout 的 source binding 负责，首次改动才在 layout part 物化 owner-local `a:xfrm`；`PpjInheritedLayoutPlaceholderProjectsEffectiveMasterFrameAndEditsOwnerLocally` 已证明 layout-only changed-part、Open XML 校验和二次投影恢复；页面顺序、sections、custom shows 已有 PPJ 语义。

直接嵌入的 master/layout 图片背景现在也沿同一 owner-local `setBackground` capability 支持 bounded crop/opacity 回写与单 owner 图片替换；`PpjSourceBoundMasterAndLayoutImageBackgroundCropOpacityEditAndReprojects` 证明只改变各自 `slideMaster`/`slideLayout` XML、不触碰 slide、关系或媒体，并在 Open XML 校验及二次投影中恢复两个 owner 的 crop/opacity；`PpjSourceBoundMasterAndLayoutImageBackgroundReplacementClosesRelationshipsAndReprojects` 进一步证明替换时只改变对应 owner XML、`.rels` 和新媒体，并删除不再引用的旧媒体。该证据不扩展到跨 owner 共享图片、外部关系或背景 effect graph。

**仍缺：** 多 master 之间的复杂继承、没有可唯一匹配 layout/master direct frame 的 inherited placeholder（当前会降为 opaque）、effective frame 与实际多级继承链的完整诊断、非文本 placeholder 的内容/图标/关系语义与完整插入删除、任意 imported Master/Layout 图编辑、theme effect/font scheme 的完整继承、handout/notes master 和打印配置。可选 rotation/flip 的存在、显式零/false 与删除已在 owner-local master/layout profile 中闭合；slide placeholder 目前也沿用有限 owner-local profile，但整个继承图和复杂 transform 仍不承诺无损编辑。

**待实现与验收：** 已补 direct master/layout background、owner-local direct placeholder frame/text、slide placeholder 的 direct x/y/width/height/有限 rotation/flip frame，以及“layout 无 direct frame、master 有唯一同 type/index frame”的 bounded effective-frame 投影与 owner-local materialization；本轮再补“slide 无 direct frame、layout/master 有唯一同 type/index frame”的 effective-frame 投影、slide-owner 物化和有限 rotation/flip 写回，并完成 master/layout owner-local transform 的存在、显式零/false、删除三态回投影；`PpjSourceBoundMasterAndLayoutPlaceholderEditsAndReprojects` 证明删除 rotation/flip 后只改 master part，保留显式 false，并在二次投影中恢复属性存在性。slide placeholder 的 `replaceText` 仍只允许固定文本拓扑，paragraph/body style、复杂 inherited transform 与复杂 owner graph 保持 source-owned。常见非文本 placeholder 已加入 source-bound 的只读/有限 frame/text 投影，但未宣称 picture/chart/table 内容关系可由 layout placeholder 自行创建或替换。下一步仍是 theme/master/layout 的完整继承诊断、无法唯一匹配的 inherited placeholder、任意 imported placeholder 图编辑和非文本 placeholder 内容/关系语义。禁止因为改 canvas 就自动缩放、重排或裁剪所有页面；每次布局变化都要有 page-level render review。

### F-03 文本、段落、列表、字段和 WordArt

**段落层级增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.level` 已接通独立赋值、显式零、删除、恢复和
单字段 style 包装删除，范围沿用 0–8 的整数。直接值优先于 authored
owner 默认值；删除清除原生 `lvl`，显式零仍保留属性。编辑要求 level 的
精确字段权限，只改变化段落，保留对齐、缩进、间距、项目符号、原始数值
写法、其它属性、run、相邻段落和非目标 XML/ZIP。非法或越界原生层级
随 no-op 和无关标量赋值/删除保留，拒绝覆盖；修复了其导致投影失败的问题。
最小实验 **3/3 通过**，包含该实验的段落、列表和母版回归
**31/31 通过，0 跳过**。Help、schema、生成资料、预览诊断、可移植性、
reference-sync 和 OpenSpec 检查通过。未重建 NativeAOT；继承列表布局和
预览仍为 partial，完整 F-03 继续开放。

**补齐段落对齐枚举（2026-09-10）：** `text.paragraphs[].style.alignment`
新增 `justifyLow` 和 `thaiDistributed`，分别写为原生 `justLow` 和
`thaiDist`，至此可表达 DrawingML 的七种段落对齐。沿用精确字段权限、
赋值、删除和恢复流程，保留其它段落状态与非目标 XML/ZIP。
表格的直接段落检查已同步，两个模式不会单独导致表格降为 opaque；
母版文本默认值也可往返保留。复用对齐实验，扩展为 **5/5 通过**；
包含该实验的段落、列表、母版和表格回归 **10/10 通过，0 跳过**。
输入诊断和实际 SVG painter 均保留两个模式的 partial 对齐提示与可读文本。
Help、schema、registry 和生成资料已同步，可移植性、reference-sync、
OpenSpec 和提交快照的生成资料/预览诊断检查通过。未重建 NativeAOT，
也未做宿主字形布局验收。完整 F-03 继续开放。

**段落对齐增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.alignment` 已验证 `left`、`center`、`right`、
`justify`、`distributed` 的赋值、删除、恢复和单字段 style 包装删除。
显式 left 保留直接属性，删除则清除它；直接段落值优先于 authored owner
默认值，每次编辑要求 alignment 的精确字段权限。只改变化段落的对齐，
保留缩进、间距、其它属性、run、相邻段落和非目标 XML/ZIP。
本轮复现并修复非法原生 `algn` 导致投影失败的问题：按原始属性文本识别
已建模值，未建模值随 no-op 和无关标量赋值/删除保留，拒绝覆盖。
当时剩余的原生 `justLow`、`thaiDist` 已由上方增量补齐。最小实验 **3/3 通过**，
包含该实验的段落与列表回归 **25/25 通过，0 跳过**。
Help、schema、生成资料、预览诊断、可移植性、reference-sync 和 OpenSpec
检查通过。未重建 NativeAOT，预览文本布局仍为 partial；完整 F-03 继续开放。

**段落悬挂缩进增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.hanging` 已接通正负值赋值、显式零、删除、恢复和
单字段 style 包装删除。范围为 −4032–4032pt，按最近 EMU 取偶数舍入；
正值表示首行向左悬挂，负值表示首行向右缩进，原生 `indent` 取相反符号。
该字段与 PPJ `indent`（左缩进）独立，每次编辑要求 hanging 的精确字段权限。
只改目标段落的直接缩进，保留左缩进、三类间距、其它段落属性、原始数值
写法、run、相邻段落和非目标 XML/ZIP。非法或越界原生缩进随 no-op 和
无关标量编辑保留，拒绝覆盖。复用左缩进实验，最小实验 **6/6 通过**，
相关回归 **350/350 通过，0 跳过**，沿用整组默认样式的已记录基线排除项。
Help、schema、生成资料、预览诊断、可移植性、reference-sync 和 OpenSpec
检查通过。未重建 NativeAOT，预览仍为 partial；完整 F-03 和宿主布局继续开放。

**段落左缩进增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.indent` 已接通赋值、显式零、删除、恢复和单字段
style 包装删除。该字段对应原生 `marL`，范围为 0–4032pt，按最近 EMU
取偶数舍入，与 `hanging` 独立；每次编辑要求 indent 的精确字段权限。
只改目标段落的左缩进，保留悬挂缩进、其它段落属性、三类间距、原始数值
写法、run、相邻段落和非目标 XML/ZIP。原生值按属性文本校验；负值、越界、
非整数等非法左缩进随 no-op 和无关标量编辑保留，拒绝覆盖。最小实验
**3/3 通过**，相关回归 **347/347 通过，0 跳过**，沿用整组默认样式的
已记录基线排除项。Help、schema、生成资料、预览诊断、可移植性、
reference-sync 和 OpenSpec 检查通过。未重建 NativeAOT，预览仍为 partial；
`hanging` 的完整生命周期、完整 F-03 和宿主布局继续开放。

**行距增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.lineSpacing`（正数，至多 1584pt）和
`lineSpacingMultiplier`（正数，至多 132 倍）已接通赋值、单位切换、删除、
恢复和单行距 style 包装删除。同一 style 只选一种行距单位，高优先级样式
同时决定单位和值；点值按百分之一 pt、倍数按十万分之一取偶数舍入，舍入后
为零则拒绝。删除字段清除直接行距，切换单位要求两个字段权限。写回保留
段前/段后间距、原始数值写法、run、相邻段落和非目标 XML/ZIP；重复节点、
未知属性/子内容及缺失、零或非法原生值随无关标量编辑保留，拒绝覆盖。
三种间距共用原始 PPTX 实验，最小实验 **12/12 通过**，相关回归
**344/344 通过，0 跳过**，沿用整组默认样式的已记录基线排除项。
Help、schema、生成资料、预览诊断、可移植性、reference-sync 和 OpenSpec
检查通过。未重建 NativeAOT，预览仍为 partial；完整 F-03 和宿主布局继续开放。

**段后间距增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.spaceAfter`（0–1584pt）和 `spaceAfterMultiplier`
（0–132 倍）已接通赋值、单位切换、删除、恢复和单间距 style 包装删除。
同一 style 的段后间距只选一种单位，高优先级样式同时决定单位和值；点值按
百分之一 pt、倍数按十万分之一取偶数舍入，显式零保留单位，切换要求两个
字段权限。写回保留段前间距、行距、原始数值写法、run、相邻段落和非目标
XML/ZIP；重复节点、未知属性/子内容及缺失或非法原生值随无关标量编辑保留，
拒绝覆盖。复用段前实验的原始 PPTX 检查，段前/段后最小实验 **8/8 通过**，
相关回归 **340/340 通过，0 跳过**，沿用整组默认样式的已记录基线排除项。
schema 已对齐原生范围并拒绝双单位声明；Help、生成资料、预览诊断、
可移植性、reference-sync 和 OpenSpec 检查通过。未重建 NativeAOT，
预览仍为 partial；行距的完整 PPJ 生命周期、完整 F-03 和宿主布局继续开放。

**段前间距增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.spaceBefore`（0–1584pt）和 `spaceBeforeMultiplier`
（0–132 倍）已接通赋值、单位切换、删除、恢复和单间距 style 包装删除。
同一 style 只选一种单位，高优先级样式同时决定单位和值；点值按百分之一 pt、
倍数按十万分之一取偶数舍入，显式零保留原生单位。切换要求两个字段权限。
写回保留其它间距的原始数值写法、run、相邻段落和非目标 XML/ZIP；未知属性/
子内容、重复节点及非法值随无关标量编辑保留，拒绝覆盖。最小原生实验
**4/4 通过**。相关回归初跑 **335/336 通过**；旧实验的 distributed 对齐断言
按既有读取器和夹具修正后，与新增实验复跑 **5/5 通过**。沿用整组默认样式的
已记录基线排除项。schema 已对齐原生范围并拒绝双单位声明；Help、生成资料、
预览诊断、可移植性、reference-sync 和 OpenSpec 检查通过。未重建 NativeAOT，
预览仍为 partial；段后间距、行距的完整 PPJ 生命周期和宿主布局继续开放。

**段落默认柔化边缘增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.softEdge` 已接通赋值、删除、恢复和
单效果包装删除。沿用 `{ radius }`，半径为 0–1000pt，按最近 EMU 舍入，
中点取偶数；显式零保留效果。只修改目标段落的 `a:softEdge`，保留其它效果、
列表属性、直接 run、相邻段落和非目标 XML/ZIP。非法或缺失半径、未知属性/
子内容、重复节点和 DAG 随无关标量编辑保留，拒绝覆盖。最小原生实验
**3/3 通过**，相关回归 **360/360 通过，0 跳过**，沿用整组默认样式的已记录
基线排除项。实验修复了未知效果共存时恢复节点的顺序，以及负原生半径被 SDK
读成零的问题。Help、schema、生成资料、预览诊断、可移植性、reference-sync
和 OpenSpec 检查通过；未重建 NativeAOT，预览仍为 partial，完整 F-03 和
宿主显示继续开放。

**段落默认文字外阴影增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.shadow` 已接通全部 11 个值：color 必填，
opacity、blur、distance、angle、alignment、rotateWithShape、scaleX/Y 和
skewX/Y 可独立省略。支持赋值、删除、恢复和单效果包装删除，保留缺省、零、
一和 false；长度、角度、比例与透明度按原生精度保存。RGB/RGBA、简单主题色、
grammar 同名颜色优先和 tint/shade 均有回投影证据。只修改目标外阴影，保留
其它效果、列表属性、直接 run、相邻段落和非目标 XML/ZIP；重复节点、DAG、
非法几何及未建模颜色内容随无关标量编辑保留，拒绝覆盖。聚焦原生检查
**23/23 通过**，相关回归 **354/354 通过，0 跳过**，沿用整组默认样式的
已记录基线排除项。Help、schema、生成资料、预览诊断、可移植性、reference-sync
和 OpenSpec 检查通过；未重建 NativeAOT，外阴影预览仍为 partial，完整 F-03
和宿主显示继续开放。

**段落默认文字反射增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.reflection` 已接通全部 14 个可选值：
blur、distance、angle、start/endOpacity、start/endPosition、fadeAngle、
scaleX/Y、skewX/Y、alignment 和 rotateWithShape。支持独立赋值、删除、恢复，
保留缺省、零、一和 false；空对象保留反射效果，删除字段或单效果包装则移除。
几何、比例和透明度按原生精度归一化；opacity token、混合效果、相邻段落、
直接 run 和非目标 XML/ZIP 均有实验。未知属性/子内容、非法值、重复节点和 DAG
随无关编辑保留，拒绝覆盖。相关原生回归首次 323/324 通过；旧全跨度 leaf
实验补上显式 startPosition/endPosition 后，与 14 个新增用例复跑 **15/15 通过**。
沿用整组默认样式的已记录基线排除项。Help、schema、生成资料、预览诊断、
可移植性、reference-sync 与 OpenSpec 检查通过；未重建 NativeAOT，
反射预览仍为 partial，完整 F-03 和宿主显示继续开放。

**段落默认文字内阴影增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.innerShadow` 支持赋值、删除、恢复，
以及仅含内阴影的包装删除。color 必填；blur 0–1000pt、distance 0–100000pt、
angle -360–360° 和 opacity/token 可独立省略，显式零和透明度一分别保留。
长度按 EMU 舍入，角度先舍入再归一化；RGB/RGBA、简单主题色、grammar
同名颜色优先和 tint/shade 均有回投影证据。修改目标效果保留其它效果及列表属性、
直接 run、相邻段落和非目标 XML/ZIP；未建模图随无关标量编辑保留并拒绝覆盖。
相关原生回归 **307/307 通过，0 跳过**，沿用整组默认样式的已记录基线排除项。
Help、schema、生成资料、预览诊断与 OpenSpec 已同步；可移植性和 reference-sync
通过，完整 reference-skills 因环境缺少 `pdftoppm` 未完成。未重建 NativeAOT；
内阴影预览仍为 partial，完整 F-03 和宿主显示继续开放。


**段落默认文字发光增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.glow` 支持赋值、删除、恢复和
仅含 glow 的包装删除。radius 接受 0–1000pt，按原生 EMU 舍入；RGB/RGBA、
颜色 token、opacity token，以及透明度缺省/零/一分别保留。简单源主题色
保留 scheme 绑定，grammar 同名颜色优先，tint/shade 解析为 RGB。
各个直接段落默认效果独立读取，修改 glow 保留其它效果及列表属性、相邻段落、
run 和非目标 XML/ZIP。重复 glow/列表、effect DAG、未建模颜色/半径/子内容
随无关标量编辑保留，拒绝覆盖。相关原生回归 **287/287 通过，0 跳过**，
包含 12 个新增 glow 用例；沿用整组默认样式的已记录基线排除项。
资料与 OpenSpec 已同步；glow 预览仍为 partial，完整 F-03 和宿主显示继续开放。


**段落默认文字渐变增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.gradient` 支持线性/居中径向渐变、
删除、恢复和单字段包装删除。2–16 个有序色标允许重复位置；RGBA/token
解析为 RGB 与色标透明度，显式 opacity 0/1 保留，角度在原生舍入后归一化。
源编辑中 grammar 同名颜色优先。纯色与渐变切换须同时拥有两个字段的权限。
修改一段保留其它段的 tileRect、alpha 精度、run 颜色和非目标 XML/ZIP。
不支持的缩放、主题色标、重复填充和未知子内容随无关标量编辑保留，拒绝覆盖。
相关 **248/248 通过，0 跳过**，沿用整组默认样式基线排除项；补强颜色 token
碰撞用例后，4 个生命周期用例再次通过。旧百分比径向渐变测试已按现有读取器
校正为“居中 50% 支持、非居中 25% 拒绝”。资料与 OpenSpec 同步，
文字渐变预览仍为 partial；完整 F-03 和宿主显示继续开放。

**段落默认文字颜色增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.color` 支持 RGB/RGBA、颜色 token、
删除、恢复和单字段包装删除。显式 alpha 0/1 与缺省分别保留；简单源主题色
保持 scheme 绑定，显式 grammar 同名颜色优先。只改目标段落，保留其它段落
原生 alpha 精度、主题绑定、run 颜色和非目标 XML/ZIP。源亮度变换、重复填充
和未建模颜色随无关标量编辑保留，拒绝覆盖；已投影的变换颜色也拒绝删除。
渐变冲突与渐变修改保留独立边界。既有相关回归 228 项通过；颜色 9 项最初
有 2 项 JSON 属性顺序断言错误，修正并补充渐变反例后 **9/9 通过，0 跳过**。
沿用整组默认样式的已记录基线排除项。资料、生成检查与 OpenSpec 已同步；
渐变生命周期、宿主显示和完整 F-03 继续开放。

**段落默认高亮增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.highlight`
支持不透明 RGB、颜色 token、删除、恢复和单字段包装删除。直接原生主题色
投影为 `{"token":"accent1"}`，源编辑保留未加变换的主题绑定；
显式 grammar 同名颜色优先，tint/shade 解析成 RGB。修改一段保留其它段落的
主题绑定、run 高亮、原文字、其它样式和非目标 XML/ZIP。
透明度结果拒绝；带变换、重复节点或未知子元素的源高亮保留并拒绝覆盖。
非法颜色裸文本会被 SDK 解析丢弃，因此无操作原字节保留、编辑拒绝输出。
相关 **228/228 通过，0 跳过**，包括既有高亮回归；沿用已记录的整组默认样式
基线排除项。资料、生成检查和 OpenSpec 已同步，宿主显示及完整 F-03 继续开放。



**段落默认下划线增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.underline`
支持现有全部下划线枚举、single/double 别名、显式 none、删除、恢复和
单字段包装删除。原生 sng/dbl 二次投影为 PPJ single/double，none 与缺省
分别保留。原文字、run 下划线、其它默认样式和非目标 XML/ZIP 保留。
未知 u 值及下划线填充/线条子节点拒绝覆盖，并随无关标量修改或删除保留；
实验包含 uFillTx 有无 u 属性的两种源文件。
相关 **214/214 通过，0 跳过**，沿用整组默认样式的已记录基线排除项。
资料、生成检查和 OpenSpec 已同步；独立下划线效果、宿主字形和完整 F-03
继续开放。



**段落默认删除线增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.strike`
支持 true/false、noStrike/sngStrike/dblStrike、显式取消、删除、
恢复和单字段包装删除。布尔别名二次投影为标准字符串，noStrike 与缺省
分别保留；原文字、run 删除线、其它默认样式、未知属性和非目标 XML/ZIP
保留。未建模原生 strike 拒绝覆盖，无关标量修改或删除仍保留它。
相关 **207/207 通过，0 跳过**，沿用已记录的整组默认样式基线失败排除项。
包含单线→显式取消→删除的连续请求；资料、生成检查和 OpenSpec 已同步。
主机字形显示仍需单独证据，完整 F-03 继续开放。



**段落默认大小写样式增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.capitalization`
支持 none/small/all、显式 none、删除、恢复和单字段包装删除。none 与缺省
分别保留；原始文字、run 大小写样式、其它默认样式、未知属性和非目标
XML/ZIP 保留。未建模原生 cap 拒绝覆盖，无关标量修改或删除仍保留它，
沿用来源已有校验警告、不引入新警告的导出规则。
相关 **202/202 通过，0 跳过**，沿用已记录的整组默认样式基线失败排除项。
资料、生成检查和 OpenSpec 已同步；主机小型大写字形仍需单独证据。



**段落默认基线偏移增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.baseline`
支持独立赋值、显式零、删除、恢复及单字段包装删除，单位为百分比，
范围 -400–400%，原生精度 0.001%，中点向偶数舍入。接近零的负值按
原生整数归一为显式零。直接 run 基线、其它默认样式、未知属性和
非目标 XML/ZIP 保留；未建模原生 baseline 拒绝覆盖，随其它标量修改或删除保留。
相关 **197/197 通过，0 跳过**，沿用已记录的整组默认样式基线失败排除项。
资料、生成检查和 OpenSpec 已同步；字体实际排版仍需主机证据。



**段落默认字符间距增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.letterSpacing`
支持独立赋值、显式零、删除、恢复及单字段包装删除，范围 -768–768pt，
原生精度 0.01pt，中点向偶数舍入。接近零的负值按原生整数归一为显式零，
修复浮点负零引起的写出后语义不一致。直接 run 间距、其它默认样式、
未知属性和非目标 XML/ZIP 保留；未建模 spc 拒绝覆盖，仍随其它字段修改保留。
相关 **192/192 通过，0 跳过**，沿用已记录的整组默认样式基线失败排除项。
资料、生成检查和 OpenSpec 已同步；字体实际排版仍需主机证据。



**段落默认 kerning 增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.kerning`
表示启用字偶距调整的最小字号阈值，单位 pt，范围 0–768。
支持独立赋值、显式零、删除、恢复和单字段包装删除；原生精度为 0.01pt，
中点向偶数舍入。直接 run 的 kerning、其它默认样式、未知属性和非目标
XML/ZIP 保留。未建模的原生 kern 拒绝覆盖，修改或删除其它标量仍保留它。
相关 **187/187 通过，0 跳过**，沿用已记录的整组默认样式基线失败排除项。
资料、生成检查和 OpenSpec 已同步；字体实际排版仍需主机证据。



**段落默认语言增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.language`
支持独立赋值、删除、恢复，以及仅包含语言的 defaultText/style 包装删除。
沿用 2–63 字符的标签语法，保留大小写；run 语言、altLang、字体/效果、
其它段和非目标 XML/ZIP 保留。来源中的未建模 lang 保留，拒绝覆盖，
修改其它默认标量仍可进行。修复了删除时误写 lang="" 的问题。
相关 **182/182 通过，0 跳过**，沿用已记录的整组默认样式基线失败排除项。
字段资料、生成检查和 OpenSpec 已同步；主机校对和复杂继承仍是独立缺口。


**段落默认复杂脚本字体增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.fontFamilyComplexScript`
支持独立赋值、删除、恢复和单字体包装删除，仅更新简单直接的 a:cs。
名字非空白、最多 255 字符，主题字体引用保留原写法；
额外元数据和隐藏子内容拒绝覆盖。拉丁/东亚字体、效果、直接 run、
其它段和非目标 XML/ZIP 保留。相关 **177/177 通过，0 跳过**，
沿用已记录的整组默认样式基线失败排除项。三类直接默认字体的字段生命周期
已分别接通；资料、生成检查和 OpenSpec 通过，现有协议不变，未重建 NativeAOT。


**段落默认东亚字体增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.fontFamilyEastAsia`
支持独立赋值、删除、恢复和单字体包装删除。source-bound 按显式字段决定存在性，
保留 fontFamily 时也能真正删除 a:ea，作者模式的字体补全不会把它重新写回。
名字非空白、最多 255 字符；额外元数据和隐藏子内容继续拒绝覆盖。
拉丁/复杂脚本字体、效果、直接 run、其它段及非目标 XML/ZIP 保留。
相关 **171/171 通过，0 跳过**，沿用已记录的整组默认样式基线失败排除项。
资料、生成检查和 OpenSpec 通过；现有协议及作者模式补全不变，未重建 NativeAOT。


**段落默认字体增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.fontFamily` 支持独立赋值、删除、恢复，
以及仅含该字体的包装对象删除。它只修改简单直接的 a:latin；
名字不能为空白，最多 255 字符，主题字体引用保留原写法。
东亚/复杂脚本字体、其它默认样式、效果、直接 run 字体和非目标 XML/ZIP 保留。
带额外元数据或隐藏子内容的字体节点拒绝替换。作者模式会补东亚字体，
故单字体来源实验显式移除该补充值，多字体实验则验证其保留。
相关 **165/165 通过，0 跳过**，沿用已记录的整组默认样式基线失败排除项。
资料、生成检查和 OpenSpec 通过；现有协议不变，字体可用性和宿主替代尚未验收。


**段落默认字号增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.size` 支持赋值、删除、恢复，
也支持移除仅含字号的 defaultText/style 包装；要求独立字段权限。
字号范围为 1..768pt，以 0.01pt 精度就近舍入，中点取偶数
（18.256→18.26，18.125→18.12）。首次实验发现 0.01pt 不满足原生字号下限，
已补前置校验并保留拒绝用例。只更新直接 sz，保留其它默认样式、效果、
直接 run、第二段、未知属性及非目标 XML/ZIP。相关 **159/159 通过，0 跳过**，
沿用下方已复现的整组默认样式基线失败排除项。资料、生成检查和 OpenSpec 通过；
未重建 NativeAOT，字体、其它默认样式及继承继续逐项补齐。


**段落默认斜体增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.italic` 支持 true、false、删除和恢复，
含仅有斜体的包装对象删除；与 bold 分别检查字段权限。共享实验验证其它默认样式、
soft-edge、直接 run、未知原生属性及非目标 XML/ZIP 保留，未改动布尔属性的 XML
写法也保持原样。相关 **155/155 通过，0 跳过**；沿用下一段记录的基线失败排除项。
schema、资料和 OpenSpec 同步通过，现有协议不变；其它默认字段与继承继续逐项补齐。


**段落默认粗体增量（2026-09-10）：** 普通文本框和形状的
`text.paragraphs[].style.defaultText.bold` 支持 true、false、删除和恢复，
也支持删除仅含 bold 的包装对象；要求对应的 `setTextParagraphStyle` 字段权限。
仅粗体变化时只更新直接 `a:defRPr/@b`，保留其它默认样式、soft-edge、run 直接样式、
原生 dirty 属性和非目标 XML/ZIP。专项 5/5，相关 151/151，0 跳过。
旧的 `ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle`
在干净基线 `bbc1065b` 仍有写后语义不匹配，已单独记录并排除在 151 项之外。
tab-stop 测试改为复用同一份来源字节，避免重复打包的哈希漂移。资料与 OpenSpec
检查通过；复用现有协议，未重建 NativeAOT，其它默认样式字段及继承仍待逐项处理。


**文字变形删除增量（2026-09-10）：** `textWarpPreset` 与
`textWarpAdjustments` 支持删除和恢复。保留 preset 时，删除或清空参数数组
只清除直接参数；删除 preset 时须同时清除依赖参数。`textNoShape` 保持显式值，
参数顺序、0 和 signed 32-bit 正负边界均可导出、重新投影。
普通文本、形状、母版/布局占位符及表格单元格共用最小实验，包含简单样式删除、
表格紧凑文字恢复和非目标 XML/ZIP 保留。相关原生 **160/160 通过，0 跳过**。
新增删除字段 46，旧 setter 37/参数 38 编码不变；同时修正该 native leaf 的
schema 范围，并拒绝参数叶节点的非法嵌套内容。协议、资料及 OpenSpec 检查通过。
需要更新后的 codec；未重建 NativeAOT，完整 WordArt 渲染仍待补齐。


**平面文字深度删除增量（2026-09-10）：** `flatTextZ` 支持显式 0、
signed 32-bit 正负边界、删除和恢复；删除只移除 canonical `a:flatTx`，
保留文字变形及其它 XML/ZIP 状态。普通文本、形状、母版/布局占位符和表格单元格
共用最小回归，包含简单样式删除和表格紧凑文字恢复。
相关原生回归 **152/152 通过，0 跳过**（SDK 8.0.128）。
wire 保留 setter 39，新增删除字段 45；边界实验同时修正该 native leaf 的
schema 数值范围，并补上 SDK 叶节点非法嵌套内容的拒绝检查。
协议、资料同步、预览限制提示及 OpenSpec 检查通过；需要更新后的 codec，
未重建 NativeAOT，未进行完整 3D 文字渲染验收。


2026-09-10 增量：`fromWordArt` 已区分 true/false/缺失，删除只移除原生
来源标记；`textArchUp` 文字变形及其它 XML/ZIP 内容保留。新增删除标记 44，
旧布尔字段 36 编码不变。共享实验新增 7 个生命周期场景和 1 个协议场景，
相关原生 144/144、0 跳过，C#/JS 编码、proto:check 与资料检查通过。
新删除需要更新 codec，本轮未重建 NativeAOT，完整 WordArt 渲染仍待补齐。


2026-09-10 增量：`compatibleLineSpacing` 已区分 true/false/缺失，
删除移除 `compatLnSpc`，保留实际行距、段前/段后间距与其它文本状态。
新增删除标记 43，旧布尔字段 35 编码不变。共享实验新增 7 个生命周期场景
和 1 个协议场景，相关原生 136/136、0 跳过，C#/JS 编码、proto:check
与资料检查通过。新删除需要更新 codec，本轮未重建 NativeAOT；
提示字段不代表宿主行距度量已验收。


2026-09-10 增量：`spaceFirstLastParagraph` 已补齐 true/false/缺失的源语义，
删除移除 `spcFirstLastPara`，保留实际段前/段后间距及其它文本属性。
新增删除标记 42，旧布尔字段 34 编码不变。共享实验新增 7 个生命周期场景
（含显式段前 7、段后 9）和 1 个协议场景，相关原生 128/128、0 跳过；
C#/JS 编码、proto:check 与资料检查通过。新删除需要更新 codec，本轮未重建
NativeAOT，也未声称宿主段落度量已验收。


2026-09-10 增量：`forceAntiAlias` 已区分 true/false/缺失，删除移除原生
`forceAA`，保留 `anchorCenter` 与垂直对齐。新增删除标记
`no_force_anti_alias=41`，旧字段 33 编码不变；false 删除及设置/删除并存被拒绝。
复用 7 个布尔生命周期场景和 1 个协议场景，相关原生 120/120、0 跳过，
C#/JS 编码、proto:check 与资料检查通过。新删除需要更新 codec；未重建
NativeAOT，提示字段不代表宿主抗锯齿效果已验收。


2026-09-10 增量：`anchorCenter` 已区分 true/false/省略；删除移除原生
`anchorCtr`，保留 `verticalAlignment`。协议新增可选删除标记
`no_anchor_center=40`，旧布尔字段 32 编码不变；拒绝 false 删除命令及设置/删除
并存。新增 7 个共享生命周期场景和 1 个原生命令实验，相关原生 112/112、
0 跳过；C#/JS 编码、proto:check 及字段资料检查通过。删除需要更新后的 codec，
本轮未重建 NativeAOT，完整继承与宿主锚点排版仍待补齐。


2026-09-10 增量：`autoFit` 与 `normalAutoFit` 已补齐源删除与恢复。
显式 `none` 保留 `noAutofit`；删除模式及其关联 profile 才移除节点。
保留 `shrink-text` 时，删除百分比只移除对应属性，删除 profile 保留空
`normAutofit`。新增 7 个共享用例覆盖五类 owner、三模式、百分比边界/小数/零、
依赖拒绝、简单样式与表格紧凑文字恢复、原源 no-op 和非目标 XML/ZIP 保留。
相关专项 104/104、0 跳过，字段说明与预览限制已同步；非规范源拓扑、继承和
宿主重排继续保留原边界。


2026-09-10 增量：`margins.left/top/right/bottom`（0..10000 points）已补齐
单边删除、空对象、整组删除和恢复。显式 0 保留，删除只移除对应原生 inset；
其它边距与正文保留。共享实验新增 7 例覆盖五类文字 owner、0/12.5/10000、
简单样式整组删除及表格紧凑文字恢复；相关原生 96/96、0 跳过，no-op 与非目标
XML/ZIP 保留通过。字段说明及预览限制已同步；完整继承与宿主重排仍待补齐。


2026-09-10 增量：`columns`（整数 1..16）补齐 source-bound 删除与恢复；
显式单栏 `1` 与缺失分开保留。删除映射为移除原生 `numCol`，保留
`columnGap` 和 `columnDirection`。现有数值属性实验增加 7 个场景，覆盖
text/shape/master/layout/table、1/3/16 恢复、简单样式整组删除、表格紧凑文字
恢复、no-op 字节及非目标 XML/ZIP 保留。相关原生回归 89/89、0 跳过；
schema、Help、registry、Skill 与预览限制已同步。完整继承和宿主分栏排版仍待补齐。


**分栏间距删除增量（2026-09-10）：** `columnGap` 以 points 表达 0..10000，显式 0 保留原生零间距，删除移除 `bodyPr/@spcCol`。文本、shape、master/layout placeholder 和表格已闭合源删除及恢复；新增 7 个共享生命周期用例验证 0/12.5/10000 的实际 EMU、原源 no-op、简单样式删除、表格纯文本后的结构化恢复，以及 columns/columnDirection/其它 XML/非目标 ZIP 保留。相关专项 82/82、零跳过；简单样式删除字段新增 columnGap。预览保留分栏诊断，宿主排版、完整继承和其它属性删除仍开放。

**垂直对齐删除增量（2026-09-10）：** `verticalAlignment` 已闭合 top/middle/bottom/省略的源编辑与恢复；middle 保持原生 Center 映射，删除移除 `bodyPr/@anchor`，独立的 anchorCenter 保留。文本、shape、master/layout placeholder 和表格共用枚举实验，新增 7 例，相关专项 75/75、零跳过；覆盖原源 no-op、三值恢复、简单样式整体删除、表格纯文本后的结构化恢复及其它 XML/非目标 ZIP 保留。简单样式删除字段新增 verticalAlignment；anchorCenter 等其它字段仍受既有删除边界约束。预览保留布局诊断，宿主排版和完整继承仍开放。

**垂直溢出删除增量（2026-09-10）：** `verticalOverflow` 的 overflow/ellipsis/clip/省略已闭合源编辑与恢复；删除移除 `bodyPr/@vertOverflow`，保留独立的 horizontalOverflow。文本、shape、master/layout placeholder 和表格共用枚举实验，新增 7 例，相关专项 68/68、零跳过；覆盖原源 no-op、三值恢复、简单样式整体删除、表格纯文本后的结构化恢复，以及其它 XML/非目标 ZIP 保留。可整体删除的简单样式字段现在包括 upright、rotation、columnDirection、verticalText、wrap、horizontalOverflow、verticalOverflow。预览保留字段诊断，宿主省略号/裁切、完整继承和其它属性删除仍开放。

**水平溢出删除增量（2026-09-10）：** `horizontalOverflow` 的 overflow/clip/省略已闭合源编辑与恢复；删除只移除 `bodyPr/@horzOverflow`，保留独立的 verticalOverflow。文本、shape、master/layout placeholder、表格共用枚举生命周期实验，新增 7 例，相关专项 61/61、零跳过；覆盖原源 no-op、双值恢复、简单样式删除、表格纯文本后的结构化恢复及其它 XML/非目标 ZIP 保留。简单样式删除字段在此前五项基础上增加 horizontalOverflow；其它属性删除仍受约束。预览保留字段诊断，宿主裁切及完整继承仍开放。

**换行模式删除增量（2026-09-10）：** `wrap` 已区分 `square`、`none` 和省略；显式 none 保留原生不换行属性，删除字段才移除 `bodyPr/@wrap`。文本、shape、master/layout placeholder 和表格复用枚举生命周期实验，新增 7 例，相关专项 54/54、零跳过，覆盖原源 no-op、双值恢复、简单样式整体删除、表格纯文本后的结构化恢复及其它 XML/非目标 ZIP 保留。可整体删除的简单样式现可由 upright、rotation、columnDirection、verticalText、wrap 组成，其它字段继续受删除边界约束。预览保留换行布局诊断，宿主自动换行和完整继承仍开放。

**文字方向删除增量（2026-09-10）：** `verticalText` 的 `horizontal`、`vertical`、`vertical270` 与省略已接通源编辑：显式 horizontal 保留 `vert=horz`，删除才移除原生 `bodyPr/@vert`。文本、shape、master/layout placeholder 和表格均复用方向生命周期实验，新增 7 例，相关专项 47/47、零跳过；检查原源 no-op、三种值恢复、简单样式整体删除、表格纯文本后的结构化恢复，以及其它正文/frame/非目标 ZIP 保留。可整体删除的简单样式字段现为 upright、rotation、columnDirection、verticalText。预览保留文字布局诊断，宿主排版、继承及其它属性删除继续开放。

**分栏方向删除增量（2026-09-10）：** `columnDirection` 的 `left-to-right`、`right-to-left` 与省略已闭合 authored/去嵌入/source-bound/二次投影；前两者保留原生 `rtlCol=false/true`，删除才移除属性。文本、shape、master/layout placeholder 和表格共用生命周期实验；只含该字段、upright、rotation 的样式可整体移除，表格变回纯文本后可按原段落/run 拓扑恢复结构化样式。新增 7 例，相关专项 40/40、零跳过，验证原源 no-op、其它正文属性/frame/非目标 ZIP 保留。预览保持分栏诊断；其它属性删除、继承和宿主排版仍开放。

**文本体 rotation 完整增改删（2026-09-10）：** `style.rotation` / `textStyle.rotation` / 表格 `text.style.rotation` 删除后移除原生 `bodyPr/@rot`，显式 0 与正负角度仍保留。文本、形状、master/layout 占位符及表格均支持原源删除和恢复；只含 rotation/upright 的样式对象可整体移除。新增 7 个最小回归核对实际 60000 分之一度数值、原源 no-op、0/负角恢复、表格普通文本规范化后的字段重建，以及 frame/正文/其它属性和非目标 ZIP 保留；相关专项 33/33 通过。预览保留文本旋转限制，未重建 NativeAOT 或做宿主布局验收，其它 body 字段删除仍逐项推进。

**文本 upright 完整增改删（2026-09-10）：** `style.upright` / `textStyle.upright` / 表格 `text.style.upright` 保留 true、false 与省略的区别；从已投影源样式删除该字段，会移除原生 `upright` 属性，恢复继承/默认行为。样式只含 upright 时也可删除整个对象。文本、形状、master/layout 占位符和表格的最小实验 7/7 通过，核对原源 no-op、删除后再添加 true/false、属性实际存在性、其它 XML 语义及非目标 ZIP 保留、缺权限和连带删除拒绝。表格删除最后一个 body 属性后可规范投影为普通文本，再以相同原生拓扑恢复结构化样式。相关专项 103/104，唯一异常合并表格断言在干净 `bbc1065b` 基线上同样失败；预览保留布局限制，未重建 NativeAOT 或做宿主验收。

**优先级：P0；状态：部分完成（rich text 子集已交付）。**

**当前进度：** PPJ 有 string/rich text、paragraph/run、字体、字号、颜色、渐变、阴影、项目符号、段落间距、缩进、文本框边距、列、方向、AutoFit、垂直文字和有限 inline LaTeX；新增 typed `run.field`（固定 `type/text`、可选花括号 UUID）可 authored 编译为 `p:fld` 并在去嵌入 PPJ 后恢复字段类型/文本/ID；`run.break: true` 现在可以表示一个不新建段落的原生 DrawingML line break，并在普通文本和固定拓扑表格 cell 中 authored/投影；source-bound 现在允许在保持 field ID/type 不变的前提下修改静态 display text，只改目标 SlidePart，并通过二次投影恢复；普通文本框、带文本形状和占位符的固定拓扑 text body 现在还可通过独立 `setTextBodyStyle` 回写直接 bodyPr 的 vertical alignment、wrap、四边 inset、columns、column gap/direction、vertical text、rotation、horizontal/vertical overflow、upright 和有限 auto-fit，保持段落/run 拓扑与未建模 XML，`PpjSourceBoundTextBodyStyleEditsTextShapeAndReprojects` 覆盖 authored、capability、SlidePart-only source-bound 编辑和二次投影；固定拓扑表格 cell 也允许 direct text/field/break 混合 body，只修改文字或字段缓存文本并保持字段 ID/type、段落/run 拓扑和其余 XML；普通文本、默认 run、表格和图表样式新增 `text.fontFamilyComplexScript`/`fontFamilyComplexScript`，authored 与 imported/source-bound 单叶回归分别证明直接 `a:cs` 的写入、投影和 token-splice 编辑；已有多个 source-bound text leaf。

**文本 bodyPr hint 增量：** `textBoxStyle.forceAntiAlias` / `textBodyForceAntiAlias` 绑定直接 `a:bodyPr/@forceAA`，只接受已有 canonical `0|1` token；`PpjSourceBoundTextBodyForceAntiAliasEditsLeafAndReprojects` 覆盖 authored → 去嵌入投影 → 单属性 token splice → SlidePart-only changed-part → Open XML → 二次投影，缺失属性不发 leaf。该字段只保留 DrawingML 的显式 anti-aliasing hint，不宣称跨宿主渲染一致性。

**文本 bodyPr 段落间距 hint 增量：** `textBoxStyle.spaceFirstLastParagraph` / `textBodySpaceFirstLastParagraph` 绑定直接 `a:bodyPr/@spcFirstLastPara`，只接受已有 canonical `0|1` token；`PpjSourceBoundTextBodySpaceFirstLastParagraphEditsLeafAndReprojects` 覆盖 authored → 去嵌入投影 → 单属性 token splice → SlidePart-only changed-part → Open XML → 二次投影，缺失属性不发 leaf。该字段只保留 DrawingML 的显式首尾段落间距 hint，不推断宿主段落度量或自动 reflow。

**文本 bodyPr 兼容行距 hint 增量：** `textBoxStyle.compatibleLineSpacing` / `textBodyCompatibleLineSpacing` 绑定直接 `a:bodyPr/@compatLnSpc`，只接受已有 canonical `0|1` token；`PpjSourceBoundTextBodyCompatibleLineSpacingEditsLeafAndReprojects` 覆盖 authored → 去嵌入投影 → 单属性 token splice → SlidePart-only changed-part → Open XML → 二次投影，缺失属性不发 leaf。该字段只保留 DrawingML 的显式兼容行距 hint，不推断宿主行距或自动 reflow。

**文本 bodyPr WordArt 标记增量：** `textBoxStyle.fromWordArt` / `textBodyFromWordArt` 绑定直接 `a:bodyPr/@fromWordArt`，只接受已有 canonical `0|1` token；`PpjSourceBoundTextBodyFromWordArtEditsLeafAndReprojects` 覆盖 authored → 去嵌入投影 → 单属性 token splice → SlidePart-only changed-part → Open XML → 二次投影，缺失属性不发 leaf。该字段只保留 DrawingML 的显式 WordArt 标记，不声称支持 WordArt 变换、效果或宿主等价渲染。

**文本 bodyPr warp 预设与调整增量：** `textBoxStyle.textWarpPreset` / `textBodyWarpPreset` 绑定直接 `a:prstTxWarp/@prst`，只接受有限 DrawingML `TextShapeValues` 词表；配套的可选 `textBoxStyle.textWarpAdjustments` / `textBodyWarpAdjustment` 绑定唯一、属性-only 的 `a:avLst/a:gd name="..." fmla="val N"` 有序列表，并按索引只替换一个 `val` 数字 token。`PpjSourceBoundTextWarpAdjustmentEditsLeafAndReprojects` 覆盖 authored → 去嵌入投影 → 单值 token splice → SlidePart-only changed-part → Open XML → 二次投影；空列表、公式、扩展、重复名和未知拓扑保持 source-owned。该字段只保留 WordArt 的预设名和有限字面调整值，不声称支持完整变换、效果或宿主等价渲染。

**文本 glow 增量：** `textStyle.glow` 已覆盖 rich-text run 的 `style` 与段落 `style.defaultText`，并沿用命名 text style、layout/master 默认 text style 的 authored 路径；它接受 direct RGB/theme color、`0..1000` point radius 和 numeric/`opacity` token，NativeAOT 写入直接 `a:effectLst/a:glow`，与已有 text outer shadow 保持 glow → outer shadow 顺序。`PpjAuthoredTextGlowEffectWritesRunAndDefaultRunOwners` 已验证 run 与 `a:defRPr` 两个 owner 的 XML、半径/alpha、嵌入 PPJ 恢复和非法半径拒绝。对去嵌入 PPJ 后的 imported/source-bound 直接 rich-text run，只有单一直接 `a:glow` 或 `a:glow` 后接一个有界 `a:outerShdw` 的 effect list 会投影为 `textStyle.glow`，并颁发 `textGlowRadiusEmu`、`textGlowColorRgb`/`textGlowColorScheme`、显式 alpha 三类 native leaf；`PpjSourceBoundTextGlowEditsDirectRunOwnerAndReprojects` 已证明三类值只改所属 SlidePart 并二次投影恢复。段落 `defaultText` 的 strict owner 现在为直接 glow radius/RGB/theme color/explicit alpha 颁发 `textDefaultGlowRadiusEmu`、`textDefaultGlowColorRgb`、`textDefaultGlowColorScheme` 和 `textDefaultGlowOpacityThousandthPercent`，同一 focused fixture 已证明它与 inline/default soft-edge/inner-shadow/reflection owner 独立回写；defaultText 的其它未建模 effect graph、WordArt 和复杂 effect graph 继续 source-owned 或 fail closed。

**文本 inner shadow 增量：** `textStyle.innerShadow` 已复用同一 rich-text run/default-run owner 路径，覆盖命名 text style、layout/master 默认 text style 的 authored 解析；它接受 direct RGB/theme color、`blur`/`distance`/`angle` 和 numeric/`opacity` token，NativeAOT 写入直接 `a:effectLst/a:innerShdw`，与 text glow、outer shadow 保持 glow → inner shadow → outer shadow 顺序。`PpjAuthoredTextInnerShadowEffectWritesRunAndDefaultRunOwners` 已验证 run 与 `a:defRPr` 两个 owner 的几何、颜色/alpha、嵌入 PPJ 恢复和非法 blur 拒绝。对去嵌入 PPJ 后的 imported/source-bound 直接 rich-text run，只有单一直接 `a:innerShdw` 或 `a:innerShdw` 后接一个有界 `a:outerShdw` 的 effect list 会投影为 `textStyle.innerShadow`，并颁发存在的 blur/distance/direction、RGB/主题色和显式 alpha native leaf；`PpjSourceBoundTextInnerShadowEditsDirectRunOwnerAndReprojects` 已证明五个可用值只改所属 SlidePart 并二次投影恢复。段落 `defaultText` 现在额外颁发 `textDefaultInnerShadowBlurRadiusEmu`、`textDefaultInnerShadowDistanceEmu`、`textDefaultInnerShadowDirectionDegrees`、`textDefaultInnerShadowColorRgb`/`textDefaultInnerShadowColorScheme` 和 `textDefaultInnerShadowOpacityThousandthPercent`，覆盖直接 inner-shadow blur/distance/direction/RGB/theme color/显式 alpha；它与 glow radius、soft-edge radius leaf 一样保留 paragraph index proof，只有已有直接 `a:alpha/@val` 才能回写，缺失或隐式 alpha 仍 source-owned；reflection 的其它标量、WordArt 和复杂 effect graph 继续 source-owned 或 fail closed。

**文本 defaultText outer shadow 增量：** 段落 `style.defaultText.shadow` 的 strict imported/source-bound owner 现在颁发 `textDefaultShadowBlurRadiusEmu`、`textDefaultShadowDistanceEmu`、`textDefaultShadowDirectionDegrees`、`textDefaultShadowAlignment`、`textDefaultShadowColorRgb`、`textDefaultShadowColorScheme`、`textDefaultShadowOpacityThousandthPercent` 与 `textDefaultShadowRotateWithShape`，仅接受 `a:p/a:pPr/a:defRPr/a:effectLst/a:outerShdw` 下已有且有界的直接 `blurRad`/`dist`/`dir`/`algn`，一个直接 `srgbClr/@val` 或 `schemeClr/@val`，该颜色下一个已有的直接 `a:alpha/@val`，以及一个已有的显式 `rotWithShape` 0/1 属性；paragraph index proof 将叶子绑定到具体 `a:defRPr`，edit plan 只 splice 对应的 `outerShdw/@blurRad`、`outerShdw/@dist`、`outerShdw/@dir`、`outerShdw/@algn`、`outerShdw/srgbClr/@val`、`outerShdw/schemeClr/@val`、已有的 `a:alpha/@val` 或 `outerShdw/@rotWithShape`，不重建 effect list。`PpjSourceBoundTextSoftEdgeEditsDirectRunOwnerAndReprojects` 的 focused fixture 另放一个 default outer-shadow 段落，证明 authored blur/distance/direction/alignment/RGB/theme color/opacity/rotate-with-shape、source-bound 单字段编辑、仅所属 SlidePart 改变、未建模负例保持 fail closed，以及二次投影恢复八项值；未知/扩展和其它 effect graph 继续 source-owned。

**文本 reflection 增量：** `textStyle.reflection` 已复用同一 rich-text run/default-run owner 路径，覆盖命名 text style、layout/master 默认 text style 的 authored 解析；它接受 `blur`、`startOpacity`/`endOpacity`、`distance` 和 `angle`，NativeAOT 写入直接 `a:effectLst/a:reflection`，使用确定性的 `stPos=0`/`endPos=100000`，并保持 glow → inner shadow → outer shadow → reflection 顺序。`PpjAuthoredTextReflectionEffectWritesRunAndDefaultRunOwners` 已验证 run 与 `a:defRPr` 两个 owner 的几何、起止 alpha、全跨度位置、嵌入 PPJ 恢复和非法 opacity 拒绝。对去嵌入 PPJ 后的 imported/source-bound 直接 rich-text run，只有一个 full-span 直接 `a:reflection` 或其前接一个有界 `a:outerShdw` 的 effect list 会投影为 `textStyle.reflection`，并颁发存在的 blur/startOpacity/endOpacity/distance/direction 五类 native leaf；`PpjSourceBoundTextReflectionEditsDirectRunOwnerAndReprojects` 已证明五个值只改所属 SlidePart、保留 full-span owner proof 并二次投影恢复。段落 `defaultText` 现在也颁发 `textDefaultReflectionBlurRadiusEmu`、`textDefaultReflectionDistanceEmu`、`textDefaultReflectionStartOpacityThousandthPercent`、`textDefaultReflectionEndOpacityThousandthPercent` 和 `textDefaultReflectionDirectionDegrees`（仅 direct full-span blur/distance/start/end opacity/direction），并分别颁发 glow radius、shadow blur、inner-shadow blur 与 soft-edge radius leaf；非 full-span、transform/unknown/extension effect graph 继续 source-owned 或 fail closed。

**文本 soft edge 增量：** `textStyle.softEdge` 已复用同一 rich-text run/default-run owner 路径，覆盖命名 text style、layout/master 默认 text style 的 authored 解析；它接受 `0..1000` point radius，NativeAOT 写入直接 `a:effectLst/a:softEdge`，并保持 glow → inner shadow → outer shadow → reflection → soft edge 顺序。`PpjAuthoredTextSoftEdgeEffectWritesRunAndDefaultRunOwners` 已验证 run 与 `a:defRPr` 两个 owner 的半径、四项 effect 顺序、嵌入 PPJ 恢复和非法半径拒绝。对去嵌入 PPJ 后的 imported/source-bound 直接 rich-text run，只有一个直接 `a:softEdge`，或其前接一个已验证的 `a:outerShdw` / full-span `a:reflection`，会投影为 `textStyle.softEdge`，并颁发 `textSoftEdgeRadiusEmu`；段落 `defaultText` 的同一 strict owner 现在颁发 `textDefaultSoftEdgeRadiusEmu`，按段落 index 定位 `a:defRPr` 并只 splice `softEdge/@rad`；扩展后的 focused fixture 还证明 `textDefaultGlowRadiusEmu`、`textDefaultGlowColorRgb`、`textDefaultGlowColorScheme`、`textDefaultGlowOpacityThousandthPercent`、`textDefaultShadowBlurRadiusEmu`、`textDefaultShadowDistanceEmu`、`textDefaultShadowDirectionDegrees`、`textDefaultShadowAlignment`、`textDefaultShadowColorRgb`、`textDefaultShadowColorScheme`、`textDefaultShadowOpacityThousandthPercent`、`textDefaultShadowRotateWithShape`、`textDefaultInnerShadowBlurRadiusEmu`、`textDefaultInnerShadowDistanceEmu`、`textDefaultInnerShadowDirectionDegrees`、`textDefaultInnerShadowColorRgb`/`textDefaultInnerShadowColorScheme`、`textDefaultInnerShadowOpacityThousandthPercent` 与 `textDefaultReflectionBlurRadiusEmu`、`textDefaultReflectionDistanceEmu`、`textDefaultReflectionStartOpacityThousandthPercent`、`textDefaultReflectionEndOpacityThousandthPercent`、`textDefaultReflectionDirectionDegrees` 的独立回写，二十四类 owner 只改所属 SlidePart 并二次投影恢复。多 effect chain、reflection→softEdge 之外的其它复杂拓扑、非 full-span/transform/unknown/extension graph 继续 source-owned 或 fail closed。

**本轮 formal text owner 增量：** 八个文字标量（含 `text.fontFamilyEastAsia`、`text.fontFamilyComplexScript` 和 `text.language`）可在声明 `stylePrecedence` 后按 `run` → `paragraph` → `element` → `styleRef` → `layout` → `master` → `theme` → `default` 首个命中 authored 编译到 native text run；普通 run、默认 run、表格和图表文本的显式复杂脚本字体还会直接落到 `a:cs`，并在 imported/source-bound 单叶路径中保持 owner；`text.language` 则落到直接 `a:rPr/@lang`；`design.theme.textStyle` 和 grammar default token 的命中/回落已有去嵌入后二次投影与 review 证据。该 profile 仍不覆盖完整语言/脚本字体回退、宿主字体求值或 source-bound 继承写回。

**仍缺：** 完整 PowerPoint field (`p:fld`) 语义、字段刷新/宿主求值、日期/页码/作者自动字段的完整类型目录、复杂 WordArt transform、除 authored、direct rich-text run 及 strict `defaultText` 已覆盖的 glow、outer shadow、inner shadow、reflection、soft-edge leaves 外的完整文本 effect、`defaultText` 的其它 source-bound effect leaf、所有语言/脚本/字体回退组合、rich text 的未知扩展和复杂列表拓扑；文本容器 bodyPr 当前已覆盖 direct vertical alignment/wrap/inset/columns/column gap-direction/vertical text、rotation、horizontal/vertical overflow、upright、anchorCenter、有限 auto-fit 及 canonical `normalAutoFit` 百分比，继承、显式删除、effect/extension graph 和自动 reflow 仍 source-owned；`run.break` 只承诺已有 line-break inline 的固定拓扑保留，不提供自动换行或段落重排；表格字段 profile 仍只允许修改已有字段的缓存 display text，字段 ID/type、字段关系和复杂字段图继续 source-owned；当前 typed field 只承诺固定可回读值，不伪装成 PowerPoint 自动计算。

**待实现与验收：** 新增字段前先定义静态值、自动值和宿主计算的边界；静态 display 的 source-bound 编辑已由 `PpjTextFieldAuthorsAndProjectsAsTypedRun` 覆盖，field 的 ID/type 变更仍 fail closed；authoring 必须二次导入恢复；第三方 field/WordArt 未识别时保持原文，不能转成普通字符串；文本布局报告必须与实际字体、边距和 AutoFit 证据绑定。

**文本 glow 增量验收：** authored run 与 paragraph default-run 均须在二次导入中恢复 `color`/`radius`/`opacity`；direct rich-text run 的 source-bound profile 还须恢复对应的 RGB/主题色、半径、已有显式 alpha leaf，并在单字段编辑后只改变所属 SlidePart；paragraph `defaultText` 的 strict owner 还须恢复 `textDefaultGlowRadiusEmu`、`textDefaultGlowColorRgb`、`textDefaultGlowColorScheme`、`textDefaultGlowOpacityThousandthPercent`，只编辑直接 `glow/@rad`、`a:srgbClr/@val`、`a:schemeClr/@val` 或已有直接 `a:alpha/@val` 并保留 paragraph index；与既有 outer shadow 同时出现时必须保持直接 effect owner 和顺序。仅单一 `a:glow` 或 `a:glow`→`a:outerShdw` 进入该 profile，超出半径/opacity、缺失显式 alpha 或带 inner/reflection/soft-edge/unknown 等未建模 effect graph 时保持 source-owned 或 fail closed，不把 glow 转成普通文本样式。

**文本 inner shadow 增量验收：** authored run 与 paragraph default-run 均须在二次导入中恢复 `color`/`blur`/`distance`/`angle`/`opacity`；direct rich-text run 的 source-bound profile 还须恢复存在的几何、RGB/主题色和显式 alpha leaf，并在单字段编辑后只改变所属 SlidePart；与 outer shadow 同时出现时必须保持 `a:innerShdw` 的独立 owner 和 canonical 顺序；paragraph `defaultText` 的 strict owner 还须恢复 `textDefaultInnerShadowBlurRadiusEmu`、`textDefaultInnerShadowDistanceEmu` 与 `textDefaultInnerShadowDirectionDegrees`、`textDefaultInnerShadowColorRgb`/`textDefaultInnerShadowColorScheme`、`textDefaultInnerShadowOpacityThousandthPercent`，只编辑直接 `innerShdw/@blurRad`、`innerShdw/@dist`、`innerShdw/@dir`、`a:srgbClr/@val`/`a:schemeClr/@val` 或已有直接 `a:alpha/@val` 并保留 paragraph index；缺失或隐式 alpha 仍 source-owned。仅单一 `a:innerShdw` 或 `a:innerShdw`→`a:outerShdw` 进入该 profile，带 glow、reflection、soft-edge、unknown/extension 等未建模 effect graph 时保持 source-owned 或 fail closed，不把 inner shadow 转成普通文本样式。

**文本 defaultText outer shadow 增量验收：** authored default-run 与 source-bound paragraph `defaultText` owner 均须恢复 outer-shadow `blur`/`distance`/`direction`/`alignment`/RGB/theme color/opacity；focused fixture 的 `textDefaultShadowBlurRadiusEmu`、`textDefaultShadowDistanceEmu`、`textDefaultShadowDirectionDegrees`、`textDefaultShadowAlignment`、`textDefaultShadowColorRgb`、`textDefaultShadowColorScheme`、`textDefaultShadowOpacityThousandthPercent` 与 `textDefaultShadowRotateWithShape` 单字段编辑必须只改变所属 SlidePart、保持直接 `a:outerShdw` 和其它 XML，并在二次投影中恢复新值。仅一个已有且有界的直接 `outerShdw/@blurRad`/`@dist`/`@dir`/`@algn`、一个直接 `a:srgbClr/@val` 或 `a:schemeClr/@val`，以及该颜色下一个已有的直接 `a:alpha/@val` 和一个已有的显式 `outerShdw/@rotWithShape` 进入该 profile；缺失目标属性、多个 outer shadow、缺失 alpha 和其它未建模 effect graph 保持 source-owned 或 fail closed。

**文本 reflection 增量验收：** authored run 与 paragraph default-run 均须在二次导入中恢复 `blur`/`startOpacity`/`endOpacity`/`distance`/`angle`；direct rich-text run 的 source-bound profile 还须恢复存在的五个 native leaf，并在单字段编辑后只改变所属 SlidePart，同时保留 `stPos=0`/`endPos=100000`；paragraph `defaultText` 的 strict owner 还须恢复 `textDefaultReflectionBlurRadiusEmu`、`textDefaultReflectionDistanceEmu`、`textDefaultReflectionStartOpacityThousandthPercent`、`textDefaultReflectionEndOpacityThousandthPercent` 与 `textDefaultReflectionDirectionDegrees`，只编辑直接 `reflection/@blurRad`/`reflection/@dist`/`reflection/@stA`/`reflection/@endA`/`reflection/@dir` 并保留 paragraph index；仅一个 full-span `a:reflection` 或 `a:outerShdw`→`a:reflection` 进入该 profile，超出几何/opacity、非 full-span、transform/unknown/extension 或带其它未建模 effect graph 时保持 source-owned 或 fail closed，不把 reflection 转成普通文本样式。

**文本 soft edge 增量验收：** authored run 与 paragraph default-run 均须在二次导入中恢复 `radius`；direct rich-text run 的 source-bound profile 还须恢复 `textSoftEdgeRadiusEmu`，paragraph `defaultText` 的 strict owner 还须恢复 `textDefaultSoftEdgeRadiusEmu`，两类单字段编辑都只改变所属 SlidePart，并保留已验证的前置 owner/段落 index；仅一个直接 `a:softEdge`、`a:outerShdw`→`a:softEdge` 或 full-span `a:reflection`→`a:softEdge` 进入该 profile，超出半径、非 full-span/transform、glow/inner/多 effect chain、unknown/extension 时保持 source-owned 或 fail closed，不把 soft edge 转成普通文本样式。

**段落 tab stop 增量：** `style.tabStops`（points + `left`/`center`/`right`/`decimal`）和 `style.noTabStops` 已进入 PPJ schema；`PpjParagraphTabStopsAuthorAndReproject` 覆盖 authored native `a:tabLst`、去嵌入投影、source-bound 单字段修改只写 `ppt/slides/slide1.xml` 及二次投影恢复。**文本容器 bodyPr 增量：** 普通文本框、带文本形状和占位符的 direct bodyPr 有界叶子通过 `setTextBodyStyle` 做 source-bound 回写，现覆盖 rotation、horizontal/vertical overflow、upright 以及 canonical `normalAutoFit` 百分比；对应的 `nativeRef.leaves[]` 现在另外暴露 `textBodyNormalAutoFitFontScale` 与 `textBodyNormalAutoFitLineSpacingReduction`，可直接 token-splice 单个 `a:normAutofit` 属性并保留另一个属性；省略字段保留原 native leaf，不把样式删除伪装成成功；`PpjSourceBoundTextBodyStyleEditsTextShapeAndReprojects` 与 `PpjSourceBoundNormalAutoFitLeavesEditAndReproject` 覆盖 bodyPr authored/投影、capability 与 changed-part。自动字段求值、WordArt、继承、显式删除和复杂文本效果仍不在该 profile。
**段落 tab stop 增量：** `style.tabStops`（points + `left`/`center`/`right`/`decimal`）和 `style.noTabStops` 已进入 PPJ schema；`PpjParagraphTabStopsAuthorAndReproject` 覆盖 authored native `a:tabLst`、去嵌入投影、source-bound 单字段修改只写 `ppt/slides/slide1.xml` 及二次投影恢复。**文本容器 bodyPr 增量：** 普通文本框、带文本形状和占位符的 direct bodyPr 有界叶子通过 `setTextBodyStyle` 做 source-bound 回写，现覆盖 rotation、horizontal/vertical overflow、upright、anchorCenter 以及 canonical `normalAutoFit` 百分比；对应的 `nativeRef.leaves[]` 现在另外暴露 `textBodyNormalAutoFitFontScale` 与 `textBodyNormalAutoFitLineSpacingReduction`，可直接 token-splice 单个 `a:normAutofit` 属性并保留另一个属性；省略字段保留原 native leaf，不把样式删除伪装成成功；`PpjSourceBoundTextBodyStyleEditsTextShapeAndReprojects`、`PpjSourceBoundNormalAutoFitLeavesEditAndReproject` 与 `PpjSourceBoundTextBodyAnchorCenterEditsLeafAndReprojects` 覆盖 bodyPr authored/投影、capability 与 changed-part。自动字段求值、WordArt、继承、显式删除和复杂文本效果仍不在该 profile。

### F-04 Shape、Custom Geometry、Connector 和 Group Transform

**连接线 custom site 绑定增量（2026-09-10）：** `from/to: {element, connectionSite}` 用 0..1023 的索引绑定可完整表达的 custom shape `geometry.connectionSites`。编译按原生公式、frame 旋转/镜像和 group 子坐标求解端点，并写入原生 target/index；去嵌入投影保留索引。源 `setConnectorEndpoints` 支持换索引、换目标、解除绑定和转为 frame anchor；语义 geometry/frame 及已支持的 frame leaf 修改会重算依赖端点。绑定目标的 geometry native-leaf 修改明确拒绝，改用语义 geometry；preset/不透明原生绑定保持既有权限边界。最小回归覆盖实际坐标、原生绑定、源 no-op、独立修改、二次投影及非目标 ZIP 保留；custom geometry/connector/preview 专项 130/130 通过，零跳过。预览明确报告限制；未重建 NativeAOT 或验收宿主拖拽，完整 F-04 仍开放。

**优先级：P0；状态：部分完成。**

**自定义路径 viewport 增量（2026-09-10）：** `geometry.paths[].viewport: {width,height}` 保留每条路径独立的坐标宽高；省略继承 geometry.viewBox，任一轴为 0 则使用原生形状坐标默认值。已接通 authored、去嵌入投影和源 paths 覆盖值增改删；投影可重新归并公共 viewBox 与覆盖值，实际原生宽高保持等价。最小实验覆盖不同宽高、单轴/双轴默认、原生上限、原源 no-op 和命令/标志/文字/frame/非目标 ZIP 保留；负数/越界与 mask/clip override 拒绝。路径宽高 native leaf 的数值上限已与原生 int32 最大值对齐，其它 leaf 上限不变；相关 geometry/connector/preview 专项 128/128 通过。预览保留字段诊断，宿主几何外观仍开放。

**自定义形状引用型路径增量（2026-09-10）：** `geometry.paths[].commands[]` 的 move/line/quadratic/cubic 坐标及 arc 半径/角度均可使用数值或原生引用字符串。数值沿用 viewBox/角度转换，字符串保留内建/adjustment/guide 名称及原生单位。已接通 authored、去嵌入投影和源 paths 引用/数值替换；最小实验逐槽回写，并验证 adjustment 20000→30000 的求值变化及路径引用 XML 不变，保留其它 graph/文字/frame/非目标 ZIP。悬空引用、非法求值弧、line/mask/clip 引用路径拒绝；相关 geometry/connector/preview 专项 127/127 通过；默认或不同 path viewport、宿主几何行为仍开放，预览明确报告引用限制。

**自定义形状 adjustment handle 增量（2026-09-10）：** `geometry.adjustmentHandles[]` 以 `kind: xy/polar` 表达两类手柄。XY 包含 X/Y 受控调整项及范围，polar 包含半径/角度受控调整项及范围，两类均有 `position: {x,y}`；数值为局部 points/角度，字符串保留原生引用。已支持 authored、去嵌入投影、源文件中成对范围的增改删和位置修改；源顺序、类型及受控名称保持固定。最小实验核对实际单位、零/省略、两类引用、原源 no-op、非法范围/身份/位置拒绝和 path XML/文字/frame/非目标 ZIP 保留，相关专项 126/126 通过。预览保留手柄诊断；引用型路径和宿主拖拽仍开放。

**自定义形状 connection site 增量（2026-09-10）：** `geometry.connectionSites` 用最多 1024 项有序 `{angle, x, y}` 表达连接点；数值是角度和形状局部 points，字符串保留内建/adjustment/guide 引用。literal-path custom shape 已支持 authored、去嵌入投影和 source-bound 逐槽值修改；空 authored 列表投影为省略。源列表数量保持固定，避免改变连接线使用的下标身份。最小实验覆盖实际单位、引用互换、原源 no-op、path XML/文字/frame/非目标 ZIP 保留，以及越界、悬空引用、列表数量变更和缺权限拒绝；相关 custom geometry/connector/preview 专项 125/125 通过。handle、引用型 path、连接点绑定的高层 authored 语法和宿主拖拽仍开放；预览给出明确限制。

**自定义路径 extrusionOk 增量（2026-09-10）：** `geometry.paths[].extrusionOk` 保留原生路径“允许拉伸”的可选布尔值；true、false 和省略分别往返。现有 paths 编辑权限支持增改删，修改坐标时保留该标志，修复 PPJ 投影及重写路径时的属性丢失。最小实验核对原源 no-op、非目标 ZIP 和其余 geometry/文字/frame 保留、非法类型拒绝；custom geometry/authored-preview 专项 67/67 通过。预览明确报告字段限制，3-D 深度/材质与宿主外观仍开放。

**自定义形状 adjustment 增量（2026-09-10）：** `kind: custom` 的 `geometry.adjustments` 使用最多 256 项有序 `{name, formula}`，先于 guides 求值，共用名称与引用检查；preset 继续使用整数数组。已支持 authored、去嵌入投影和 source-bound 列表增改删，空列表投影为省略。最小实验验证 adjustment → guide → 文字矩形链路的 10→20 pt 实际求值、原源 no-op、协调删除、跨列表重名/前向/悬空引用拒绝，以及 paths/文字/frame/非目标 ZIP 保留；custom/preset geometry 与 authored-preview 专项 68/68 通过。handle/site、引用型 path 坐标和宿主交互仍开放，预览保留公式限制诊断。

**自定义形状公式 guide 增量（2026-09-10）：** `geometry.guides` 承载有序 `{name, formula}` 列表，复用原生有界公式图；文字矩形可引用声明的 guide，顺序和引用通过 authored/去嵌入投影保留。`setGeometry` 支持列表增改删，删除仍被引用的 guide 会拒绝，可同时删除依赖矩形完成合法清理；空列表投影为省略。最小实验验证 `w/10→w/5` 的实际求值、原源 no-op、path XML/文字/frame 和非目标 ZIP 保留，相关专项 65/65 通过。handle/site 和引用型 path 坐标的 PPJ 整体表达仍开放；预览保留明确公式/布局诊断。

**自定义形状文字矩形增量（2026-09-10）：** `geometry.textRectangle` 以 `left/top/right/bottom` 表达形状局部点值、原生内建引用或混合边界；省略保留原生默认。literal custom shape 的 authored/去嵌入投影和 `setGeometry` 增改删已接通，不再因为存在 text rectangle 就丢失可编辑形状身份。`PpjCustomGeometryTextRectangleTests` 核对原源 no-op、修改/删除/再新增、路径 XML/文字/frame 和非目标 ZIP 保留，相关专项 64/64 通过；无效边界、未知引用和 mask/clip owner 拒绝。custom guide/handle 图的完整 PPJ 表达及宿主布局仍开放；预览四边均保留明确诊断。

**连接线弯折调整增量（2026-09-10）：** `bendAdjustment` 保留 elbow/curved 三段路径的直接 `adj1` 原生整数值，支持 signed int32、显式零与省略；`setConnectorType` 可独立修改/删除，换为 straight 清理旧值，显式冲突拒绝。最小原源实验覆盖 `25000→0`、负值、删除、native guide、再投影和仅目标 SlidePart 改动，相关回归 69/69 通过，零跳过。计算公式和端点归一化无法保留弯折轴的旋转源仍 opaque；内部预览对非默认弯折明确给出 `preview.scene.paint.connector-bend`，不代画中点路线。完整路由及宿主外观仍开放。

**连接线类型增量（2026-09-10）：** `connectorType` 通过 `setConnectorType` 支持 source-bound 的 straight/elbow/curved 切换，复用原生 canonical geometry 替换，保留端点、绑定、箭头和线样式。相关 Connector/preview 回归 67/67 通过；最小回归按 straight→elbow→curved→straight 核对实际 `prstGeom`、每个源的 no-op、SlidePart-only 差异与二次投影；非法类型及修改过的权限证据继续拒绝。此字段仍必填，不代表自动避障或任意 bend/guide 编辑。

**连接线箭头尺寸增量（2026-09-10）：** PPJ 新增 `startArrowWidth/startArrowLength/endArrowWidth/endArrowLength`，各取 `sm/med/lg`，省略保持原生默认且再投影不补默认值。`setConnectorArrows` 支持逐个修改或删除尺寸属性；整端箭头删除时清掉该端尺寸，显式尺寸与箭头删除冲突则拒绝。`PpjConnectorObjectAnchorTests` 覆盖创建、去嵌入投影、同一原源逐字段增改删、no-op 字节相同、绑定和非目标 ZIP 保留；相关专项 68/68 通过，零跳过。生产预览仍以 `preview.connector.limited` 标记部分支持，未重建 NativeAOT 或做宿主外观验收；F-04 整体继续开放。

**连接线箭头增量（2026-09-10）：** `startArrow/endArrow` 通过 `setConnectorArrows` 支持 source-bound 增改删；省略或 `none` 删除该端及其尺寸，换形状保留原宽高，另一端与端点绑定保持不变。PPJ `open` 与原生 `arrow` 双向转换，修复 authored/去嵌入投影的枚举断层。相关回归 67/67 通过；`PpjConnectorObjectAnchorTests` 用同一原源分别修改、删除和新增箭头，核对 native 尺寸/绑定、SlidePart-only 和二次投影；opaque owner 不因此取得编辑权限。

**有符号端点增量（2026-09-10）：** `connector.from/to.x/y` 及对象锚点换算结果保留负值，旋转跨组后的端点可落在 childFrame 原点之前；写入、导入、源绑定端点和 frame 修改使用同一坐标范围。极端越界在运算前拒绝，原生异常对象保持 opaque；若其 frame 超过 PPJ 上限，则投影明确失败。最小实验沿用 `PpjConnectorObjectAnchorTests`，相关 Connector/preview 回归 65/65 通过，核对负端点的真实 XML、原源 no-op、修改后再投影和非目标 ZIP 成员保留；宿主范围兼容性与正式预览接入仍独立验证。

**对象锚点增量（2026-09-10）：** `connector.from/to` 的 `{element, anchor}` 现在决定实际端点；`top/right/bottom/left/center` 按目标 frame 解析，`auto` 在 slide 空间选最短侧中点对，同距按起点再终点的 top/right/bottom/left 顺序。旋转、翻转、group childFrame 和组件展开保持各自坐标空间。原生扩展保留目标和原始 anchor，去掉嵌入 PPJ 后仍可投影回来；source-bound 的 `setConnectorEndpoints` 支持改锚点、换目标、改为坐标以解除绑定，目标及祖先 frame/childFrame 的语义和已颁发 native leaf 修改同步更新端点。`PpjConnectorObjectAnchorTests` 与相关回归共 69/69 通过，覆盖原源 no-op、编辑后再投影和只改目标 SlidePart。原有 native connection-site 绑定保留独立权限；不接受 opaque/connector 目标或组件实例虚拟端口；后续有符号端点增量已解除负局部坐标限制（见下）。此处是编译库证据，未重建 NativeAOT、未验证 PowerPoint 拖拽吸附，正式预览接入和完整路由仍待完成。

**当前进度：** PPJ 有 178 个 preset geometry、调整值、custom path、arcTo、frame rotation/flip、straight/elbow/curved connector、组和组件；本轮新增独立 `line` literal path 及 Kimi `points/viewBox/curve` authored lowering 的 authored/投影 profile，并覆盖 bent connector 等有限拓扑。对无 guide/handle/connection-site/text-rectangle 的 literal custom path，source-bound 现在可在保持 custom geometry owner 的情况下改写 paths，只改目标 SlidePart 并二次投影恢复；对已识别 custom geometry 的 `a:avLst`，本轮再开放 `val N` 调整项的 `customGeometryAdjustment` native leaf；对已识别 preset shape 的完整 `a:avLst`，新增逐槽 `presetGeometryAdjustment` native leaf；对结构合法但 partial 或含计算公式的 preset `avLst`，shape 继续保持 source-owned，同时只为独立 literal `val N` sibling 颁发同一 native leaf，公共 `geometry.adjustments` 不暴露不完整向量；对严格 image-fill shape 的 partial/formula custom geometry 也保留 direct `a:avLst`，只为其中独立 literal sibling 颁发 `customGeometryAdjustment`，其余公式和 path graph 保持 source-owned；这些叶子均可 token-splice 修改数值而不重写路径、guide 名称或拓扑；普通 group 也会在 source binding 可证明时暴露外层 `off/ext`、显式 `rot/flipH/flipV` 以及局部 `chOff/chExt` 的 `leftEmu/topEmu/widthEmu/heightEmu/rotationDegrees/flipHorizontal/flipVertical/childLeftEmu/childTopEmu/childWidthEmu/childHeightEmu` 叶子，PPJ 以可选 `childFrame` 保留子坐标矩形；编辑只改目标 SlidePart，保留子树并由二次投影恢复外层/子空间值；复杂 guide/formula/handle、多路径拓扑、子空间联动和自动 descendant rescale 仍不授予 capability。

**仍缺：** 高层 freeCurve authored sugar 的无损保留、完整 DrawingML guide/formula/handle/connection site 语义、preset shape partial/formula/extension guide 图以及 image-fill custom geometry partial/formula graph 的整体语义编辑（目前只允许其中独立 literal sibling 的 native leaf）、group `chOff/chExt` 变化对 descendants 的自动重算/继承联动、多路径闭合和拓扑修改、3D/bevel/soft-edge、完整 connector routing 和 source-bound 复杂路径修改。

**待实现与验收：** K-01 的有界 literal/points profile 已闭合；`SourceBoundLiteralCustomGeometryPathEditChangesOnlySlideAndReprojects` 证明 literal path 的 source-bound changed-part/二次投影闭环，`PpjSourceBoundLiteralCustomGeometryAdjustmentLeafEditsAndReprojects` 证明 custom geometry `val N` 调整叶子的 SlidePart-only token splice、Open XML 校验和二次投影；新增 preset shape 的 `presetGeometryAdjustment` 也沿同一 proof 规则，只接受已知 preset 和结构合法的 ordered guide list，完整列表可作为 PPJ geometry，partial/formula 列表只为 literal sibling 发 native leaf，并保持非目标 XML；`PpjSourceBoundPartialPresetGeometryLiteralSiblingEditsAndReprojects` 覆盖 partial list、目标 token splice、其余 guide 保留、Open XML 校验、二次投影和非 literal sibling 不发 leaf。严格 image-fill custom geometry 的同类路径由 `SourceBoundPartialCustomGeometryLiteralSiblingEditsAndReprojects` 覆盖：不支持的 sibling formula 保留，literal target 只改所属 SlidePart 并通过二次投影恢复。`PpjGroupReadingOrderAuthorsAndReordersLocalShapeTree` 现在同时覆盖普通 group 外层 frame leaf、`childFrame`/`chOff` 单字段回写、外层 `off` 不变、SlidePart-only footprint 和二次投影。后续继续把非 `val` 的 guide formula 整体编辑、partial/extension guide 图整体编辑、child-space descendant rescale/继承、guide/handle、多路径拓扑保持 fail-closed；复杂 group 只提供已证明的 frame、childFrame、text 或子元素局部叶子。

**本轮新增窄闭环：** `customGeometryGuide` 只对应 source-bound custom geometry 的一个直接、有序、唯一命名且 `fmla="val N"` 的 `a:gdLst/a:gd` guide 数值。它只 token-splice 该 guide 的公式数值，保留公式图、handle、connection-site、text rectangle、path 和其它拓扑；计算公式、重复名称、扩展、异常子节点与不完整图不发 capability。`PpjSourceBoundLiteralCustomGeometryGuideLeafEditsAndReprojects` 覆盖 authored → 去嵌入投影 → 单字段编辑 → SlidePart-only → 二次投影。

**本轮新增窄闭环：** `customGeometryGuideFormula` 绑定 source-bound custom geometry 的一个直接、有序、唯一命名且 canonical 的非 `val N` `a:gdLst/a:gd/@fmla` 公式；编辑只替换该公式字符串 token，并用现有有界公式图重新校验，guide 引用、其它 guide、handle、connection-site、text rectangle、path 和其它拓扑保持 source-owned。`PpjSourceBoundCustomGeometryGuideFormulaEditsAndReprojects` 覆盖一个公式 guide 的 authored → 去嵌入投影 → 单 token 编辑 → SlidePart-only → Open XML → 二次投影；不支持的公式、非 canonical 空白、重复名和复杂图仍 fail closed。

**本轮新增窄闭环：** `customGeometryAdjustmentFormula` 绑定 fully recognized custom geometry 的一个直接、有序、唯一拓扑中的非 `val N` `a:avLst/a:gd/@fmla` 调整公式；编辑只替换该 adjustment 的公式 token，并用现有有界公式图重新校验，guide 引用、其它 adjustment/guide、handle、connection-site、text rectangle、path 和其它拓扑保持 source-owned。`PpjSourceBoundCustomGeometryAdjustmentFormulaEditsAndReprojects` 覆盖 authored → 去嵌入投影 → 单 token 编辑 → SlidePart-only → Open XML → 二次投影；partial image-fill fallback、非 canonical 空白、未知公式和复杂图仍 fail closed。

**本轮新增窄闭环：** `customGeometryPathLineToX` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:lnTo/a:pt/@x` canonical signed coordinate；只 token-splice 选中 line-to 点的 `@x`，保留配对 `@y`、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathLineToXEditsAndReprojects` 覆盖两个 line-to leaf、单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryPathLineToY` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:lnTo/a:pt/@y` canonical signed coordinate；只 token-splice 选中 line-to 点的 `@y`，保留配对 `@x`、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathLineToYEditsAndReprojects` 覆盖两个 line-to leaf、单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryPathMoveToX` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:moveTo/a:pt/@x` canonical signed coordinate；只 token-splice 选中 move-to 点的 `@x`，保留配对 `@y`、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathMoveToXEditsAndReprojects` 覆盖单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryPathMoveToY` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:moveTo/a:pt/@y` canonical signed coordinate；只 token-splice 选中 move-to 点的 `@y`，保留配对 `@x`、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathMoveToYEditsAndReprojects` 覆盖单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryPathQuadraticEndX` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:quadBezTo` 第二个 `a:pt/@x` canonical signed quadratic end-point coordinate；只 token-splice 选中终点的 `@x`，保留 control point、终点 `@y`、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathQuadraticEndXEditsAndReprojects` 覆盖单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryPathQuadraticEndY` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:quadBezTo` 第二个 `a:pt/@y` canonical signed quadratic end-point coordinate；只 token-splice 选中终点的 `@y`，保留 control point、终点 `@x`、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathQuadraticEndYEditsAndReprojects` 覆盖单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryPathQuadraticControlX` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:quadBezTo` 第一个 `a:pt/@x` canonical signed quadratic control-point coordinate；只 token-splice 选中控制点的 `@x`，保留控制点 `@y`、终点、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathQuadraticControlXEditsAndReprojects` 覆盖单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryPathQuadraticControlY` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:quadBezTo` 第一个 `a:pt/@y` canonical signed quadratic control-point coordinate；只 token-splice 选中控制点的 `@y`，保留控制点 `@x`、终点、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathQuadraticControlYEditsAndReprojects` 覆盖单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryPathCubicEndX` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:cubicBezTo` 第三个 `a:pt/@x` canonical signed cubic end-point coordinate；只 token-splice 选中终点的 `@x`，保留两个 control point、终点 `@y`、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathCubicEndXEditsAndReprojects` 覆盖单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryPathCubicEndY` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:cubicBezTo` 第三个 `a:pt/@y` canonical signed cubic end-point coordinate；只 token-splice 选中终点的 `@y`，保留两个 control point、终点 `@x`、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathCubicEndYEditsAndReprojects` 覆盖单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryPathCubicControl1X` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:cubicBezTo` 第一个 `a:pt/@x` canonical signed cubic first-control-point coordinate；只 token-splice 选中第一个控制点的 `@x`，保留该点 `@y`、第二个 control point、终点、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathCubicControl1XEditsAndReprojects` 覆盖单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryPathCubicControl1Y` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:cubicBezTo` 第一个 `a:pt/@y` canonical signed cubic first-control-point coordinate；只 token-splice 选中第一个控制点的 `@y`，保留该点 `@x`、第二个 control point、终点、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathCubicControl1YEditsAndReprojects` 覆盖单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryPathCubicControl2X` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:cubicBezTo` 第二个 `a:pt/@x` canonical signed cubic second-control-point coordinate；只 token-splice 选中第二个控制点的 `@x`，保留该点 `@y`、第一个 control point、终点、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathCubicControl2XEditsAndReprojects` 覆盖单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryPathCubicControl2Y` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:cubicBezTo` 第二个 `a:pt/@y` canonical signed cubic second-control-point coordinate；只 token-splice 选中第二个控制点的 `@y`，保留该点 `@x`、第一个 control point、终点、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、越界或 unsupported command 继续 source-owned；`PpjSourceBoundCustomGeometryPathCubicControl2YEditsAndReprojects` 覆盖单字段、SlidePart-only、Open XML 校验和二次投影。

**本轮新增窄闭环：** `customGeometryConnectionSiteAngle60000` 对应 source-bound custom geometry 中一个直接、有序 `a:cxnLst/a:cxn/@ang` 的 canonical bounded integer；只 token-splice 角度，保留 position、guide/formula、handle、path 和 site 拓扑。`PpjSourceBoundLiteralCustomGeometryConnectionSiteAngleLeafEditsAndReprojects` 覆盖单字段和二次投影；公式化、异常或 stale 绑定 fail closed。

**本轮新增窄闭环：** `customGeometryConnectionSiteXEmu` 对应一个直接、有序 `a:cxnLst/a:cxn/a:pos/@x` 的 canonical 非负整数（shape-local EMU）；只 token-splice 选中的 x，保留 y、angle、引用型 sibling、guide/formula、handle、path 和 site 拓扑。`PpjSourceBoundLiteralCustomGeometryConnectionSiteXLeafEditsAndReprojects` 覆盖单字段、二次投影、引用型 sibling、stale 与越界 fail closed。

**本轮新增窄闭环：** `customGeometryConnectionSiteYEmu` 对应一个直接、有序 `a:cxnLst/a:cxn/a:pos/@y` 的 canonical 非负整数（shape-local EMU）；只 token-splice 选中的 y，保留 x、angle、引用型 sibling、guide/formula、handle、path 和 site 拓扑。`PpjSourceBoundLiteralCustomGeometryConnectionSiteYLeafEditsAndReprojects` 覆盖单字段、二次投影、引用型 sibling、stale 与越界 fail closed。

**本轮新增窄闭环：** `customGeometryAdjustmentHandleXEmu` 对应一个直接、有序 `a:ahXY/a:pos/@x` 的 canonical 非负整数（shape-local EMU）；只 token-splice 选中的 handle x，保留 `gdRef` identity、min/max、y、handle kind/order、polar sibling、guide/formula、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryAdjustmentHandleXLeafEditsAndReprojects` 覆盖单字段、二次投影、identity 保留、stale 与越界 fail closed。

**本轮新增窄闭环：** `customGeometryAdjustmentHandleYEmu` 对应一个直接、有序 `a:ahXY/a:pos/@y` 的 canonical 非负整数（shape-local EMU）；只 token-splice 选中的 handle y，保留 `gdRef` identity、min/max、x、handle kind/order、polar sibling、guide/formula、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryAdjustmentHandleYLeafEditsAndReprojects` 覆盖单字段、二次投影、identity 保留、stale 与越界 fail closed。

**本轮新增窄闭环：** `customGeometryAdjustmentHandleMinXEmu` 对应一个直接、有序 `a:ahXY/@minX` 的 canonical 非负整数（shape-local EMU）；只 token-splice 选中的 minX，保留 `gdRef` identity、maxX、y bounds、position、handle kind/order、polar sibling、guide/formula、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryAdjustmentHandleMinXLeafEditsAndReprojects` 覆盖单字段、有效范围替换、二次投影、identity 保留、stale 与越界 fail closed。

**本轮新增窄闭环：** `customGeometryAdjustmentHandleMaxXEmu` 对应一个直接、有序 `a:ahXY/@maxX` 的 canonical 非负整数（shape-local EMU）；只 token-splice 选中的 maxX，保留 `gdRef` identity、minX、y bounds、position、handle kind/order、polar sibling、guide/formula、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryAdjustmentHandleMaxXLeafEditsAndReprojects` 覆盖单字段、有效范围替换、二次投影、identity 保留、stale 与越界 fail closed。

**本轮新增窄闭环：** `customGeometryAdjustmentHandleMinYEmu` 对应一个直接、有序 `a:ahXY/@minY` 的 canonical 非负整数（shape-local EMU）；只 token-splice 选中的 minY，保留 `gdRef` identity、x bounds、maxY、position、handle kind/order、polar sibling、guide/formula、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryAdjustmentHandleMinYLeafEditsAndReprojects` 覆盖单字段、有效范围替换、二次投影、identity 保留、stale 与越界 fail closed。

**本轮新增窄闭环：** `customGeometryAdjustmentHandleMaxYEmu` 对应一个直接、有序 `a:ahXY/@maxY` 的 canonical 非负整数（shape-local EMU）；只 token-splice 选中的 maxY，保留 `gdRef` identity、x bounds、minY、position、handle kind/order、polar sibling、guide/formula、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryAdjustmentHandleMaxYLeafEditsAndReprojects` 覆盖单字段、有效范围替换、二次投影、identity 保留、stale 与越界 fail closed。

**本轮新增窄闭环：** `customGeometryAdjustmentHandlePolarMinRadiusEmu` 对应一个直接、有序 `a:ahPolar/@minR` 的 canonical 非负 DrawingML 整数（shape-local EMU）；只 token-splice 选中的 minR，保留 `gdRefR`/`gdRefAng`、maxR、angle bounds、position、handle kind/order、XY sibling、guide/formula、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryAdjustmentHandlePolarMinRadiusLeafEditsAndReprojects` 覆盖单字段、有效范围替换、二次投影、identity 保留、stale 与非法输入 fail closed。

**本轮新增窄闭环：** `customGeometryAdjustmentHandlePolarMaxRadiusEmu` 对应一个直接、有序 `a:ahPolar/@maxR` 的 canonical 非负 DrawingML 整数（shape-local EMU）；只 token-splice 选中的 maxR，保留 `gdRefR`/`gdRefAng`、minR、angle bounds、position、handle kind/order、XY sibling、guide/formula、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryAdjustmentHandlePolarMaxRadiusLeafEditsAndReprojects` 覆盖单字段、有效范围替换、二次投影、identity 保留、stale 与非法输入 fail closed。

**本轮新增窄闭环：** `customGeometryAdjustmentHandlePolarMinAngle60000` 对应一个直接、有序 `a:ahPolar/@minAng` 的 canonical 1/60000 度整数（±360°）；只 token-splice 选中的 minAng，保留 `gdRefR`/`gdRefAng`、radial bounds、maxAng、position、handle kind/order、XY sibling、guide/formula、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryAdjustmentHandlePolarMinAngleLeafEditsAndReprojects` 覆盖单字段、有效范围替换、二次投影、identity 保留、stale 与越界 fail closed。

**本轮新增窄闭环：** `customGeometryAdjustmentHandlePolarMaxAngle60000` 对应一个直接、有序 `a:ahPolar/@maxAng` 的 canonical 1/60000 度整数（±360°）；只 token-splice 选中的 maxAng，保留 `gdRefR`/`gdRefAng`、radial bounds、minAng、position、handle kind/order、XY sibling、guide/formula、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryAdjustmentHandlePolarMaxAngleLeafEditsAndReprojects` 覆盖单字段、有效范围替换、二次投影、identity 保留、stale 与越界 fail closed。

**本轮新增窄闭环：** `customGeometryAdjustmentHandlePolarXEmu` 对应一个直接、有序 `a:ahPolar/a:pos/@x` 的 canonical 非负整数（shape-local EMU）；只 token-splice 选中的 polar position x，保留 `gdRefR`/`gdRefAng`、radial/angular bounds、y、handle kind/order、XY sibling、guide/formula、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryAdjustmentHandlePolarXLeafEditsAndReprojects` 覆盖单字段、in-frame 替换、二次投影、identity 保留、stale 与越界 fail closed。

**本轮新增窄闭环：** `customGeometryAdjustmentHandlePolarYEmu` 对应一个直接、有序 `a:ahPolar/a:pos/@y` 的 canonical 非负整数（shape-local EMU）；只 token-splice 选中的 polar position y，保留 `gdRefR`/`gdRefAng`、radial/angular bounds、x、handle kind/order、XY sibling、guide/formula、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryAdjustmentHandlePolarYLeafEditsAndReprojects` 覆盖单字段、in-frame 替换、二次投影、identity 保留、stale 与越界 fail closed。

**本轮新增窄闭环：** `customGeometryTextRectangleLeftEmu` 对应一个 recognized custom geometry 的有序 `a:rect/@l` literal left edge；既支持 canonical direct numeric token，也支持 OfficeKit authored 的 `l="officeKitTextLeft"` + scaled `gd` profile。四边必须为 in-frame canonical 非负整数且保持 left<right/top<bottom；只 token-splice 左边界，保留 direct/private 表示、top/right/bottom、guide、handle、connection-site、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryTextRectangleLeftLeafEditsAndReprojects` 覆盖 authored/private profile、单字段、二次投影、representation 保留、stale 与越界/无序 fail closed。

**本轮新增窄闭环：** `customGeometryTextRectangleTopEmu` 对应同一 recognized custom geometry 的有序 `a:rect/@t` literal top edge；既支持 canonical direct numeric token，也支持 OfficeKit authored 的 `t="officeKitTextTop"` + scaled `gd` profile。四边必须为 in-frame canonical 非负整数且保持 left<right/top<bottom；只 token-splice 上边界，保留 direct/private 表示、left/right/bottom、guide、handle、connection-site、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryTextRectangleTopLeafEditsAndReprojects` 覆盖 authored/private profile、direct representation、单字段、二次投影、stale 与越界/无序 fail closed。

**本轮新增窄闭环：** `customGeometryTextRectangleRightEmu` 对应同一 recognized custom geometry 的有序 `a:rect/@r` literal right edge；既支持 canonical direct numeric token，也支持 OfficeKit authored 的 `r="officeKitTextRight"` + scaled `gd` profile。四边必须为 in-frame canonical 非负整数且保持 left<right/top<bottom；只 token-splice 右边界，保留 direct/private 表示、left/top/bottom、guide、handle、connection-site、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryTextRectangleRightLeafEditsAndReprojects` 覆盖 authored/private profile、direct representation、单字段、二次投影、stale 与越界/无序 fail closed。

**本轮新增窄闭环：** `customGeometryTextRectangleBottomEmu` 对应同一 recognized custom geometry 的有序 `a:rect/@b` literal bottom edge；既支持 canonical direct numeric token，也支持 OfficeKit authored 的 `b="officeKitTextBottom"` + scaled `gd` profile。四边必须为 in-frame canonical 非负整数且保持 left<right/top<bottom；只 token-splice 下边界，保留 direct/private 表示、left/top/right、guide、handle、connection-site、path 和其它拓扑。`PpjSourceBoundLiteralCustomGeometryTextRectangleBottomLeafEditsAndReprojects` 覆盖 authored/private profile、direct representation、单字段、二次投影、stale 与越界/无序 fail closed。

**本轮新增窄闭环：** `customGeometryPathWidth` 对应 recognized custom geometry 中按 path 顺序排列的直接正整数 `a:path/@w`，单位是 DrawingML path-coordinate units，不误称 EMU；只 token-splice 选中 path 的 `@w`，保留 `@h`、fill/stroke/extrusionOk、commands、sibling paths 和其它拓扑。省略/零值、malformed、extension-bearing 或 unsupported path 继续 source-owned；`PpjSourceBoundLiteralCustomGeometryPathWidthLeafEditsAndReprojects` 覆盖 authored/imported、单字段、SlidePart-only、二次投影、stale 与非法值 fail closed。

**本轮新增窄闭环：** `customGeometryPathHeight` 对应 recognized custom geometry 中按 path 顺序排列的直接正整数 `a:path/@h`，单位同样是 DrawingML path-coordinate units；只 token-splice 选中 path 的 `@h`，保留 `@w`、fill/stroke/extrusionOk、commands、sibling paths 和其它拓扑。省略/零值、malformed、extension-bearing 或 unsupported path 继续 source-owned；`PpjSourceBoundLiteralCustomGeometryPathHeightLeafEditsAndReprojects` 覆盖 authored/imported、单字段、SlidePart-only、二次投影、stale 与非法值 fail closed。

**本轮新增窄闭环：** `customGeometryPathArcWidthRadius` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:arcTo/@wR` positive canonical literal width radius；只 token-splice 选中 arc 的 `@wR`，保留 `@hR`、start/sweep angle、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、非正或 unsupported arc state 继续 source-owned；`PpjSourceBoundLiteralCustomGeometryPathArcWidthRadiusLeafEditsAndReprojects` 覆盖单字段、SlidePart-only、二次投影和非法值 fail closed。

**本轮新增窄闭环：** `customGeometryPathArcHeightRadius` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:arcTo/@hR` positive canonical literal height radius；只 token-splice 选中 arc 的 `@hR`，保留 `@wR`、start/sweep angle、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、非正或 unsupported arc state 继续 source-owned；`PpjSourceBoundLiteralCustomGeometryPathArcHeightRadiusLeafEditsAndReprojects` 覆盖单字段、SlidePart-only、二次投影和非法值 fail closed。

**本轮新增窄闭环：** `customGeometryPathArcStartAngle60000` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:arcTo/@stAng` canonical signed 1/60000-degree literal angle；只 token-splice 选中 arc 的 `@stAng`，保留 `@wR/@hR`、sweep angle、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、超出一整圈或 unsupported arc state 继续 source-owned；`PpjSourceBoundLiteralCustomGeometryPathArcStartAngleLeafEditsAndReprojects` 覆盖单字段、SlidePart-only、二次投影和非法值 fail closed。

**本轮新增窄闭环：** `customGeometryPathArcSweepAngle60000` 对应 recognized custom geometry 中按 path/command 顺序定位的直接 `a:arcTo/@swAng` canonical signed non-zero 1/60000-degree literal sweep angle；只 token-splice 选中 arc 的 `@swAng`，保留 `@wR/@hR`、start angle、path 属性、其它 command、sibling path 和其它拓扑。reference-backed、malformed、extension-bearing、零值、超出一整圈或 unsupported arc state 继续 source-owned；`PpjSourceBoundLiteralCustomGeometryPathArcSweepAngleLeafEditsAndReprojects` 覆盖单字段、SlidePart-only、二次投影和非法值 fail closed。

**本轮新增窄闭环：** `customGeometryPathFill` 对应 recognized custom geometry 中按 path 顺序排列的直接 `a:path/@fill`，仅接受 `norm`/`none` 并在 PPJ 中表现为 boolean；只 token-splice 选中 path 的 `@fill`，保留 `@w/@h`、stroke/extrusionOk、commands、sibling paths 和其它拓扑。省略、unsupported、malformed 或 extension-bearing fill 继续 source-owned；`PpjSourceBoundLiteralCustomGeometryPathFillLeafEditsAndReprojects` 覆盖 authored/imported、单字段、SlidePart-only、二次投影、stale 与非法值 fail closed。

**本轮新增窄闭环：** `customGeometryPathStroke` 对应 recognized custom geometry 中按 path 顺序排列的显式 `a:path/@stroke` boolean；只 token-splice 选中 path 的 `@stroke`，保留 `@w/@h`、fill/extrusionOk、commands、sibling paths 和其它拓扑。省略、malformed、unsupported 或 extension-bearing stroke 继续 source-owned；`PpjSourceBoundLiteralCustomGeometryPathStrokeLeafEditsAndReprojects` 覆盖 authored/imported、单字段、SlidePart-only、二次投影、stale、非法与 unchanged 请求 fail closed。

**本轮新增窄闭环：** `customGeometryPathExtrusionAllowed` 对应 recognized custom geometry 中按 path 顺序排列的显式 `a:path/@extrusionOk` boolean；只 token-splice 选中 path 的 `@extrusionOk`，保留 `@w/@h`、fill/stroke、commands、sibling paths 和其它拓扑。省略、malformed、unsupported 或 extension-bearing extrusion 继续 source-owned；`PpjSourceBoundLiteralCustomGeometryPathExtrusionAllowedLeafEditsAndReprojects` 覆盖 authored/imported、单字段、SlidePart-only、二次投影、stale、非法与 unchanged 请求 fail closed。

### F-05 图片、背景、Fill、Crop、Mask 和 Effects

**优先级：P0；状态：部分完成。**

**本轮增量：** F-05 的复合透明度补上 source-bound 普通无可见文字 shape 的窄 owner profile：只有 fill/gradient/image/outline/shadow alpha 可证明一致时才投影 `compositing.opacity`，修改只落所属 SlidePart 并二次投影恢复；同时补上 source-free authored 的 shape/image/line/picture `glow`、`innerShadow`、`reflection` 与 `softEdge` 字段，分别使用有限 color/radius/opacity、color/blur/distance/angle/opacity、blur/startOpacity/endOpacity/distance/angle 或 radius 直接写入 `a:effectLst`，并保持 glow、inner shadow、outer shadow、reflection、soft edge 的独立顺序；本轮又补上 imported/source-bound 普通 shape、line、picture 的直接 glow 或 glow→outer-shadow profile、直接 inner shadow 或 inner-shadow→outer-shadow profile，以及 direct soft edge、outer-shadow→soft-edge、full-span reflection→soft-edge profile，按颜色拓扑、几何标量和显式 alpha 或 radius 颁发 shape/image native leaves，只做所属 SlidePart token splice；其余 imported/source-bound effect graph 和混合 owner 继续 source-owned。

**当前进度：** 支持图片 asset、SVG/raster fallback、页面/母版/布局背景、image fill、cover/contain/stretch/tile/none、crop、预设/自定义 mask、渐变、透明度、边框、阴影；本轮增加受限 `compositing` 声明和 normal opacity authored 编译，并补上 source-bound solid slide background 的 RGB、`color` grammar token 与 opacity 单页闭环，直接嵌入图片背景的 crop/opacity 与单 owner 资源替换也已形成 SlidePart、SlideMasterPart、SlideLayoutPart 及各自关系/媒体闭包和二次投影证据；部分 source-bound 图片 frame、SVG 和 image-filled custom geometry 已有证据。recognized picture 的默认矩形 mask 会规范化为 `imageMaskPreset: "rect"` 叶子；对受支持预设（包括完整、按固定 guide 顺序的 adjustments），可用 `image.mask.preset` 与 `image.mask.adjustments` 在既有 picture `a:prstGeom` 内做 source-bound identity/参数变更，并由 `PpjSourceBoundImageMaskPresetAndCustomIdentityChangesAndReprojects` 证明无 adjustments 与不同 guide arity 的调整预设切换、预设↔literal custom 双向转换都只改目标 SlidePart、再投影恢复。现在完整的 `a:avLst/a:gd fmla="val N"` 还会为每个调整槽颁发 `imageMaskAdjustment` native leaf；编辑只 token-splice 对应 `fmla`，保留 preset、guide 顺序、mask 拓扑和图片关系。即使一个 mask 因 partial/formula adjustment 退回 opaque，只要 preset、`avLst` 和每个 `gd` 仍是简单直接属性图，也会为独立 literal `val N` sibling 颁发同一叶子；`PpjSourceBoundImageMaskPresetAndCustomIdentityChangesAndReprojects` 已覆盖完整和 partial 两种路径，均只改目标 SlidePart 并二次投影恢复。recognized picture 的直接 RGB/主题色边框与单一 outer shadow 已加入 `setImageEffects`，可在已有 `a:spPr` 内 source-bound 增改/清除；普通非文本框形状和线条的单一 outer shadow 也通过 `setShapeEffects` 具备同样的 SlidePart-only source-bound 闭环，并以 changed-part 和二次投影回归证明；blend/isolation/clip closure 仍明确 fail closed。

**仍缺：** 多层背景、通用 blend/isolation/clip stack、imported/source-bound shape/image 的 3-D effect，以及 direct text-run glow/inner shadow/reflection/soft edge 之外的完整文本 effect、任意 SVG filter、preset mask 中非简单直接 `avLst`、非 literal 目标 guide、扩展/子节点/unknown geometry 的编辑（这些不会伪装成 `imageMaskAdjustment`）、带公式/扩展/复杂拓扑的 source-bound 自定义 mask 路径修改、fallback pair 的新增/删除和完整替换；无 guide/handle 的 literal custom mask、direct shape/image glow、inner shadow、reflection 和 soft edge 已有窄路径，复杂 effect graph 仍按 source-owned/fail-closed 处理。

**待实现与验收：** K-04 的 normal opacity、connector `compositing.opacity` 的单 owner authored/规范 projection、icon/placeholder 的 shape-backed authored lowering、literal custom-mask path、recognized picture preset identity（含完整 adjustments）、完整及简单 partial/formula mask 的 literal sibling leaf、border/shadow、ordinary-shape shadow、source-free authored shape/image/line/picture `glow`/`innerShadow`/`reflection`/`softEdge`、imported/source-bound 普通 shape/line/picture direct glow、direct inner shadow、reflection 与 soft edge、source-bound solid background opacity，以及 direct embedded image background 的 crop/opacity/单 owner replacement 已有最小闭环；`PpjConnectorCompositingOpacityAuthorsAndReprojects` 覆盖 connector 的 grammar token、既有 stroke alpha 乘法、去嵌入 PPJ 后的 `stroke.opacity` 规范投影和类型保留，`PpjShapeBackedCompositingOpacityAuthorsAndReprojects` 覆盖 icon/placeholder 的 native shape paint alpha 与既有 projection boundary，`PpjAuthoredGlowEffectWritesShapeAndPictureOwners` 覆盖 shape/picture 的 `rad`、颜色、alpha、glow→outer shadow 顺序、回读和半径拒绝，`PpjSourceBoundShapeImageGlowEditsOwnersAndReprojects` 覆盖去嵌入 PPJ 后 shape/picture 三类 glow leaf 的 owner-local token splice、SlidePart-only changed-part、外层 shadow 保留、Open XML 校验和二次投影，`PpjSourceBoundShapeImageGlowLeavesStayOpaqueForComplexEffectLists` 覆盖复杂 effect list 不发 glow leaf/capability且 picture 保持 opaque；`PpjSourceBoundShapeImageInnerShadowEditsOwnersAndReprojects` 覆盖去嵌入 PPJ 后 shape/picture 五类 inner-shadow leaf 的 owner-local token splice、SlidePart-only changed-part、外层 shadow 保留、Open XML 校验和二次投影，`PpjSourceBoundShapeImageInnerShadowLeavesStayOpaqueForComplexEffectLists` 覆盖 glow+inner-shadow 复杂 effect list 不发 inner-shadow leaf/capability且 picture 保持 opaque；`PpjAuthoredSoftEdgeEffectWritesShapeAndPictureOwners` 覆盖 shape/picture 的 `rad`、soft-edge 顺序、回读和半径拒绝，`PpjSourceBoundShapeImageSoftEdgeEditsOwnersAndReprojects` 覆盖 direct/outer-shadow→soft-edge shape/picture 的 `shapeSoftEdgeRadiusEmu`/`imageSoftEdgeRadiusEmu`、owner-local token splice、SlidePart-only changed-part、前置效果保留、Open XML 校验和二次投影，`PpjSourceBoundShapeImageSoftEdgeLeavesStayOpaqueForComplexEffectLists` 覆盖复杂 effect list 不发 soft-edge leaf/capability且 picture 保持 opaque；`PpjAuthoredInnerShadowEffectWritesShapeAndPictureOwners` 覆盖 shape/picture 的 blur/distance/direction、颜色、alpha、glow→inner shadow→outer shadow→soft edge 顺序、回读和半径拒绝，`PpjAuthoredReflectionEffectWritesShapeAndPictureOwners` 覆盖 shape/picture 的 blur、起止 opacity、distance、direction、全跨度位置、glow→inner shadow→outer shadow→reflection→soft edge 顺序、回读和拒绝；`SourceBoundImageBackgroundCropAndOpacityEditOnlySlideAndReprojects` 和 `PpjSourceBoundMasterAndLayoutImageBackgroundReplacementClosesRelationshipsAndReprojects` 分别证明背景 owner 的字段回写与关系/媒体闭包。下一步继续按 native editable、source-preserved、lossy-only 三类结果补齐 compositing matrix、shape/image 3-D、direct text-run glow/inner shadow/reflection/soft edge 之外的复杂文本 effect、非简单 partial/formula mask adjustments、复杂 mask/effect graph 和背景层关系；任何 lossy 转换必须显式声明，不得把整页截图当作可编辑背景。

**本轮 source-bound 增量：** 普通 shape/line 的单一直接 outer shadow 现在额外发出 `shadowRotateWithShape` 布尔叶子；图片的同一 owner 也发出 `imageShadowRotateWithShape`、`imageShadowBlurRadiusEmu`、`imageShadowDistanceEmu`、`imageShadowDirectionDegrees`、`imageShadowAlignment` 与 `imageShadowOpacityThousandthPercent`。七者只接受已有严格直接 `outerShdw` 中的显式 `rotWithShape` `0/1`、非负整数 `blurRad`/`dist`、DrawingML `dir`、九宫格 `algn` 或 direct RGB/theme color `alpha` `0`–`100000` token，编辑只替换目标 token，保留其它 shadow 属性和 SlidePart 外的内容；`PpjSourceBoundShapeShadowRotateWithShapeEditsLeafAndReprojects`、`PpjSourceBoundImageShadowRotateWithShapeEditsLeafAndReprojects`、`PpjSourceBoundImageShadowBlurEditsLeafAndReprojects`、`PpjSourceBoundImageShadowDistanceEditsLeafAndReprojects`、`PpjSourceBoundImageShadowDirectionEditsLeafAndReprojects`、`PpjSourceBoundImageShadowAlignmentEditsLeafAndReprojects` 与 `PpjSourceBoundImageShadowOpacityEditsLeafAndReprojects` 分别覆盖 shape/line、picture、changed-part 和二次投影。

**图片 outer-shadow RGB 增量：** `imageShadowColorRgb` 现在绑定 `p:pic/a:spPr/a:effectLst/a:outerShdw/a:srgbClr/@val`，只在单一 direct outer shadow、单一显式六位 RGB color token 的拓扑下投影；source-bound 编辑只替换该 `val`，保留图片 payload、mask、border、blur/distance/direction/alignment、rotateWithShape、alpha 及 SlidePart 外内容。`PpjSourceBoundImageShadowColorEditsLeafAndReprojects` 已验证 `17324D` → `0F6B5B`、Open XML、SlidePart-only changed-part、非目标 ZIP part byte-identical 和二次投影；缺失/复杂 color topology 继续 source-owned/fail-closed。

**图片 outer-shadow theme-color 增量：** `imageShadowColorScheme` 现在绑定同一 owner 的 `a:schemeClr/@val`，只接受一个直接 canonical theme token；`PpjSourceBoundImageShadowColorSchemeEditsLeafAndReprojects` 已验证 `accent1` → `accent2` 的单字段 token splice、图片及其它 shadow 状态保留、SlidePart-only changed-part、Open XML 和二次投影。缺失/复杂 color topology、transforms 和其它 effect graph 继续 source-owned/fail-closed。

**本轮新增窄闭环：** `shape3dExtrusionHeightEmu` 绑定普通 shape 的直接、无子节点 `a:sp3d/@extrusionH` canonical 非负坐标；source-bound 编辑只 token-splice 该高度，保留 `z`、`contourW`、材质和其它 3-D markup，复杂/扩展/异常 `sp3d` 继续 source-owned。`PpjSourceBoundShape3dExtrusionHeightLeafEditsAndReprojects` 覆盖 native leaf、SlidePart-only changed-part、Open XML 校验和二次投影；这不是完整 3-D 场景编辑器。

**本轮继续拆出的窄闭环：** `shape3dDepthEmu` 绑定同一普通 shape 的直接、无子节点 `a:sp3d/@z` signed 32-bit EMU 坐标；source-bound 编辑只 token-splice 深度值，保留 `extrusionH`、`contourW`、材质和其它 3-D markup，复杂/扩展/异常 `sp3d` 继续 source-owned。`PpjSourceBoundShape3dDepthLeafEditsAndReprojects` 覆盖负值 canonical leaf、SlidePart-only changed-part、Open XML 校验和二次投影；这仍不是完整 3-D 场景编辑器。

**本轮继续收窄 F-05 的图片 3-D 残差：** 已有 `shape3dDepthEmu` 也绑定严格图片 owner 的 `p:pic/p:spPr/a:sp3d/@z`；新增 additive `PresentationImage.shape_3d_depth_emu` source-bound 载体，`PpjSourceBoundPictureShape3dDepthLeafEditsAndReprojects` 验证负值深度的单 `@z` token splice、仅目标 SlidePart、图片关系/mask/effect 与其它 3-D 状态保留、Open XML 和二次投影。图片的 bevel、scene、颜色和复杂/扩展 3-D graph 仍 source-owned。

**本轮继续拆出同一图片 owner 的第二个 3-D 标量：** 已有 `shape3dExtrusionHeightEmu` 也绑定严格图片 owner 的 `p:pic/p:spPr/a:sp3d/@extrusionH`；新增 additive `PresentationImage.shape_3d_extrusion_height_emu` source-bound 载体，`PpjSourceBoundPictureShape3dExtrusionHeightLeafEditsAndReprojects` 验证非负高度的单 `@extrusionH` token splice、仅目标 SlidePart、图片关系/mask/effect 与其它 3-D 状态保留、Open XML 和二次投影。图片的 bevel、scene、颜色和复杂/扩展 3-D graph 仍 source-owned。

**本轮再拆出的窄闭环：** `shape3dContourWidthEmu` 绑定同一普通 shape 的直接、无子节点 `a:sp3d/@contourW` canonical 非负坐标；source-bound 编辑只 token-splice 轮廓宽度，保留 `z`、`extrusionH`、材质和其它 3-D markup，复杂/扩展/异常 `sp3d` 继续 source-owned。`PpjSourceBoundShape3dContourWidthLeafEditsAndReprojects` 覆盖 native leaf、SlidePart-only changed-part、Open XML 校验和二次投影；这仍不是完整 3-D 场景编辑器。

**本轮补齐同一 owner 的有限枚举：** `shape3dPresetMaterial` 绑定直接、无子节点 `a:sp3d/@prstMaterial`，只接受 canonical DrawingML preset-material token；source-bound 编辑只 token-splice 材质名，保留深度、挤出高度、轮廓、其它 3-D markup 和关系。`PpjSourceBoundShape3dPresetMaterialLeafEditsAndReprojects` 覆盖 `metal` → `matte`、SlidePart-only changed-part、Open XML 校验和二次投影；未知 token、复杂/扩展/异常 `sp3d` 继续 source-owned。这仍不是完整 3-D 场景或材质编辑器。

**本轮继续拆出同一图片 owner 的有限 3-D 枚举：** 已有 `shape3dPresetMaterial` 也绑定严格图片 owner 的 `p:pic/p:spPr/a:sp3d/@prstMaterial`；新增 additive `PresentationImage.shape_3d_preset_material` source-bound 载体，`PpjSourceBoundPictureShape3dPresetMaterialLeafEditsAndReprojects` 验证 `metal` → `matte` 的单 token splice、仅目标 SlidePart、图片关系/mask/effect 与其它 3-D 状态保留、Open XML 和二次投影。未知 token、图片 bevel、scene、颜色和复杂/扩展 3-D graph 仍 source-owned。

**本轮拆出的 bevel 窄闭环：** `shape3dBevelTopWidthEmu` 绑定普通 shape 的单一直接 `a:sp3d/a:bevelT/@w` canonical 非负坐标；`bevelT` 只允许保留 `w`、`h`、`prst` 直接属性，source-bound 编辑只 token-splice `w`，保留高度、预设和根部 3-D markup。`PpjSourceBoundShape3dBevelTopWidthLeafEditsAndReprojects` 覆盖 native leaf、SlidePart-only changed-part、Open XML 校验和二次投影；底面 bevel、颜色/扩展/未知属性、异常值和完整 3-D 场景仍 source-owned。

**本轮继续拆出同一图片 owner 的 bevel 标量：** 已有 `shape3dBevelTopWidthEmu` 也绑定严格图片 owner 的 `p:pic/p:spPr/a:sp3d/a:bevelT/@w`；新增 additive `PresentationImage.shape_3d_bevel_top_width_emu` source-bound 载体，`PpjSourceBoundPictureShape3dBevelTopWidthLeafEditsAndReprojects` 验证 `1200` → `2400` 的单 `bevelT/@w` token splice、仅目标 SlidePart、图片关系/crop/mask/effect 与其它 3-D 状态保留、Open XML 和二次投影。未知属性、额外子节点、底面 bevel、scene、颜色和复杂/扩展 3-D graph 仍 source-owned。

**本轮继续拆出同一图片 owner 的第二个 bevel 标量：** 已有 `shape3dBevelTopHeightEmu` 也绑定严格图片 owner 的 `p:pic/p:spPr/a:sp3d/a:bevelT/@h`；新增 additive `PresentationImage.shape_3d_bevel_top_height_emu` source-bound 载体，`PpjSourceBoundPictureShape3dBevelTopHeightLeafEditsAndReprojects` 验证 `800` → `1600` 的单 `bevelT/@h` token splice、仅目标 SlidePart、图片关系/crop/mask/effect 与其它 3-D 状态保留、Open XML 和二次投影。未知属性、额外子节点、底面 bevel、scene、颜色和复杂/扩展 3-D graph 仍 source-owned。

**本轮继续拆出的 bevel 窄闭环：** `shape3dBevelTopHeightEmu` 绑定同一普通 shape 的单一直接 `a:sp3d/a:bevelT/@h` canonical 非负坐标；source-bound 编辑只 token-splice `h`，保留宽度、预设和根部 3-D markup。`PpjSourceBoundShape3dBevelTopHeightLeafEditsAndReprojects` 覆盖 native leaf、SlidePart-only changed-part、Open XML 校验和二次投影；底面 bevel、颜色/扩展/未知属性、异常值和完整 3-D 场景仍 source-owned。

**本轮再拆出的 bevel 窄闭环：** `shape3dBevelTopPreset` 绑定同一普通 shape 的单一直接 `a:sp3d/a:bevelT/@prst` 有限 DrawingML bevel 预设 token；source-bound 编辑只 token-splice `prst`，保留宽高、根部 3-D markup 和其余包内容。`PpjSourceBoundShape3dBevelTopPresetLeafEditsAndReprojects` 覆盖 `angle` → `softRound`、SlidePart-only changed-part、Open XML 校验和二次投影；底面 bevel、颜色/扩展/未知属性、异常 token 和完整 3-D 场景仍 source-owned。

**本轮继续拆出同一图片 owner 的 bevel 枚举：** 已有 `shape3dBevelTopPreset` 也绑定严格图片 owner 的 `p:pic/p:spPr/a:sp3d/a:bevelT/@prst`；新增 additive `PresentationImage.shape_3d_bevel_top_preset` source-bound 载体，`PpjSourceBoundPictureShape3dBevelTopPresetLeafEditsAndReprojects` 验证 `angle` → `softRound` 的单 `bevelT/@prst` token splice、仅目标 SlidePart、图片关系/crop/mask/effect、bevel 尺寸与其它 3-D 状态保留、Open XML 和二次投影。未知属性、额外子节点、底面 bevel、scene、颜色和复杂/扩展 3-D graph 仍 source-owned。

**本轮开始拆出的 bottom bevel 窄闭环：** `shape3dBevelBottomWidthEmu` 绑定普通 shape 的单一直接 `a:sp3d/a:bevelB/@w` canonical 非负坐标；`bevelB` 只允许保留 `w`、`h`、`prst` 直接属性，source-bound 编辑只 token-splice `w`，保留高度、预设和根部 3-D markup。`PpjSourceBoundShape3dBevelBottomWidthLeafEditsAndReprojects` 覆盖 native leaf、SlidePart-only changed-part、Open XML 校验和二次投影；顶面 bevel、颜色/扩展/未知属性、异常值和完整 3-D 场景仍 source-owned。

**本轮继续拆出同一图片 owner 的 bottom bevel 标量：** 已有 `shape3dBevelBottomWidthEmu` 也绑定严格图片 owner 的 `p:pic/p:spPr/a:sp3d/a:bevelB/@w`；新增 additive `PresentationImage.shape_3d_bevel_bottom_width_emu` source-bound 载体，`PpjSourceBoundPictureShape3dBevelBottomWidthLeafEditsAndReprojects` 验证 `1200` → `2400` 的单 `bevelB/@w` token splice、仅目标 SlidePart、图片关系/crop/mask/effect 与其它 3-D 状态保留、Open XML 和二次投影。未知属性、额外子节点、顶面 bevel、scene、颜色和复杂/扩展 3-D graph 仍 source-owned。

**本轮继续拆出的 bottom bevel 窄闭环：** `shape3dBevelBottomHeightEmu` 绑定同一普通 shape 的单一直接 `a:sp3d/a:bevelB/@h` canonical 非负坐标；source-bound 编辑只 token-splice `h`，保留宽度、预设和根部 3-D markup。`PpjSourceBoundShape3dBevelBottomHeightLeafEditsAndReprojects` 覆盖 native leaf、SlidePart-only changed-part、Open XML 校验和二次投影；顶面 bevel、颜色/扩展/未知属性、异常值和完整 3-D 场景仍 source-owned。

**本轮继续拆出同一图片 owner 的第二个 bottom bevel 标量：** 已有 `shape3dBevelBottomHeightEmu` 也绑定严格图片 owner 的 `p:pic/p:spPr/a:sp3d/a:bevelB/@h`；新增 additive `PresentationImage.shape_3d_bevel_bottom_height_emu` source-bound 载体，`PpjSourceBoundPictureShape3dBevelBottomHeightLeafEditsAndReprojects` 验证 `800` → `1600` 的单 `bevelB/@h` token splice、仅目标 SlidePart、图片关系/crop/mask/effect 与其它 3-D 状态保留、Open XML 和二次投影。未知属性、额外子节点、顶面 bevel、scene、颜色和复杂/扩展 3-D graph 仍 source-owned。

**本轮继续拆出的 bottom bevel 窄闭环：** `shape3dBevelBottomPreset` 绑定同一普通 shape 的单一直接 `a:sp3d/a:bevelB/@prst` 有限 DrawingML bevel 预设 token；source-bound 编辑只 token-splice `prst`，保留宽高、根部 3-D markup 和其余包内容。`PpjSourceBoundShape3dBevelBottomPresetLeafEditsAndReprojects` 覆盖 `angle` → `softRound`、SlidePart-only changed-part、Open XML 校验和二次投影；顶面 bevel、颜色/扩展/未知属性、异常 token 和完整 3-D 场景仍 source-owned。

**本轮继续拆出同一图片 owner 的 bottom bevel 枚举：** 已有 `shape3dBevelBottomPreset` 也绑定严格图片 owner 的 `p:pic/p:spPr/a:sp3d/a:bevelB/@prst`；新增 additive `PresentationImage.shape_3d_bevel_bottom_preset` source-bound 载体，`PpjSourceBoundPictureShape3dBevelBottomPresetLeafEditsAndReprojects` 验证 `angle` → `softRound` 的单 `bevelB/@prst` token splice、仅目标 SlidePart、图片关系/crop/mask/effect、bevel 尺寸与其它 3-D 状态保留、Open XML 和二次投影。未知属性、额外子节点、顶面 bevel、scene、颜色和复杂/扩展 3-D graph 仍 source-owned。

### F-06 Table、Cell Style 和 Table Layout

**文字变形删除增量（2026-09-10）：** 表格 `text.style.textWarpPreset/textWarpAdjustments` 的删除、清空和恢复闭合，简单样式删除后可恢复结构化文字；见 F-03 的 160/160 证据。


**平面文字深度删除增量（2026-09-10）：** 表格 `text.style.flatTextZ` 的零值、正负边界、删除和恢复闭合，简单样式删除后可恢复结构化文字；见 F-03 的 152/152 证据。


2026-09-10：`text.style.fromWordArt` 删除/布尔恢复与简单样式删除后的
紧凑文字恢复已闭合，文字变形和固定段落结构保留；见 F-03 的 144/144。


2026-09-10：`text.style.compatibleLineSpacing` 删除/布尔恢复与简单样式
删除后紧凑文字恢复已闭合，固定段落与实际间距保留；见 F-03 的 136/136。


2026-09-10：表格 `text.style.spaceFirstLastParagraph` 删除/恢复及简单样式
删除后的紧凑文字恢复已闭合，固定段落结构和间距保留；见 F-03 的 128/128。


2026-09-10：`text.style.forceAntiAlias` 删除与布尔恢复已进入共享回归，
包含简单样式整组删除和表格紧凑文字恢复；见 F-03 的 120/120 专项。


2026-09-10：表格 `text.style.anchorCenter` 删除/布尔恢复、简单样式删除后
紧凑文字恢复已进入共享回归，保留垂直对齐和固定文字拓扑；见 F-03 的 112/112。


2026-09-10：表格 `text.style.autoFit/normalAutoFit` 模式和百分比删除恢复
进入共享生命周期回归，包含删除后紧凑文字恢复；详见 F-03 的 104/104 专项。


2026-09-10：表格 `text.style.margins` 四边删除、空对象/整组删除和零值恢复
已闭合，包含简单样式删除后的紧凑文字恢复；详见 F-03 的 96/96 相关回归。


2026-09-10：表格 `text.style.columns` 的删除、显式 1/3/16 恢复及简单样式
删除后的紧凑文字恢复已进入共享回归，见 F-03；保持固定段落/run 拓扑与编辑权限边界。


**分栏间距删除增量（2026-09-10）：** 表格 `text.style.columnGap` 的显式 0、小数、上限、删除和恢复已闭合，列数和方向保留；简单样式删除及纯文本规范化/恢复证据见 F-03。

**垂直对齐删除增量（2026-09-10）：** 表格 `text.style.verticalAlignment` 的三值、删除和恢复已闭合，并保留 anchorCenter；简单样式删除及纯文本规范化/恢复证据见 F-03。

**垂直溢出删除增量（2026-09-10）：** 表格 `text.style.verticalOverflow` 的三种显式值、删除和恢复已闭合，并保留 horizontalOverflow；简单样式删除及纯文本规范化/恢复证据见 F-03。

**水平溢出删除增量（2026-09-10）：** 表格 `text.style.horizontalOverflow` 的 overflow/clip/省略和恢复已闭合，verticalOverflow 保持独立；简单样式删除与纯文本规范化/恢复证据见 F-03。

**换行模式删除增量（2026-09-10）：** 表格 `text.style.wrap` 的 square/none/省略和恢复已闭合；与文本 owner 共用最小实验，简单样式删除及纯文本规范化/恢复证据见 F-03。

**文字方向删除增量（2026-09-10）：** 表格 `text.style.verticalText` 支持三种显式方向、删除和恢复；与 columnDirection 共用最小实验，简单样式删除及纯文本规范化/恢复证据见 F-03。

**分栏方向删除增量（2026-09-10）：** 表格 structured `text.style.columnDirection` 已支持双向显式值、删除和恢复；简单样式整体移除及表格纯文本规范化/恢复证据见 F-03 的共享生命周期实验。

**文本旋转删除增量（2026-09-10）：** 表格 `text.style.rotation` 支持省略删除及显式 0/负角度恢复；与 upright 合并移除简单样式对象后，可按原生固定拓扑恢复结构化文本。详见 F-03 rotation 增量；其它布局字段删除仍开放。

**文本体删除增量（2026-09-10）：** 表格 `text.style.upright` 已与其它文本 owner 对齐，支持 true/false/删除及移除 upright-only style；删除最后一个属性后的普通文本规范化与恢复路径已验证。证据和已有失败见 F-03 upright 增量，其它 body style 删除与复杂继承仍开放。

**优先级：P1；状态：部分完成（固定矩形表格 profile 已交付）。**

**当前进度：** 支持矩形表格、合并、列宽、行高、cell text、边框、填充、对齐、headerRows、banding、图片填充、多 header、部分 source-bound 表格格式叶子。本轮补齐 recognized rectangular table 的 `setTableStyle`：`headerRows`（仅 0/1）、`bandedRows`、`bandedColumns`、`firstColumnEmphasis`、`lastColumnEmphasis` 可在已有 `a:tblPr` 内 source-bound 增改/清除，只改所属 SlidePart，并通过二次投影恢复；新增 `setTableGeometry`，在保持列/行 ID、物理网格和 merge topology 不变时，可 source-bound 修改已有列宽和行高，仍只写所属 SlidePart，且不宣称自动文字 reflow。`setTableCellStyle` 覆盖直接单元格视觉样式窄 profile：每个物理 cell 的 direct `a:noFill`、单一 RGB `a:solidFill`（含 alpha）、有界 `a:gradFill`、单一直接嵌入 `a:blipFill` image paint（含 crop、stretch/tile、alpha），`lnL/lnT/lnR/lnB` 的颜色、宽度、preset dash、cap、join，以及跨一个或多个段落、所有直接文本 run 共享同一受限样式时的字号、粗体、斜体、RGB 字色、字体和下划线/删除线，可在保留原 `a:tcPr`、文本 body 其余节点、段落/run 数和文本拓扑的前提下 source-bound 修改。现在还支持一个更窄的 mixed-run body profile：当 cell 只有固定段落/直接纯文本 run、每个 run 的直接样式均属于同一受限集合但彼此不一致时，投影保留 `text.paragraphs[].runs[].style`，`setTableCellStyle/table.cell.textStyle` 可回写每个原 run 的样式；编译保持 run/段落拓扑和未建模 XML，不把混合样式压成一个 cell-level uniform style。对每个段落由一个或多个直接纯文本/field/break inline 构成的 cell，文本替换也会保留段落数/run 数和原有样式；固定拓扑单段落的嵌入式 picture bullet 通过 `style.bullet = { type: "picture", asset }` 投影，source-bound 文字编辑保持该 marker，并将 PPJ asset ID 按内容哈希映射回 native picture-bullet 资源；同一窄 profile 还允许 picture bullet 的直接 `fontFamily`、RGB/主题色（含 alpha）、follow-text 和 `size`/`sizePercent` 样式在已有 `a:pPr` 内 source-bound 改写，保持 marker 和关系不变。外部 URI picture bullet、图片 bullet 的未知/效果化子图、未授权的 line-break topology change 和富文本或空 cell 加文仍 fail-closed。图片 replacement 会同步所属 SlidePart、其 `.rels` 和媒体 part，并清理不再引用的旧关系/媒体；当同一 SlidePart 的多个 cell 共享旧 image relationship 时，采用 copy-on-write：只为被编辑 cell 建立新 relationship/media，旧关系和媒体继续服务其他 cell。`PpjSourceBoundTableCellStylesEditOnlySlideAndReproject` 证明填充、边框、固定拓扑多-run 文本替换、跨段落样式一致多-run 文本样式和表格几何闭环，`PpjSourceBoundTableCellMixedRunStylesEditOnlySlideAndReproject` 证明混合 run body 的样式/文字回写、单 SlidePart footprint 和二次投影，`PpjSourceBoundTableCellImageFillReplacesRelationshipAndReprojects` 证明单 owner 图片关系/媒体替换，`PpjSourceBoundTableCellSharedImageFillPreservesOtherReference` 证明多 cell 复用同一关系时旧 owner 不被误删，`PpjSourceBoundTableCellPictureBulletPreservesAssetAndEditsText` 证明嵌入式 picture bullet 的 PPJ 投影、文字编辑、直接 bullet 样式 source-bound 替换与二次投影，`PpjSourceBoundTableCellLineBreakPreservesInlineAndEditsText` 证明固定拓扑 `a:br` 的 PPJ `run.break` 投影、文字编辑、SlidePart-only footprint、二次投影和 topology-change 拒绝。

**本轮增量（段落叶）：** mixed-run body 现在允许一个已存在的直接段落 `alignment`（`left`/`center`/`right`/`justify`/`distributed`）、`indent`/`hanging`、direct `spaceBefore`/`spaceAfter`/`lineSpacing` 的 points 或 multiplier 形式、direct character/auto-number bullet（保留有限 scheme 与 `startAt`）以及固定拓扑单段落的嵌入式 picture bullet；character/auto-number/picture bullet 均支持受限的 direct bullet style（字体、RGB/主题色、跟随正文、字号/百分比），picture bullet 只接受已存在的 `asset` 关系并保持其 native picture marker 拓扑。`defaultText` 仍限于字号、字体、RGB 颜色。固定拓扑段落还投影 `style.tabStops`（points + `left`/`center`/`right`/`decimal`）并可用 `style.noTabStops: true` 删除 modeled tab-stop list。投影保留 `text.paragraphs[].style` 对应字段，source-bound `setTableCellStyle/table.cell.textStyle` 只改已有 `a:pPr` 子节点（`@algn`、`@marL`、`@indent`、`a:buChar`/`a:buAutoNum`/`a:buBlip`、`a:buFont`/`a:buClr`/`a:buSz*`、`a:tabLst`、`a:spcBef`、`a:spcAft`、`a:lnSpc`、`a:defRPr`），同时保持段落/run 拓扑和其余 XML 不变。未知/效果化 bullet 子图、defaultText 的主题/效果/复杂装饰以及其它未建模段落属性仍不在该 profile。

本轮新增窄闭环：`tableLastRow` 绑定矩形表格直接 `p:graphicFrame/a:graphicData/a:tbl/a:tblPr/@lastRow`，并在 `table.style.lastRow` 保留可选的 authored/source-bound 布尔语义；只替换已有 `0|1` token，保持表格网格、单元格内容、其它标志和关系不变，缺失/重复/非法/不规则表格不发 leaf。

**本轮增量（文本容器叶）：** 固定拓扑 mixed-run table cell 现在可以在结构化 rich text 的顶层暴露 `text.style`，把已有直接 `a:bodyPr` 的有界字段映射为 PPJ `textBoxStyle`：`verticalAlignment`、`wrap`、`margins`（四边 inset）、`columns`、`columnGap`、`columnDirection`、`verticalText`、`rotation`、`horizontal/verticalOverflow`、`upright`、有限 `autoFit` 和 `normalAutoFit.fontScale/lineSpacingReduction` 百分比。`nativeRef.leaves[]` 对普通文本框/带文本形状/占位符还提供两个 canonical normAutofit 百分比叶子；它们只改变对应直接属性，不创建新节点或改动另一个属性。`setTableCellStyle/table.cell.textStyle` 通过同一 `TextBody` 回写这些已有 body-property leaves，只改所属 SlidePart；不会新建 text body、推断继承值，也不会因为 PPJ 省略字段而删除 native leaf。`PpjSourceBoundTextBodyStyleEditsTextShapeAndReprojects` 与 `PpjSourceBoundNormalAutoFitLeavesEditAndReproject` 现在覆盖普通 text body 的 authored/投影和 rotation/overflow/upright source-bound 修改；表格 cell 复用同一 body-property writer。显式 body-property 删除、继承/未知/效果化 bodyPr 和自动 reflow 仍 source-owned/fail-closed。

**字段增量：** 固定拓扑 table cell 现在可以投影 direct text/field 混合 body；`run.field` 保留 typed `type`、`text` 和 ID，source-bound 只允许修改缓存 display text，字段 ID/type 与段落/run 拓扑保持不变。`PpjSourceBoundTableCellFieldPreservesIdentityAndEditsCachedText` 覆盖 authored → 去嵌入 PPJ → 缓存文本写回 → 二次投影，以及字段身份变更的 fail-closed 拒绝；复杂字段图、字段关系和自动求值仍不开放。

**仍缺：** 完整 table style inheritance、带段落/列表/字段/布局属性的 mixed-run body、未落入受限直接 run 样式集合的效果/主题图、文本 reflow/自动换行、跨 owner 或外部 image relationship 的完整治理（同一 SlidePart 内多 cell 复用关系的 copy-on-write 已覆盖）、picture bullet 的未知/效果化 marker 子图、复杂主题/效果和混合/扩展 paint、动态 row/column reflow（当前 geometry 只改固定网格的宽高，不重算内容）、未知 table extension、宿主自动调整后的精确恢复，以及与通用 layout solver 的联动；PPJ schema 中 `headerRows > 1` 的语义仍只适用于 authored/嵌入 PPJ 恢复，第三方表格不能凭一个 PowerPoint 布尔属性反推多 header。

这里的“带段落属性”不再包含上面已闭合的 direct alignment、indent/hanging、三类 direct spacing、character/auto-number/picture bullet、受限 bullet style、受限 defaultText 和固定拓扑 `run.break`；其余段落/列表/字段/布局属性仍保持 source-owned/fail-closed。

这里的“布局属性”也不再包含上一段已闭合的 bounded text-container `text.style`：vertical alignment、wrap、四边 inset、columns/column gap/direction、vertical text、rotation、horizontal/vertical overflow、upright、有限 auto-fit 和 canonical `normalAutoFit` 百分比。未落入该 bodyPr profile 的继承、显式删除以及未知/效果化扩展仍属于待实现边界。

**待实现与验收：** 继续将 table cell 的 visual style、text layout、row/column geometry 分离；已完成的 `setTableStyle`、`setTableGeometry` 与 bounded `setTableCellStyle` 都保留最小 source-bound 闭环（声明的字段、SlidePart/`.rels`/media footprint、二次投影），其中 geometry 只接受同一列/行 ID 和同一物理网格，列宽总和/行高总和仍须满足原 frame 约束，不执行自动 reflow；固定拓扑文本替换可覆盖一个或多个段落内的直接纯文本多 run，按原 run 数保留样式与文本拓扑；text-style capability 还可覆盖所有直接文本 run 共享同一受限 RGB/字体/装饰样式的多 run cell，也可覆盖固定段落、无列表/字段/高级效果的样式不一致 mixed-run body，并将各自样式回写到每个原 run。image capability 仅接受由该表格所属 SlidePart 直接拥有、可解析且可安全替换的嵌入关系；同一 SlidePart 内的多引用关系采用 copy-on-write 并保留仍被使用的旧媒体，外部、跨 owner、未知扩展和 ambiguous closure 仍撤回对应 capability。带段落/列表/字段/效果的文本 body、自动 reflow、共享/外链关系仍撤回对应 capability。下一步按段落样式、跨 owner 关系治理、style inheritance、reflow 分别定义所有权。任何 topology-changing edit、未知 `tcPr` paint/effect 或未证明的 relationship closure 都要显式列为未支持。

**F-06 后续边界：** bounded text-container bodyPr 已有独立闭环；后续工作聚焦 style inheritance、未建模 bodyPr/paragraph/list/effect graph、自动 reflow 与跨 owner/shared relationship 治理，不再把已支持的 direct body layout 计作缺口。

**F-06 外链 picture bullet 边界补充：** 已有外链 URI marker 可在固定拓扑中编辑相邻文字、direct marker 样式并替换为新的 owner-local URI；source-bound export 采用 append-only 关系策略保留旧外链关系，以维持原图闭包。跨 owner 关系治理和未知/效果化 marker 子图仍 fail-closed。

**F-06 段落 tab stop 补充：** 固定拓扑表格/文本段落共用 `a:tabLst` 的有界 profile；PPJ 可读写 points 位置和四种对齐（包括 `distributed` 段落对齐），并以 `noTabStops: true` 显式清除 modeled 列表。`PpjSourceBoundTableCellMixedRunStylesEditOnlySlideAndReproject` 现在覆盖表格段落的 distributed 对齐、两个直接 tab stop 的投影以及位置/对齐的 SlidePart-only source-bound 回写。其余列表继承、自动换行/reflow 和未知段落扩展仍 source-owned。

### F-07 Chart、ChartML、嵌入工作簿和扩展图表

**图表外阴影变换增量（2026-09-09）：** `chartTextStyle.shadow` 新增可选 `scaleX/scaleY/skewX/skewY`。缩放为有符号倍率（1 为原尺寸），以原生 0.00001 精度保存；倾斜用度表达，按 1/60000 度舍入，舍入后仍须严格位于 -90～90 度内。省略、零和负值独立保留。现有 line/combo 最小生命周期实验扩展到变换修改、清空、删除和重建，检查发光、内阴影、倒影与柔化边缘共存且独立保留，并核对二次投影、原文字、原字节 no-op 和非目标 ZIP；样式优先级及 vector 默认/片段覆盖贯通。普通文字、形状、图片的 imported 变换保护条件保持；宿主显示和完整 F-07 继续待补。

**图表倒影变换增量（2026-09-09）：** `chartTextStyle.reflection` 新增可选 `fadeAngle`、`scaleX/scaleY`、`skewX/skewY`、`alignment`、`rotateWithShape`，补齐直接倒影的 14 个原生属性。缩放为有符号倍率，保留零和负数；倾斜以度表达，原生精度为 1/60000 度，舍入后仍须严格位于 -90～90 度内；旋转的 `false` 与省略独立保留。现有 line/combo 生命周期实验覆盖变换属性修改、清空、删除和重建，检查二次投影、其他效果、文字及非目标 ZIP 保留；样式优先级、vector 默认值和片段覆盖沿用同一对象。普通文字/形状/图片的 imported 变换保护条件继续保留，任意效果图和宿主显示仍待补；F-07 整项继续开放。下方早期增量中的“缩放/倾斜待补”由本增量接续。

**图表倒影渐变位置增量（2026-09-09）：** `chartTextStyle.reflection.startPosition/endPosition` 以 0～1 表达透明度渐变上的起止位置，按原生千分之一百分比精度保存；两端可独立设置、移除，并保留相同、反向及显式 0/1。图表倒影不再限定显式 0/100000，省略位置属性保留原生缺省；普通文字、形状、图片的 imported 完整跨度保护条件继续保留。现有普通图表/组合图生命周期实验已扩展到这两个属性的修改、移除、二次投影、样式优先级和 vector 默认值，检查其余倒影参数、其他效果、文字及非目标 ZIP 保留。位置不是裁切或幻灯片坐标；缩放/倾斜等其他倒影属性和宿主显示仍待补。

**图表文字倒影增量（2026-09-09）：** `chartTextStyle.reflection` 承载完整跨度倒影，可选 `blur`、`startOpacity`、`endOpacity`、`distance`、`angle`，范围沿用普通文字倒影。`{}` 保留倒影及原生缺省，显式零独立保存，省略整个字段只删除倒影；起止透明度支持 opacity token。标题、图例、轴、数据标签、趋势线富文本和 vector 默认/片段样式贯通，原生顺序为 glow → innerShdw → outerShdw → reflection → softEdge。普通图表/组合图实验检查清空属性、删除、重建、每步二次投影、原字节 no-op、其余效果、文字和非目标 ZIP 保留。同时修复既有文字倒影角度叶子的单位归类冲突，45.5° → 90.25° 的回写按原生角度单位验证。原生起止位置为 0/100000；不同跨度、缩放/倾斜等额外属性、重复/乱序或未知后代继续保留原始内容。宿主显示和 F-07 整项仍待补。

**图表文字内阴影增量（2026-09-09）：** `chartTextStyle.innerShadow` 必填 `color`，可选 `blur`（0～1000 pt）、`distance`（0～100000 pt）、`angle`（-360～360 度）、`opacity`（数值或透明度 token）。省略属性保留原生缺省，显式零独立保存；省略整个字段只移除内阴影。主题色、grammar 颜色和透明度贯通标题、图例、轴、数据标签、趋势线富文本及 vector 默认/片段样式。原生顺序为 glow → innerShdw → outerShdw → softEdge。普通图表/组合图实验检查每步二次投影、独立删除/重建、原字节 no-op、文字和非目标 ZIP 保留；重复、乱序、超范围和未知后代继续保留原始内容。完整效果图、宿主显示和 F-07 整项仍待补。

**图表文字柔化边缘增量（2026-09-09）：** `chartTextStyle.softEdge` 复用 `{ radius }`，半径为 0～1000 pt，按原生 EMU 精度保存；显式零与省略字段分别保留。标题、图例、轴、数据标签、趋势线富文本及 vector 文字共用字段，片段的零半径可覆盖标题默认值。原生按 glow → outerShdw → softEdge 排列，三种效果可分别修改、删除和重建。同时修正新建图表时标准主题阴影/发光颜色被资源引用校验提前拒绝的问题，错误类型的 grammar token 和不支持主题色的前景颜色仍会拒绝。普通图表/组合图实验检查每步二次投影、原字节 no-op、原文字及非目标 ZIP 保留。缺失半径、重复/乱序效果及未知后代继续保留原始内容；宿主显示和完整 F-07 仍待补。

**图表文字发光增量（2026-09-09）：** `chartTextStyle.glow` 复用普通文字的 `{ color, radius, opacity? }`，半径为 0～1000 pt，保存显式零半径/透明度及透明度缺省。颜色和 opacity token、主题身份贯通标题、图例、轴、数据标签、趋势线富文本及 vector 文字。发光与外阴影可以共存并独立删除、恢复；原生按 glow → outerShdw 排列。普通图表/组合图最小实验检查同一效果列表的两个字段、原字节 no-op、二次投影、文字及非目标 ZIP 保留。缺少原生半径、重复/乱序效果及未知后代继续保留原始内容；宿主显示和完整 F-07 仍待补。

**图表文字阴影增量（2026-09-09）：** `chartTextStyle.shadow` 承载直接外阴影：必填颜色，可选模糊、距离、角度、透明度、对齐及随形状旋转。省略属性保留原生缺省，显式零/false 独立保存；省略整个字段删除阴影。RGB/grammar 颜色、主题身份和 opacity token 可贯通标题、图例、轴、数据标签、趋势线富文本及 vector 文字，片段阴影优先于标题默认值。普通图表/组合图最小实验检查修改、清空属性、删除、重建、原字节 no-op、二次投影与非目标 ZIP 保留。混合/未知效果保持 source-owned；宿主阴影显示及完整 F-07 继续待补。

**图表文字高亮增量（2026-09-09）：** `chartTextStyle.highlight` 复用普通文字颜色及颜色 token，支持 tint/shade 解析，以不透明 RGB 保存文字高亮。高亮与前景填充、字体分别保留，省略字段删除高亮；标题、图例、轴、数据标签、趋势线及其富文本样式共用字段，vector 标题片段可覆盖默认高亮。普通图表/组合图实验覆盖创建、换色、删除、重建、颜色 token 和二次投影，检查原文字、原字节 no-op 与非目标 ZIP 条目；混合样式另检查原生子节点顺序。半透明或透明颜色、主题高亮、原生颜色变换和未知结构继续拒绝不受支持的编辑，宿主显示与完整 F-07 继续待补。

**图表文字 kerning 增量（2026-09-09）：** `chartTextStyle.kerning` 以 pt 表达启用字偶距调整的最小字号阈值，范围 `0`～`768`；原生保存为百分之一点，细于 `0.01pt` 的输入按最近值舍入，中点取偶数。显式 `0` 保留直接阈值，省略字段移除设置。标题、图例、轴、数据标签、趋势线及其富文本样式共用字段；vector 文字保留该值，片段设置优先于标题默认值。普通图表/组合图实验覆盖创建、修改、零值、删除、重建、精度和二次投影，检查原文字、原字节 no-op 与非目标 ZIP 条目；负值和未知字符属性继续拒绝编辑。字体配对度量、宿主排版和完整 F-07 继续待补。

**图表文字字距增量（2026-09-09）：** `chartTextStyle.letterSpacing` 以 pt 表达 `-768`～`768` 的字符间距，复用普通文字范围；原生保存为百分之一点，细于 `0.01pt` 的输入按最近值舍入，中点取偶数。显式 `0` 与省略字段分别表示归零和移除直接设置。标题、图例、轴、数据标签、趋势线及其富文本样式共用字段，vector 文字也能传递，片段字距优先于标题默认值。普通图表/组合图实验覆盖正负值、归零、删除、重建、精度和二次投影，并检查原文字、原字节 no-op 与非目标 ZIP 条目。宿主字形排布和完整 F-07 继续待补。

**图表文字大小写增量（2026-09-09）：** `chartTextStyle.capitalization` 复用普通文字的 `none`、`small`、`all`，保存原生 `cap` 显示样式，原文字保持不变。显式 `none` 与省略字段分别表示取消和移除直接设置。标题、图例、轴、数据标签、趋势线及其段落/片段/段末样式共用字段；vector 文字保留该值，片段设置优先于标题默认值。普通图表/组合图实验覆盖创建、修改、取消、删除、重建和二次投影，检查原文字、原字节 no-op 与非目标 ZIP 条目。字体字形、宿主显示和完整 F-07 继续待补。

**图表文字基线增量（2026-09-09）：** `chartTextStyle.baseline` 复用普通文字的 `-400`～`400` 百分比范围，原生保存为千分之一百分比整数；更细的小数按最近值舍入，中点取偶数。显式 `0` 与省略字段分别表示直接归零和移除设置。标题、图例、轴、数据标签、趋势线及其段落/片段/段末样式共用字段，vector 文字也能传递，片段值优先于标题默认值。普通图表/组合图实验覆盖正负值、归零、删除、重建、精度转换和二次投影，保留原字节 no-op 与非目标 ZIP 条目；宿主字形排版和完整 F-07 继续待补。

**图表文字删除线增量（2026-09-09）：** `chartTextStyle.strike` 接受 `true`、`false` 或 `sngStrike`、`dblStrike`、`noStrike`，复用普通文字的删除线语义。`false` 写入显式 `noStrike`，省略字段则移除直接属性；回投影统一保留原生枚举值。标题、图例、轴、数据标签、趋势线及其富文本样式共用字段，vector 标题和标签也能传递。普通图表/组合图的最小实验覆盖增改、取消、删除、重建、原字节 no-op 和目标 ChartPart 以外内容保留；宿主显示和完整 F-07 继续待补。

**图表文字语言增量（2026-09-09）：** `chartTextStyle.language` 使用现有语言标签或 `string` grammar token，直接承载字符 `lang`。标题、图例、轴、数据标签和趋势线样式，以及趋势线富文本的段落、片段、段末样式共用此字段；显式 `en-US`、大小写和省略状态保留。普通图表和组合图的最小实验覆盖创建、修改、删除、重建、原字节 no-op 和二次投影；语言编辑只改目标 ChartPart。vector 标题默认语言让位于片段显式语言。宿主拼写检查、字典和完整 F-07 仍待补。

**趋势线标签增量（2026-09-09）：** `data.series[].trendlines[].label` 已承载可选 `text`、`numberFormat`、`textStyle`、`fill`、`line`。省略 text 时保留自动公式/R² 内容，`{}` 保留默认标签容器，省略 label 删除标签；text/numberFormat 接受 string grammar token。普通线图和 categorical combo 的最小实验检查 authored、修改、默认容器、删除、重新添加、原生节点和去嵌入后的回投影，label-only 编辑只改目标 ChartPart。JS 工作簿接口的其他图表编辑也保留原有 wire 标签。公式文本、布局扩展、复杂效果和 extension 仍按原始数据保留，不开放 analytics 编辑；自动布局和宿主视觉效果仍未验收。

**趋势线标签布局增量（2026-09-09）：** `label.layout.manual` 承载 `target`、`xMode/yMode/widthMode/heightMode` 和 `x/y/width/height`。数字保留 ChartML 比例值和显式零；省略 layout、空 layout、空 manual 三种状态分别保留。普通线图和 categorical combo 的最小实验覆盖 authored、修改、清空、删除、重建、原生结构和 fresh projection，并检查其他标签字段和非目标 ZIP 字节。扩展、未知布局内容及非法数值仍不开放 analytics 编辑；宿主定位和 SVG 布局效果仍未验收。

**趋势线标签格式标志增量（2026-09-09）：** `label.numberFormatSourceLinked` 配合 `numberFormat` 表达原生 `sourceLinked`：true/false 写入显式值，null 保留原生属性缺省；省略 PPJ 字段沿用 false，回投影也省略 false。普通线图和 categorical combo 覆盖 authored、切换、格式修改、删除、重建及 fresh projection，其他标签状态和非目标 ZIP 字节保持。该字段只保存 ChartML 状态；[微软说明 Office 不使用趋势线标签上的此标志](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/baa94e34-1d3e-4d59-a139-d9955fcdcf79)，不据此宣称工作簿格式同步或宿主视觉效果。

**趋势线标签富文本增量（2026-09-09）：** `label.text` 已扩展为 `{paragraphs:[{style?,runs:[{text,style?}|{break:true,style?}],endStyle?}]}`，文字和样式复用 grammar token；段落默认样式、片段样式和段末样式复用 chart text style，alignment 只用于段落。顺序、空段落、空片段、空格及显式 false 样式保留。单段单个无样式非空片段（不超过 255 字符）回投影为原字符串；省略 text 恢复自动内容。普通线图和 categorical combo 的最小实验覆盖 authored、修改、字符串转换、删除、重建和二次投影，只改目标 ChartPart 的标签文本。公式引用、field、超链接、复杂 body/list 属性和未知字符效果仍保留原始数据，不开放 analytics 编辑；宿主排版和完整 F-07 继续待补。

**后续增量（2026-09-09）：** 导入的 custom `plus/minus` 现在可携带 `formula`、完整 `values` 和可选 `formatCode`，支持固定的本地绝对单行/单列引用。修改误差值会同时写回目标 ChartPart 缓存和唯一嵌入 XLSX 的数值单元格；普通柱/条/线、categorical combo 和分组图表均有回归。无修改保留原始字节，只改端帽等样式不改工作簿。共享/缺失工作簿、缓存与单元格不一致、重叠消费者、公式单元格及未支持的依赖图会拒绝数值编辑且不输出文件。`PpjErrorWorkbookSynchronizesCacheAndCells` 检查两轮写回、零值、原始公式、其他内外层 ZIP 条目和去嵌入后的再次投影。公式创建、删除、改指向、source-free 公式编写及通用 workbook/formula topology 仍未完成；本轮只记录结构与往返证据。

**本轮补充（2026-09-09）：** `errorBars.valueType: "custom"` 与 `plus/minus: { values, formatCode? }` 已表达每个数据点的本地正负误差，数组与系列点数严格对齐，保留零值；`both` 要求双侧，`plus/minus` 只接受对应侧。普通柱/条/线及 categorical combo 支持 authored、导入后数组/格式修改、单双侧切换、custom/scalar 互换、删除和重建；formatCode 接受 string grammar token。`PpjCustomErrorDataAuthorsAndEditsLiteralOwners` 检查原生 numLit 数值与格式、节点顺序、去嵌入后的投影，以及目标 owner 以外的 XML/ZIP 字节保护。公式型 numRef 保留引用身份，未知或不规则缓存不开放 analytics 编辑；固定公式的同步范围见上方后续增量。

**本轮补充（2026-09-09）：** `data.series[].errorBars` 补齐导入后的对象增删：recognized bar/column、line 及 categorical combo 系列可以省略字段删除，再新增或修改 fixed-value、percentage、standard-deviation、standard-error 对象。`PpjErrorBarsLifecycleRoundTrips` 检查四类图表的删除→新增零值→切换类型/方向/端帽/直接 stroke→删除→重新添加，每步验证原生节点顺序、去嵌入投影、目标 owner 以外的 ChartML 与 ZIP 条目保持不变。重复或非法节点、未知 extension 保留原始 ChartPart；普通误差线不能覆盖未投影的公式型 custom plus/minus 数据；本地数组由后续增量补齐。SVG 预览对误差线明确报告 `chart-error-bars-not-rendered`，本轮记录原生结构和往返证据，不声称视觉或 PowerPoint 宿主验收。

**本轮补充（2026-09-09）：** `data.series[].trendlines` 补齐导入后的列表增删：recognized bar/column、line 及 categorical combo 中的这些系列，现在可插入、排序、替换、删除单项，或通过空数组/省略字段清空后重新添加。`PpjTrendlineListSupportsSourceBoundLifecycle` 检查实际 ChartML 节点、去嵌入后的回投影、趋势线在误差线/数据之前的顺序，以及其他系列、误差线、轴和 ZIP 条目保持不变。未知 label/extension、重复或非法 type 节点保留原始 ChartPart，不开放 analytics 编辑；error-bar 增删由下述后续增量补齐。

**本轮补充（2026-09-09）：** 对数轴进入 PPJ `yAxis.logBase`，散点/气泡数值 X 轴、组合图次值轴及雷达 `spokeAxis.logBase` 共用同一字段。底数接受 2–1000 的数值或 size grammar token，删除字段恢复线性轴；显式 min/max 必须为正数。`PpjAxisLogBaseAuthorsEditsAndRemoves` 验证 authored → 去嵌入投影 → 修改/删除/重新添加 → 再投影，并检查原生节点顺序、ChartPart 内仅 logBase 变化以及其他 ZIP 条目字节不变。非法参数和重复/额外属性/子节点的原生 owner 另有拒绝及保留实验。记录的是原生结构和往返证据。

**本轮补充（2026-09-09）：** ChartSpace 的 `c:roundedCorners` 已接入 PPJ `chart.roundedCorners`。普通图和组合图保留缺失、显式 false、true 三态；语义哈希及内容比较均保留字段存在性，修复 false→删除被当成无改动的问题。`PpjChartRoundedCornersPreservesPresenceAcrossEdits` 分别验证两类图的 false→删除→false→true→删除，每步检查原生 XML、去嵌入后的投影，以及除目标 ChartPart 外所有部件字节不变；重新添加时保持 roundedCorners 在 style 前的顺序。`PpjChartRoundedCornersRejectsAmbiguousNativeOwners` 验证非法值、重复节点、额外属性和子节点不获得 setChartPlot，未修改时保留原始 ChartPart。

**优先级：P0；状态：部分完成。**

普通轴标题、轴与数据标签的 `numberFormat` grammar token 已纳入同一 bounded ChartPart profile；它只解决声明值的可复用解析，不扩展为 source-linked workbook 格式同步。

**当前进度：** PPJ 有 16 类 chartType，包括 Kimi 的 13 类 series 以及 doughnut、combo；已覆盖 trendline、error bar、labels、bubble scale、axis lines/arrowheads、radar spoke、waterfall、heatmap、candlestick、treemap、sunburst、sankey 等有限 authored/native/vector profile。本轮新增 chart `dataset/encoding/dataFilter/seriesDefaults` 归一化、ChartSpace `style.frame` 的 fill/line/shadow profile，以及 `strRef/numRef` 的安全本地公式引用投影与 ChartPart-only source-bound 编辑；opaque native chart 的安全 category cache profile 已覆盖 bar/line/area/pie/doughnut/radar 及 bounded column/line/area combo，并可同步唯一 embedded worksheet cell；scatter/bubble 的 X/Y/size cache channel 也在独立公式、缓存和工作簿单元格证明后开放双 footprint 编辑；普通 x/y/secondary 轴的有限位置、数值边界、方向、标签可见性、轴标题、numberFormat 和轴/网格线字段现在同时支持 authored/source-bound grammar token，并保持目标 ChartPart-only 回写；现已补齐 `setChartSeriesAnalytics` 的窄路径：recognized bar/column、line（含 categorical combo 系列）的 trendline 列表可 source-bound 增删、排序及替换；scalar 与 literal custom error-bar 对象可新增、删除、重新添加及替换逐点数据、参数、显示标记和直接 stroke，并通过二次投影恢复；完整 ChartML、workbook footprint 和更复杂组合仍缺。

**本轮补充：** `showBubbleSize` 已进入 PPJ schema、protobuf、authored compiler、projector 和 source-bound `setChartLabels` 回写；plot、series 默认值和 sparse point override 都保留字段 presence，`true` 只对 bubble chart/series 放行，显式 `false` 也能原样回投影。最小回归已验证 plot/series 两个 `c:showBubbleSize` 节点、去嵌入后的二次投影、只改 ChartPart 的 false 编辑和 Office 2021 Open XML 校验。

**本轮补充：** Kimi 的 `showLeaderLines` 已进入 PPJ schema、protobuf、authored compiler、projector 和 source-bound `setChartLabels` 回写；字段 presence 在 plot 与 series 默认值两个原生 scope 保留，显式 `false` 也能原样回投影。Office 2021 拒绝把它放进单点 `c:dLbl`，因此 point scope 保持 fail-closed；leader-line 的几何、布局和效果图仍归 source-owned。

**本轮补充：** Kimi 的 `ChartAxisConfig.minorUnit` 已进入 PPJ schema、protobuf、authored compiler、projector 和 source-bound `setChartAxis` 回写；标准 native value axis 保留正数 `minorUnit` 的字段 presence，authoring 可写入 `<c:minorUnit>`，去嵌入后的 source-bound 修改只改目标 ChartPart 并由二次投影恢复。category axis、generated vector axis 和非正/不规则值保持 fail-closed。

**本轮补充：** Kimi 的 `ChartAxisConfig.tickLabelPosition` 已进入 PPJ schema、protobuf、authored compiler、projector 和 source-bound `setChartAxis` 回写；普通 native x/y/secondary 轴保留 `<c:tickLblPos>` 的 `nextTo`、`high`、`low`、`none` 四种有界值，去嵌入后的 source-bound 修改只改目标 ChartPart 并由二次投影恢复。`none` 与旧 `tickLabelsVisible=false` 兼容别名保持同一隐藏语义；自定义位置、雷达简写和不规则轴拓扑继续 fail-closed。

**本轮补充：** Kimi 的 `ChartConfig.displayBlanksAs` 已进入 PPJ schema、protobuf、authored compiler、projector 和 source-bound `setChartPlot` 回写；chart-level 的 `<c:dispBlanksAs>` 保留 `zero`、`gap`、`span` 三种有界值，也支持 string grammar token。普通 native 与 combo ChartPart 的字段 presence 会在 authored → 去嵌入 → source-bound 编辑 → 二次投影链路中保留；最小回归覆盖 combo 图表且只改目标 ChartPart。`displayBlanksAs` 不改 series/cache/workbook 数据，vector fallback、复杂/不规则 ChartML 继续 fail-closed。

**本轮补充：** Kimi 的 `ChartLegendConfig.position` 现在补齐 `topRight`；PPJ 的 `style.legend` 将其映射为 `<c:legendPos val="tr">`，普通 native/combo ChartPart 的 authored 与 source-bound 回投影均保留该值。最小回归覆盖 authored → 去嵌入 → source-bound 改为 `left` → 二次投影；vector fallback 仍只接受其已声明的 `none`/`right` 有界范围。

**本轮补充：** Kimi 的 `ChartLegendConfig.overlay` 现在补齐为 PPJ `style.legendOverlay`；普通 native/combo ChartPart 以 `<c:legend><c:overlay val="1|0"/></c:legend>` 保留字段 presence，authored 与 source-bound 回投影均保留显式 `true`/`false`。vector fallback 没有同等图表容器语义，继续 fail-closed。

**本轮补充：** Kimi 的 `ChartLegendConfig.fill` 现在补齐为 PPJ `style.legendFill`；普通 native/combo ChartPart 以 `<c:legend><c:spPr>` 承载有界的 none、solid RGB 或 literal gradient fill，authored 与 source-bound 回投影均保留该字段。source-bound 只改目标 ChartPart，vector fallback 没有同等图表容器语义，继续 fail-closed。

**本轮补充（2026-09-07）：** ChartML 原生 `c:overlap` 现在补齐为 PPJ `style.overlap`；普通柱图和 combo 中的 column plot 保留 `-100..100` 的有符号整数，authored 与 source-bound 回投影均恢复该值。source-bound 只改目标 ChartPart；非柱图和 vector fallback 不伪造 overlap 语义，继续 fail-closed。`PpjChartOverlapAuthorAndEditSourceChart` 覆盖 combo 与普通 column 的 authored → 去嵌入 → 负值 source-bound 修改 → 二次投影。

**本轮补充（2026-09-07）：** ChartML 原生 `c:varyColors` 现在补齐为 PPJ `style.varyColors` 的柱/column 窄字段；普通 bar/column 和含 column plot 的 categorical combo 保留 presence-aware true/false，authored 与 source-bound 回投影均恢复该值。source-bound 只改目标 ChartPart；line chart 继续使用既有 `line_options`，vector fallback 和无 column 的 combo 不伪造该语义，继续 fail-closed。`PpjChartVaryColorsAuthorAndEditSourceChart` 覆盖 combo 与普通 column 的 authored → 去嵌入 → false source-bound 修改 → 二次投影。

**本轮补充（2026-09-07）：** 同一 `scatterOptions.varyColors` gap 也补齐到普通 scatter；直接 `c:scatterChart/c:varyColors` 的 true/false 会进入 PPJ `style.varyColors`，并完成 authored → 去嵌入投影 → source-bound 回写 → 二次投影，只改目标 ChartPart。bubble/radar、数值 combo、vector fallback 和不规则 ChartML 不扩张该语义，继续 fail-closed。`PpjChartVaryColorsAuthorAndEditSourceChart` 现在覆盖 combo、普通 column 与普通 scatter。

**本轮补充（2026-09-07）：** Kimi 的 `areaOptions.varyColors` gap 也补齐到普通 area；直接 `c:areaChart/c:varyColors` 的 true/false 会进入 PPJ `style.varyColors`，并完成 authored → 去嵌入投影 → source-bound 回写 → 二次投影，只改目标 ChartPart。area combo 无 column plot、vector fallback 和不规则 ChartML 不扩张该语义，继续 fail-closed。`PpjChartVaryColorsAuthorAndEditSourceChart` 现在覆盖 combo、普通 column、普通 area 与普通 scatter。

**本轮补充（2026-09-07）：** ChartML 原生 `c:majorTickMark` 现在补齐为 PPJ `chartAxis.majorTickMark`；普通 category/value/secondary 轴保留 presence-aware `cross`、`in`、`out`、`none`，authored 与 source-bound 回投影恢复该值。source-bound 只改目标 ChartPart；雷达 spoke、vector fallback 和不规则 axis topology 不伪造该语义，继续 fail-closed。`PpjChartAxisMajorTickMarksAuthorAndEditSourceChart` 覆盖 x/y 轴的 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮补充（2026-09-07）：** ChartML 原生 `c:minorTickMark` 现在补齐为 PPJ `chartAxis.minorTickMark`；普通 category/value/secondary 轴保留 presence-aware `cross`、`in`、`out`、`none`，authored 与 source-bound 回投影恢复该值。source-bound 只改目标 ChartPart；雷达 spoke、vector fallback 和不规则 axis topology 不伪造该语义，继续 fail-closed。`PpjChartAxisMinorTickMarksAuthorAndEditSourceChart` 覆盖 x/y 轴的 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮补充（2026-09-07）：** Kimi 的 `ChartAxisConfig.minorGridlines` 现在补齐为 PPJ `chartAxis.minorGridLine`；普通 native category/value/secondary 轴保留布尔可见性与 bounded direct RGB stroke，authored 与 source-bound 回投影恢复该字段。source-bound 只改目标 ChartPart；雷达 spoke、vector fallback 和不支持的 style graph 不伪造该语义，继续 fail-closed。`PpjChartAxisMinorGridlinesAuthorAndEditSourceChart` 覆盖 x/y 轴的 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮补充（2026-09-07）：** Kimi 的 `ChartAxisConfig.position` 现在补齐为 PPJ `chartAxis.position`；普通 native category/value/secondary 轴以 `<c:axPos>` 保留水平轴 `bottom`/`top` 和垂直轴 `left`/`right`，authored 与 source-bound 回投影恢复该位置。source-bound 只改目标 ChartPart，并将匹配 native 默认的位置规范化为省略字段；稳定轴 ID 的 categorical combo 也能安全识别，雷达 spoke、vector fallback 和无法证明轴组归属的第三方 combo 继续 fail-closed。`PpjChartAxisPositionsAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 改回默认位置 → 二次投影。

**本轮补充（2026-09-07）：** Kimi 的 `ChartLegendConfig.line` 现在补齐为 PPJ `style.legendLine`；普通 native/combo ChartPart 以 `c:legend/c:spPr/a:ln` 承载有界的直接 RGB、宽度和预置虚线，authored 与 source-bound 回投影恢复该轮廓。source-bound 只改目标 ChartPart；vector fallback、主题/效果线图和复杂 ChartML 继续 fail-closed。`PpjChartLegendLineAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮补充（2026-09-07）：** Kimi 的 `ChartDataLabelsConfig.line` 现在补齐为 PPJ `style.dataLabels.line`；普通 native/combo ChartPart 以 plot-level `c:dLbls/c:spPr/a:ln` 承载有界的直接 RGB、宽度和预置虚线，authored 与 source-bound 回投影恢复该轮廓。source-bound 只改目标 ChartPart；series/point label 线、主题/效果图和复杂 shape-properties graph 继续 fail-closed。`PpjChartDataLabelLineAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮补充（2026-09-07）：** Kimi 的 `ChartDataLabelsConfig.fill` 现在补齐为 PPJ `style.dataLabels.fill`；普通 native/combo ChartPart 以 plot-level `c:dLbls/c:spPr` 承载有界的 none、solid RGB 或 literal gradient paint，authored 与 source-bound 回投影恢复该填充，并可与已支持的 label line 共存。source-bound 只改目标 ChartPart；图片、主题/效果图和复杂 shape-properties graph 继续 fail-closed。`PpjChartDataLabelFillAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮补充（2026-09-07）：** Kimi 的 `ChartTextStyleConfig.alignment` 现在补齐为 PPJ `chartTextStyle.alignment`；普通 native/combo ChartPart 以标题、图例、数据标签、轴标题和刻度标签的直接 `a:pPr/@algn` 承载 `left`、`center`、`right`、`justify` 四种有界值，authored 与 source-bound 回投影恢复该对齐。source-bound 只改目标 ChartPart；未知 rich-text 段落图和 vector fallback 不伪造该语义，继续 fail-closed。`PpjChartTextAlignmentAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮补充（2026-09-07）：** Kimi 的 `ChartTextStyleConfig.fill` 现在补齐为 PPJ `chartTextStyle.fill`；普通 native/combo ChartPart 以标题、图例、数据标签、轴标题和刻度标签的直接 `a:noFill`、`a:solidFill` 或 literal `a:gradFill` 承载有界文字填充，authored 与 source-bound 回投影恢复该 paint。solid 输入继续兼容既有 `color` 投影，none/gradient 保留 `fill` 形状；source-bound 只改目标 ChartPart，图片、主题/效果图和未知 rich-text paint 图继续 fail-closed。`PpjChartTextFillAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮点级数据标签填充增量（2026-09-07）：** Kimi 的 `dataLabelOverrides[].fill` 现在补齐为 PPJ `series[].dataLabels.points[].fill`；普通 native/combo ChartPart 的稀疏 `c:dLbl` 以直接 `c:spPr` 承载 bounded none、solid RGB 或 literal gradient paint，authored 与 source-bound 回投影恢复点级填充。source-bound 只改目标 ChartPart；point label line、图片/主题/效果图和复杂 label graph 继续 fail-closed。`PpjChartPointDataLabelFillAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮点级数据标签轮廓增量（2026-09-07）：** Kimi 的 `dataLabelOverrides[].line` 现在补齐为 PPJ `series[].dataLabels.points[].line`；普通 native/combo ChartPart 的稀疏 `c:dLbl` 以直接 `c:spPr/a:ln` 承载 bounded RGB、宽度和预置虚线，并可与点级 fill 共存。source-bound 只改目标 ChartPart；自定义 line/effect graph、leader-line 几何和 vector fallback 继续 fail-closed。`PpjChartPointDataLabelLineAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮点级数据标签文本增量（2026-09-07）：** Kimi 的 `dataLabelOverrides[].text` 现在补齐为 PPJ `series[].dataLabels.points[].text`；普通 native/combo ChartPart 的稀疏 `c:dLbl` 以单段单 run 的 literal `c:tx/c:rich` 承载文本，并可与点级 fill/line 共存。source-bound 只改目标 ChartPart；公式引用、多段/多 run rich text、布局/效果图和 vector fallback 继续 fail-closed。`PpjChartPointDataLabelTextAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮补充（2026-09-07）：** Kimi 的 `scatterOptions.style` 现在补齐为 PPJ `style.scatterStyle`；普通 PPTX `c:scatterChart` 的 `line`、`lineWithMarkers`、`marker`、`smooth`、`smoothWithMarkers` 映射到 native `line`、`lineMarker`、`marker`、`smooth`、`smoothMarker`，authored 与 source-bound `setChartPlot` 回投影均只改目标 ChartPart。XLSX 仍保持 marker-only 可编辑 profile；numeric combo、3D/extension/effect/不规则 ChartML 继续 fail-closed。

**本轮补充（2026-09-07）：** Kimi 的 `LinearSeriesBase.nullHandling` 现在补齐为 PPJ `data.series[].nullHandling`；普通 line、area、radar ChartPart 将 `zero`、`gap`、`connect` 映射到 chart-level `<c:dispBlanksAs>` 的 `zero`、`gap`、`span`，并在去嵌入后的 source-bound `setChartPlot` 编辑中恢复该 series 字段。一个图内已声明的 line/area/radar series 值要求一致，省略项跟随图表模式；combo、scatter/bubble、candlestick、streamgraph/vector fallback、extension 和不规则 ChartML 继续 fail-closed。`PpjChartNullHandlingAuthorProjectAndEditOnlyChartPart` 覆盖 authored → 去嵌入 → source-bound 改值 → 二次投影且只改目标 ChartPart。

**仍缺：** 通用 dataset/encode 的完整 13 类组合、图表容器 frame 的共享/外部 image relationship 治理与完整 effect graph、任意 ChartML extension、3D charts、stock/特殊 radar 变体、复杂 combo/多坐标轴、公式/链接工作簿的通用同步编辑（当前仅安全本地公式引用的 ChartPart-only profile）、图表 build 动画、自动标签布局和完整 number format/effect graph。

**待实现与验收：** K-02 已有 bounded frame profile，K-03 已有 dataset/encode、安全本地公式引用、`setChartSeriesStyle` 的有限 series marker/stroke 窄路径和 `setChartSeriesAnalytics` 的 trendline 列表增删改与 scalar/literal custom error-bar 对象增删改；`NativeChartDataLeavesCoverCircularAndRadarCategoryPlots`、`NativeChartDataLeavesCoverCategoricalComboPlots` 与 `NativeChartDataLeavesCoverScatterAndBubbleNumericChannels` 已补足单 family category、bounded column/line/area combo 及 scatter/bubble numeric channel 识别，`PpjSourceBoundNativeChartDataLeafEditsCacheAndWorkbookAndReprojects` 证明安全点的双 footprint 编辑，`PpjGapProfilesCompileAndReproject` 证明 combo line series marker/stroke 的增改删只改 ChartPart 并可二次投影。下一阶段是为每个 chart family 区分 native ChartPart、vector lowering 和 opaque projection，并在确认 embedded workbook 所有权后证明更完整的 ChartPart/cache/worksheet footprint；error-bar 的固定本地公式型 plus/minus 已支持缓存/工作簿同步；复杂 trendline label/effect graph 和 workbook/formula topology 变更仍缺。不能从任意 DrawingML 形状猜 chart semantics，也不能把只改公式字符串误报为 workbook 同步。

### F-08 SmartArt、DiagramML 和四部件关系

**优先级：P1；状态：部分完成（8 类布局 profile 已交付）。**

**当前进度：** authored `smartArt` 支持 1–64 节点、8 类有限布局、自定义 `office-kit/smartart-definition/v1`、gap/column/reverse 等 operator、四个标准 diagram parts、缓存 drawing；source-bound 支持已证明的 text/connection/frame/style identity 编辑和显式 `detachToShapes`。对 OfficeKit 自有 `picture` 布局，缓存 drawing 中每个节点的单一嵌入 blip 会按媒体内容哈希投影为 PPJ `nodes[].asset`，其 canonical `a:blipFill` paint 还可投影为 `nodes[].image`（`fit: stretch|tile`、crop、opacity）；在节点数量、ID、文本、连接和布局保持不变时，`setSmartArtImage` 和 `setSmartArtImagePaint` 可分别替换已有 asset 或缓存 paint，重建受控 diagram/cache/media 闭包并通过二次投影恢复。

**仍缺：** 任意 DiagramML operator、formula-backed layout、quick style/color 复杂继承、picture 节点的新增/删除/多 blip 或效果化缓存、未知 content/relationship graph、共享/外部 closure、完整 SmartArt animation 和从缓存几何反推语义。第三方图只有在同一有限闭包、单一嵌入图片关系和稳定节点身份都被证明时才可走该 asset profile，否则继续 typed read-only/opaque。

**待实现与验收：** 扩展 operator 前先为每种 operator 建立定义资产和 copy-on-write 规则；source-bound 只编辑被 capability 明确授权的 part，picture asset/paint 替换还必须证明缓存 drawing 的嵌套关系源、媒体新增/删除和旧闭包清理；未知图、复杂 blipFill（多 blip、effect、外链或共享关系）继续 typed read-only 或 opaque；detach 必须提示语义损失并单独验证。

### F-09 Transition、Timing、Animation、Morph 和触发器

**优先级：P1；状态：部分完成（有限 timing profile 已交付）。**

**当前进度：** PPJ 有 23 类 transition、5 类对象动画效果、入口/退出/强调、顺序、click/previous、delay/stagger、paragraph/chart build 和 Morph pair；本轮增加 typed `timingGraph` sugar 并正规化到现有动画数组，compact `animations[]` 与 `timing.nodes[]` 都保留规范化 `trigger` 字段，结构化 round-trip 已有证据。有限 profile 已覆盖 linear/ease-in/ease-out/ease-in-out、repeat 1–8、autoReverse，以及可验证的 trigger→start 映射；authored media 另有 `playback.trigger` 的 onClick/onSlideStart lowering。完整 `p:timing` closure、motion path、媒体复杂 timing 和条件 trigger 仍缺。

**仍缺：** 完整 `p:timing` 树、parallel/sequence container、条件触发、shape trigger closure、motion path、custom effect、imported media/audio timing、SmartArt/Chart 全量 build、宿主 UI 行为；authored media 目前只有初始 onClick/onSlideStart 条件，不代表完整媒体播放图；repeat/easing/autoReverse 目前只有有限 `p:cTn` 映射，不代表全量 PowerPoint 语义。

**待实现与验收：** 采用 K-07 的 typed timing graph；每个 trigger/sequence 都要有稳定 ID 和 closure；未知 timing 不降级为 fade；review 结果分别记录“结构正确”“宿主识别”“真实播放”。

### F-10 Audio、Video、媒体关系和播放行为

**优先级：P2；状态：部分完成（source-bound clone 与 authored 基础 profile 已交付，编辑仍缺）。**

**当前进度：** PPJ authored 有 audio/video asset、poster、start/end、loop、mute，以及 bounded `playback.trigger`；source-bound 支持一个 canonical embedded MP4 clone leaf，并保持媒体关系和海报图。

**仍缺：** 音频/视频添加和 payload 编辑、trim/volume/fade、字幕/章节/替代文本轨道、imported/复杂播放触发、控件、视频背景、跨宿主播放验证；当前 `playback.trigger` 只承诺 authored `onClick`/`onSlideStart` lowering。

**待实现与验收：** 先把 media metadata 与 media timing 分开；payload replacement 必须绑定 MIME/hash/relationship closure；不执行转码、不下载远程资源、不宣称播放通过；任何共享或外部媒体图保持 opaque。

### F-11 OLE、嵌入 Office 文件、3D、Ink、Custom XML、ActiveX 和宏

**优先级：P2；状态：只覆盖少数安全 profile。**

**当前进度：** 对唯一绑定的内部 XLSX/DOCX OLE 有 payload extraction/replacement profile；支持 canonical MP4 和 InkML 的 unchanged clone leaf；PPJ 有 `ole`、`opaque` 等保留类型。

**仍缺：** 任意 OLE 激活、preview regeneration、链接更新、Excel/Word 宿主行为、3D model、Ink stroke 编辑、任意 Custom XML、ActiveX、VBA/macros、表单控件和嵌入对象 UI 状态。

**待实现与验收：** 每种新嵌入格式单独建立 content type、关系、所有权、预览和宿主边界；不把 payload 替换扩展成 arbitrary OLE editor；未知 content part 继续保留并拒绝危险操作。

### F-12 Notes、Comments、Sections、Custom Shows、Notes Master 和 Handouts

**优先级：P1；状态：部分完成（PPJ authored 与 source-bound 的有界 profile 已交付）。**

**当前进度：** PPJ 有 page notes、presentation comments、sections、custom shows；已有受控的备注、评论、页面顺序和 section/custom-show 变更路径。本轮把已有 NativeAOT Office 2021 modern-comment profile 接入 PPJ：导入后的 `kind: "modern"` 根评论和 direct replies 会保留 element/text-range anchor、parent、位置、作者/日期身份和精确的 `active|resolved|closed` 状态；`replaceText` 与 `setCommentStatus` 分开发放，source-bound 只改 modern comments part，并由 `PpjModernCommentsProjectAndReprojectTextAndStatus` 完成去嵌入 PPJ→编译→changed-part→二次投影回读。source-free PPJ 现在也可 author 有界 modern root/direct-reply：以 PPJ comment ID 和 author 名称确定性生成 Office GUID/person 元数据，将目标元素映射到单一 Drawing moniker（含 text-range），并由 `PpjSourceFreeModernCommentsAuthorAndReproject` 验证 Office 2021 包、锚点、线程和二次投影；这不会把复杂线程或关系图降级成普通 legacy comment。

**仍缺：** notes master/handout master 的完整语义、复杂评论线程/mentions/reactions/身份扩展、跨页评论锚点、rich notes、页面动作和 custom-show 触发器的完整 PowerPoint 行为；nested/branched/connected modern graph、共享/外部关系和现代评论的更多任务字段仍不开放。

**待实现与验收：** 将评论/备注正文、身份、锚点、线程拓扑分开建模；source-bound 只改现有已证明文本和状态叶；线程或锚点拓扑变化没有 capability 时保持原文。当前 modern PPJ 窄路径已证明 authored GUID/person/anchor 生成、status/text 的独立 capability、comments-part-only footprint 和二次投影恢复；后续只在有明确数据/宿主需求时扩展复杂 thread 或 notes/handout 闭包。

### F-13 Hyperlink、Action Setting、Macro Button 和交互导航

**优先级：P1；状态：部分完成。**

**当前进度：** PPJ rich text 支持有限 hyperlink；shape-producing 的 text/shape/line/icon/placeholder 另有 typed `action`（click）和 `hoverAction`（mouse-over），可 authored 编译并从 PPTX 投影 URI、内部 slide、custom show（含 returnToSlide）和有限 action verb；URI scheme 与目标引用在 validator/codec 入口校验，未知元素类型会 fail closed。对已识别为安全形状的 source-bound click/hover action，编译器现在可以只替换或移除 URL、页内页、custom show 或有限 verb，并清理不再引用的超链接/页跳转关系；changed-part 只包含目标 slide XML 和必要的 slide relationship part，二次投影恢复新目标。authored `media.playback.trigger` 现在补上 `onClick`/`onSlideStart` 两值字段，后者写入现有 `p:cMediaNode` 的 immediate start condition；媒体 payload、poster 和关系仍由媒体 codec 所有。transition 支持 click/after timing；元素有 accessibility metadata 和稳定 ID。

**仍缺：** 声音、运行宏、媒体触发之外的完整条件/书签触发、媒体 timing graph、字幕/章节和完整安全 URL policy；当前媒体触发只覆盖 authored `onClick`/`onSlideStart` 两值，不开放 imported timing 图。现有 action profile 不处理宏或任意 action graph。未知 sound/macro/extension action 继续拒绝，不会被重建成普通 hyperlink。

**待实现与验收：** 已完成安全 click/hover-action 的 source-bound 关系闭包：区分 URL、slide、custom show 和有限 verb，替换/移除时只修改目标 shape 与必要关系，并用二次投影验证目标恢复；authored media trigger 另有 `playback.trigger` 最小 profile，`onSlideStart` 以 `p:cMediaNode` 的 `delay="0"` 证明 lowering，省略字段保持旧 click-start。继续区分 imported timing、macro/sound 等高风险图并保持 fail closed。当前 focused acceptance 覆盖四种 authored click 目标、hover URL 的 authored/投影、source-bound click/hover URL 替换、关系清理和一个非法 image action 的拒绝；媒体 profile 由 authored media XML 与 embedded PPJ recovery 回归覆盖。

### F-14 Accessibility、Reading Order 和可访问性验证

**优先级：P1；状态：部分完成（基础 metadata 子能力已交付，完整语义仍缺）。**

**当前进度：** PPJ/asset 有 title、description、rights 等 metadata；元素有 accessibility 和稳定 ID；`pages[].readingOrder` 已成为显式、完整 direct-element permutation，普通 `group.readingOrder` 也成为显式、完整 direct-child permutation，authored/native projection 会保持两者序列。对 source-bound 页面，安全的完整 permutation 现在映射为现有 shape-tree 的 z-order；对 recognized group，完整 child permutation 映射为 group 内部 shape-tree 的 z-order；编译后只改对应 SlidePart，二次投影以元素语义顺序证明回写生效；不存在另一个独立的 PowerPoint reading-order XML owner。JS review 新增 reading-order 完整性、图片/图表/表格/SmartArt/media/OLE alternative-text 和 decorative 决策的机器检查，并列出需要人工/宿主检查的边界。对导入媒体，若 picture 形状的 `cNvPr` residual profile 可证明安全，PPJ 保持 `opaque/nativeKind: "media"` 但会投影 accessibility 并颁发 `setAccessibility`；source-bound 只改 title、description、decorative，保留媒体关系、poster、播放 action 和 timing 图。

**仍缺：** SmartArt 内部朗读顺序、装饰图标的全量宿主语义、表格/图表的宿主可访问性语义、PowerPoint Accessibility Checker 等价验证和人类意图判断；媒体的 payload、captions、播放触发/完整 timing 与宿主行为仍不开放；对存在不可移动 shape-tree 子节点、组合拓扑或未知宿主顺序语义的页面仍保持 fail closed。已识别的普通 group direct children 和 residual-profile 媒体 metadata 不再属于这两个窄字段缺口，但 group 的 `readingOrder` 只表示并回写本地 shape-tree 顺序，不声称额外的宿主朗读语义。

**待实现与验收：** 继续将 reading order 作为显式可审计序列；在本 bounded profile 中，页面只有明确完整 permutation 且 shape-tree 全部可移动时才用 z-order 作为物理 owner，普通 group 则要求 direct children 的完整 permutation，并把它回写为 group 内部 child z-order；两者都不对未知宿主语义猜测。为图片、图表、表格、SmartArt 分别定义必填/可选 metadata；机器检查与人工检查分开，不声称 WCAG/PPT conformance 已完成。当前 focused acceptance 覆盖页面和 group 的显式 permutation、source-bound z-order 回写和二次投影、非法 permutation、缺少 alternative text 和 decorative-with-text warning；媒体 gap experiment 另外覆盖 opaque media 的 capability→SlidePart metadata 写回→二次投影，并证明 playback closure 不变。

### F-15 Theme、Font Scheme、Color Transform 和 Effect Style

**优先级：P1；状态：部分完成。**

**当前进度：** PPJ 有 theme colors、fonts、直接 RGB/theme color、named styles、渐变、部分阴影和文本效果；grammar color token 可以作为 authored color fallback，并按声明顺序执行有界 `tint`（向白）再 `shade`（向黑）变换，native 输出和二次投影已有证据。

本轮又为六个 formal 文字标量增加直接 `design.theme.textStyle` authored fallback，其中包含 `text.fontFamilyEastAsia`；theme 命中与缺字段回落到 default token 均保持在 native run 并可由二次投影恢复，但不扩展为原生 `theme1.xml` 样式图编辑。

本轮新增显式 `design.theme.fontScheme.major/minor` authored 字段；它们优先于 `design.fonts` 的位置约定，写入真实 `a:majorFont`/`a:minorFont`。其中 `major`、`minor` 分别提供各自的 Latin 槽位；可选 `majorEastAsia`、`minorEastAsia` 只覆盖对应 role 的 East Asian 槽位，`majorComplexScript`、`minorComplexScript` 只覆盖对应 role 的 complex-script 槽位，缺省时回落到各自 role family；缺少整个显式 scheme 时仍保留旧回退。source-bound 不把主题 XML 猜成可写能力。

本轮再增加显式 `design.theme.accentColors.accent1..accent6` authored 字段；它们优先于 `design.theme.colors` 的旧顺序约定，写入真实 `a:accent1Color`..`a:accent6Color`。该字段只声明六个 RGB/RGBA accent role，八位 `#RRGGBBAA` 后缀只降低为对应 `a:alpha`，source-bound 仍拒绝新主题 owner，不扩展为完整 `theme1.xml` 编辑。

本轮再补一个有限的主题颜色变换字段 `design.theme.accentTransforms.accent1..accent6`；每个角色可声明 `tint` 和/或 `shade`，按 PPJ 的 0..1 分数写入对应 `a:srgbClr` 下的真实 `a:tint`/`a:shade`，并按“先 tint、后 shade”保留顺序。它只处理 authored 六个 accent role，未声明的角色保持旧输出；不把完整的 transform/effect scheme、继承 cascade 或 imported `theme1.xml` 变成可写对象。

在此基础上再补 `lumMod` 和 `lumOff` 两个有限变换字段；前者使用 0..1 分数，后者使用 -1..1 的有符号分数，写入 `a:lumMod`/`a:lumOff`，并排在 tint/shade 后面。该增量只扩展已有 authored accent owner，保留变换字段的独立 presence 和顺序，不宣称完成其它颜色变换族或宿主颜色管理。

再补 `alphaMod` 和 `alphaOff` 两个有限 alpha 变换字段；前者使用 0..1 分数，后者使用 -1..1 的有符号分数，写入 `a:alphaMod`/`a:alphaOff`。如果 accent 基础色带有 RGBA 后缀，原生 `a:alpha` 仍作为绝对 alpha owner 保留在前面；这组字段不把绝对 alpha 与相对变换合并。

再补 `satMod` 和 `satOff` 两个有限饱和度变换字段；前者使用 0..1 分数，后者使用 -1..1 的有符号分数，写入 `a:satMod`/`a:satOff`，并保留已有颜色变换的独立 presence。该字段仍只属于 authored 六角色 accent owner。

再补 `redMod`/`redOff`、`greenMod`/`greenOff` 和 `blueMod`/`blueOff` 六个有限 RGB 通道变换字段；三个 `*Mod` 使用 0..1 分数，三个 `*Off` 使用 -1..1 的有符号分数，分别写入对应的 `a:redMod`/`a:redOff`、`a:greenMod`/`a:greenOff`、`a:blueMod`/`a:blueOff`。每个字段保持独立 presence，未声明的通道不补写；该增量仍只属于 authored 六角色 accent owner。

再补五个无数值的离散颜色变换字段 `gray`、`comp`、`inv`、`gamma` 和 `invGamma`；它们只接受显式 `true`，分别写入对应的 `a:gray`、`a:comp`、`a:inv`、`a:gamma` 和 `a:invGamma` 空子节点。省略的操作不补写，`false` 和空对象在 schema/编译入口失败；该增量仍只属于 authored 六角色 accent owner。

再补 `hueMod` 和 `hueOff` 两个有限色相变换字段；前者使用 0..1 分数写入 `a:hueMod`，后者使用 -360..360 度数并转换为 DrawingML 的 1/60000 度单位写入 `a:hueOff`。两个字段保持独立 presence，未声明的一方不补写；该增量仍只属于 authored 六角色 accent owner。

本轮继续增加显式 `design.theme.colorRoles.dark1/light1/dark2/light2/hyperlink/followedHyperlink` authored 字段；它们写入真实 `a:clrScheme` 的六个剩余 bounded role，缺省值沿用 OfficeKit 原有 clean-room 默认。该字段接受六位 RGB 或八位 `#RRGGBBAA`，只将 alpha 写入 `a:alpha`，不推断 transforms、effect scheme、继承或 imported `theme1.xml` 的可写 owner；source-bound 仍 fail closed。

文字字体映射再补了一个窄的复杂脚本 owner：显式 `fontFamilyComplexScript` 直接写入/读取/编辑 `a:cs/@typeface`，覆盖 authored 普通/图表文本、imported native leaf 和 source-bound 单叶 token-splice；这只缩小语言/脚本映射残差，不引入宿主字体回退或完整 theme font cascade。

**仍缺：** 完整 `theme1.xml` color/font/effect scheme、更多颜色变换、继承后的 effect style、更完整的逐脚本字体回退、字体嵌入/替换、复杂语言字体映射的其余组合、WordArt 和宿主字体回退差异；当前 color roles、alpha、tint/shade、major/minor East Asian slots、major/minor complex-script slots 与复杂脚本 direct owner 都只覆盖 PPJ 的有限 authored/native lowering，不会改写任意 imported theme XML。

**待实现与验收：** 在已有 grammar color transform、六角色 accent 色板、六角色 dark/light/hyperlink 色板（含 authored alpha）和 font scheme 之上建立更完整的 theme transform canonical profile；所有直接颜色和 theme token 的覆盖顺序可验证；未知 theme/effect 子树必须 source-preserved；导出时记录实际字体和缺失字体 warning。当前 focused acceptance 覆盖 token fallback、tint/shade、两组六角色颜色、六位 RGB 与八位 RGBA 的 native alpha、major/minor 及其 East Asian/complex-script slot 结果和非法值拒绝。

### F-16 通用布局、遮挡、重排和响应式页面

**优先级：P1；状态：部分完成（只读 review 已交付）。**

**当前进度：** 有 frame、layout、placeholder、component expansion、AutoFit 和 review warning；`reviewPpjArtifact` 现在能报告原始 frame、保守 visual bounds、越界、z-order、遮挡和确定性文本溢出估算，且能区分轴对齐 frame、旋转/阴影与有限箭头头型的 visual-bounds 检测。组件 repeat 支持 authored `grid` 的列数、列间距和行间距、有限 flow/anchor，以及 horizontal/vertical `layout.weights` weighted stack，输出普通 PPJ frame 并已二次投影。canvas 修改不会自动重排，review 不产生写操作。真实宿主文本测量、跨对象约束、通用 solver 和显式 apply 仍未完成。

**仍缺：** PowerPoint Designer 式自动排版、约束传播、智能对齐、组内自适应、文本/图片/图表联动、遮挡修复、响应式换页和内容密度重排。

**待实现与验收：** K-08 的只读 visual-bounds、确定性文本估算、有限箭头轮廓代理、越界/重叠建议和 authored component grid/flow/anchor/weighted stack 已交付；后续再做真实字体测量、有限跨对象约束求解和显式 `layoutApply` 操作。source-bound 默认不自动移动对象；每个修复都要能回放、审计和二次导入。

### F-17 文档级属性、保护、签名、加密和发布设置

**优先级：P2；状态：大部分未建模或只读。**

**当前进度：** PPJ 有 source/provenance/hash/revision；部分页面和评论元数据可编辑。

**仍缺：** 完整 core/app/custom properties、文档保护、密码/加密、数字签名、IRM、嵌入字体、打印/手册/演讲者视图设置、广播/发布配置和所有 revision history 语义。

**待实现与验收：** 每个文档级属性单独定义 source hash、签名失效和保存策略；涉及安全边界的功能不放进普通 PPJ authored compiler；无能力时只读报告或 fail closed。

### F-18 PowerPoint 宿主行为与跨平台验收（独立证据，不计入 PPJ gap）

**优先级：P2；状态：证据边界，不是 PPJ 待实现项。**

**当前进度：** 已有 NativeAOT/Office Open XML 校验、模型 SVG/PNG review、LibreOffice/Poppler 和 Keynote 的部分结构或渲染证据；动画有 authored/second-import 和 Keynote 观察。

**尚未提供：** Windows PowerPoint 打开、编辑、播放、Morph、动画触发、字体、图表、SmartArt 和媒体的真实宿主证据；这不计入 PPJ/codec gap，也不从完成度扣分。

**后续条件：** 只有用户明确开启 Windows lane 后，才记录对应宿主证据；本阶段不运行 Windows 验收，也不用 macOS/Keynote 结果替代 Windows PowerPoint 证据。没有这条证据时，语义能力仍按 PPJ/codec 的结构和 source-bound 证据计分。

## 4. 实施顺序与里程碑

### M0：基线冻结（当前）

**状态：已完成。**

- 固定 `main@9b86939a`；
- 固定 PPJ schema、PPJ reference、coverage 和 Kimi `pptd.md` 的路径；
- 将 authored、source-bound、opaque、host evidence 分开统计；
- 不进行 Windows 验收，不启动完整发布门禁。

### M1：先补 Kimi 的直接原语差距（进行中）

**状态：部分完成。** K-01、K-02 已完成各自 bounded source-bound profile（literal/高层 points line 与 solid/gradient/image chart frame 均有 authored、changed-part/二次投影证据），K-03 已完成有限 dataset/encoding 的 parser/归一化、13 类单 family authored 回归、bounded column/line/area 混合，并同时具备高层本地公式 ChartPart-only 和 opaque native chart numeric leaves（ChartPart cache + embedded worksheet）两条 source-bound 证据；后者的 category profile 已覆盖 bar/line/area/pie/doughnut/radar 及 bounded column/line/area combo，scatter/bubble 的 X/Y/size 也有独立通道回归，并补有圆形/雷达/组合/散点/气泡识别证据；K-03 的更广泛跨 family 组合、轴数组无损恢复、共享/外链 workbook 和完整 footprint 仍未闭合。

**目标：K-01、K-02、K-03。**

交付顺序：

1. 独立 `freeCurve/linePath`；
2. chart container frame；
3. dataset/encode/filter/seriesDefaults 兼容层。

这一阶段完成后，PPJ 才能说在 Kimi 的七类元素上具有较接近的一对一语言映射；在此之前，`connector` 不能被描述成 Kimi `line` 的完全替代。这里的“一对一映射”只针对元素和核心字段，不表示模板、布局、宿主行为或完整 PowerPoint 已完成。

### M2：图层合成与图片语义（部分启动）

**状态：部分完成。** K-04 已有 compositing schema、诊断和 shape/image 的 normal opacity authored lowering；非 normal blend、isolation、clip stack 仍 fail closed。图片已有 imagePolicy 校验和 authored 槽位投影窄路径，但模板槽位搜索/替换和 effect closure 尚未闭合。

**目标：K-04、F-05。**

- 先完成统一 opacity/crop/mask 顺序和 clip closure；
- 再按 native 支持度实现 blend/isolation；
- 对不支持的组合输出显式 unsupported/lossy 记录；
- 增加 image slot 和 image-fill source-bound profile。

### M3：模板语法与布局审查（部分启动）

**状态：部分完成。** K-05 已交付 typed grammar 声明/校验、只读 token/predicate evaluator、grammar color 的有限 tint/shade authored lowering 和形状/图表/表格有界 style precedence；K-06 已交付 slot policy 校验、typed `image.crop` 组件绑定、schema-v3 图片示例槽位、replacement plan、显式 PPJ apply transaction、source-free fixture 以及确定性 imported/source-bound fixture 的 plan→NativeAOT compile/project 事务边界，K-08 已交付带旋转/阴影 visual-bounds 的只读 layout review 与 authored component grid/flow/anchor/weighted stack；完整 style cascade/writeback、跨导入焦点裁剪推导、共享/歧义/外部关系治理、source-owned accessibility、文本测量、跨对象约束、solver 和显式 layout apply 仍待实现。

**目标：K-05、K-06、K-08、F-16。**

- 先把 `designGrammar` 中可验证的规则结构化，并把已实现的 token transform 与 diagnostics 分离；
- 交付只读 layout/occlusion report；
- 再在现有 authored grid/flow/anchor repeat 之外逐步开放 bounded weighted stack，再处理完整约束布局；
- imported source 只允许显式 layout operation，不自动重排。

### M4：完整 PowerPoint 的高价值语义

**目标：F-02、F-03、F-04、F-07、F-08、F-09、F-12、F-13、F-14、F-15。**

原则是按对象闭包推进，而不是按 XML 标签数量推进：

1. 先补一个完整对象闭包的 authored + second import；
2. 再为同一闭包增加 source-bound 的局部叶子；
3. 最后才扩展到共享关系、嵌入工作簿或复杂宿主行为。

### M5：媒体、嵌入对象和文档级能力

**目标：F-10、F-11、F-17。**

这部分风险最高，且经常牵涉宿主、外部数据、安全和签名。除非有真实用户场景，不把它们混入普通 PPJ 核心语言；优先做 inspect/preserve/typed replacement，避免声称拥有完整编辑器。

### M6：宿主证据（暂不启动，不计入 PPJ gap）

**目标：F-18；F-18 不是 PPJ 待实现项。**

此阶段只有在明确安排 Windows PowerPoint 环境后才启动。当前所有“宿主未验收”只保留为证据标记，不从 PPJ/codec 完成度扣分，也不用其他渲染器冒充 Windows 证据。

## 5. 每个新原语的 Definition of Done

完成状态分三层，不能混写：

1. **结构/PPJ 完成**：schema、wire、compiler、projection 和 authored/second-import 闭环已经成立；这只能说明语言和结构化编解码达到目标 profile。
2. **source-bound 完成**：在已有第三方包上，能够证明 source hash、所有权、依赖 closure、changed-part footprint、非目标 residual 和二次导入恢复；没有这组证据，只能写“结构完成”或“部分完成”。
3. **宿主完成**：目标 Office 宿主实际打开、编辑、播放或交互通过；模型渲染、LibreOffice、Poppler、Keynote 的结果只能作为结构/视觉辅助证据。Windows PowerPoint 作为单独验收 lane，本轮不启动。

一个新字段或 C# primitive 只有在声明了属于哪一层，并满足对应的以下条件后，才能从“待实现”升级：

1. **Schema**：PPJ schema、wire contract、字段约束和预算已更新。
2. **Compiler**：NativeAOT writer 有确定的 canonical lowering；不接受任意 OOXML、XPath、JS 或网络指令。
3. **Projection**：导入能恢复 PPJ，或者明确返回 typed source-bound/opaque 以及阻塞原因。
4. **Edit capability**：source-bound 变更有 source hash、closure、所有权和 changed-part footprint。
5. **Round-trip**：至少有 authored compile → import → controlled edit（如适用）→ second import。
6. **Package proof**：证明非目标 part、关系、媒体和嵌入内容未被静默改写。
7. **Visual review**：有模型渲染或目标宿主渲染；无法验证的宿主明确标记。
8. **Docs/Skill**：更新 PPJ reference、`docs/coverage.md`、相关 Skill 和 capability registry。
9. **Failure behavior**：未知 topology、共享 closure、扩展 effect、公式路径或超预算输入 fail closed，不拍平。

### 5.1 验收记录的最小格式

每个 K/F 条目完成时，都要在实现记录或覆盖文档中留下同一组字段，避免“测试跑过”却无法判断测到了哪一层：

| 记录项 | 必须写清的内容 |
| --- | --- |
| Fixture | 输入 PPJ、输入 PPTX（如为 source-bound）、版本/哈希和是否含共享关系、外部链接或未知扩展 |
| Expected package scope | 允许修改的 part、relationship、media、embedding 路径；例如图表至少列出 `ppt/charts/*` 与依赖 workbook，SmartArt 列出四个 diagram part 和 rels |
| Positive assertions | schema 校验、authoring 编译、导入投影、受控编辑和二次导入后要恢复的字段/ID |
| Negative assertions | 公式路径、未知扩展、共享 closure、外部链接、超预算输入应返回的 capability 或 fail-closed 原因 |
| Residual proof | 非目标 part、关系、媒体和嵌入内容的 canonical hash/residual；不能只比较整个 ZIP 的偶然字节序 |
| Visual evidence | 至少一个模型或目标渲染结果、尺寸/裁剪/遮挡/文字溢出检查；图片预览不能替代语义检查 |
| Host lane | 结构、模型/Keynote/LibreOffice 观察和 Windows PowerPoint 验收分别记账；未做的 lane 明确写“未验收” |

高风险对象的最小 package scope 约定如下，后续实现可以收窄，不能无故扩大：

| 对象 | 最小 fixture 族 | 重点 package scope |
| --- | --- | --- |
| 文本/形状/布局 | 纯 authored + 一个含 master/layout/theme 的 imported slide | `ppt/slides/slide*.xml`、`ppt/slideMasters/*`、`ppt/slideLayouts/*`、`ppt/theme/*` |
| 图片/背景/mask | raster、SVG fallback、image fill、预设和自定义 mask 各一例 | `ppt/slides/*`、`ppt/media/*`、相关 rels；页面/母版背景另列 closure |
| 图表 | Kimi 13 类兼容集 + doughnut/combo + 一个嵌入 workbook | `ppt/charts/chart*.xml`、`ppt/embeddings/*`、chart/slide rels |
| SmartArt | authored 8 布局 + 一个 source-bound 四部件样本 | `ppt/diagrams/*`、diagram rels、slide graphic frame 和缓存 drawing |
| 动画/过渡 | 单对象、段落构建、click/after、Morph 各一例 | slide `p:timing`、transition、相关 shape ID；未知 `p:extLst` 必须保留 |
| 媒体/OLE | embedded MP4、XLSX/DOCX OLE 各一例 | `ppt/media/*` 或 `ppt/embeddings/*`、preview、content type 和所有 rels |
| 交互/可访问性 | URL、内部 slide、custom show、reading order 各一例 | hyperlink/action rels、稳定 ID、显式 reading-order 字段和 review report |

## 6. 非目标与决策边界

- 不以整页图片证明 PPJ 可编辑；图片只能作为 preview/calibration/evidence。
- 不把 codec 已有的内部能力直接计入 PPJ；必须有 schema、编译、投影和 Agent 可达路径。
- 不为了“支持更多文件”而把未知 DiagramML、ChartML、timing、OLE 或 Custom XML 猜成普通 shape/group。
- 不自动修改 source-bound 页面的坐标、z-order、布局或背景层；所有重排都是显式操作。
- 不在本轮执行 Windows PowerPoint 验收；不把 Keynote、LibreOffice 或模型渲染升级为 Windows 证据。
- 不以完整 npm/release 发布作为上述语言缺口的前置条件；只在需要时做 office PPJ 范围内的验证。

## 7. 跟踪表

以下表是下一轮实现时的入口，完成一项后必须回填证据链接和实际 changed-part 范围；证据入口按 1.5 索引，若没有对应证据，不得只因为代码或 schema 已出现就升级状态。

| ID | 目标 | 当前状态 | 下一步 | 优先级 |
| --- | --- | --- | --- | --- |
| K-01 | 独立自由曲线 | 有界 profile 完成（typed path + Kimi points；literal 直线、单段 quadratic/cubic 和 5–128 点高阶 Bézier 多段 smooth 均可在最多 48 点且可稳定反解的边界保留 compact projection） | 超过 48 点或量化后无法稳定反解的 compact sugar、复杂 geometry/group transform 仍由 F-04 覆盖 | P0 |
| K-02 | 图表容器 frame | 有界 profile 完成（solid/gradient/image fill、line/shadow）；image 的 crop/stretch/tile/opacity，以及单 owner image 素材替换的 ChartPart `.rels`/media 清理和二次投影均有证据 | 共享/外部关系治理、复杂 effect graph 与 F-07 完整 ChartML | P0 |
| K-03 | dataset/encode/filter | 部分完成（13 类单 family authored、bounded column/line/area 混合、candlestick overlay 与一/二级轴数组已覆盖；有限 `xAxisIndex/yAxisIndex` 0/1 会随 canonical series 回投影；无公式/无高级 series 拓扑的原生分类图及简单 scatter/bubble 会在保留旧字段的同时投影 canonical `dataset + encoding`；opaque native chart 安全点可通过 `nativeRef.leaves[].chartDataCategory/chartDataValue/chartDataXValue/chartDataYValue/chartDataBubbleSize` 按通道双 footprint 编辑，category profile 覆盖 bar/line/area/pie/doughnut/radar 及 bounded column/line/area combo，scatter/bubble 覆盖 X/Y/size；其中 `chartDataCategory` 只接受 `c:cat/c:strRef` 与直接 inline/string worksheet cell，shared-string、公式、外链和不规则 workbook 保持 opaque；普通 x/y/secondary 轴的有限数值边界和线/标签布尔字段新增 `size`/`boolean` token authored/source-bound 回投影；本轮新增 `setChartSeriesStyle`，允许 line/scatter/radar marker 与非 scatter direct series stroke 的增改删，保持 ChartPart-only footprint） | 更广泛跨 family series、完整轴数组原始形状与 ChartPart/workbook footprint round-trip；共享/外链 workbook 和高层公式仍不做通用同步 | P0 |
| K-04 | opacity/blend/clip/isolation | 部分完成（normal opacity authored，含 shape/image/line/icon/placeholder/connector 的 `compositing.opacity`，connector 以 `stroke.opacity` 为规范投影 owner，image/line opacity token、source-bound solid background opacity；image 单个非 inverse preset 或 bounded custom `compositing.clipStack` 已落到 picture mask，并在去快照投影为 `image.mask`，保留调整值或 path commands） | native support matrix、multi-level/超出 custom codec 的 clip closure 与显式 lossy boundary | P0 |
| K-05 | 结构化 style grammar | 部分完成（evaluator + color tint/shade + 文字/形状/image/line stroke/shadow/chart 标题基本属性与颜色 token authored lowering；图表常用 legend/legendOverlay/legendFill/stacking/bubble 枚举、gap/切片/孔径/气泡整数、普通轴的数值边界/方向/标签、`numberFormat` 与 axis/grid line、以及 line options 也接受 typed grammar token 并做有限值回检；chart/radar data-label number format 同样支持 `string` token；文字 run、形状/图表（含图表常用轴/图例/标签/绘图区字段）的首个来源 precedence 与显式 chart 文字样式嵌套字段浅合并、表格有界 style precedence；本轮新增 `design.styles.image`/`image.styleRef`，对 fit/crop/focus/opacity/border/shadow 做字段级 precedence，并覆盖 authored/source-bound 图片 paint） | 扩展 token 字段与类型，落地完整 cascade、样式写回和 diagnostics/apply 分离 | P1 |
| K-06 | 模板 image slots | 部分完成（slot policy + authored projection + typed `image.crop`/`image.focus` component binding + schema-v3 示例槽位/replacement plan + pure PPJ apply transaction + source-free 与确定性 imported/source-bound fixture 的本机 NativeAOT compile/project/语义回投影） | 跨导入焦点语义、共享/歧义/外部关系和 source-owned accessibility 仍不自动推导；继续扩大 fixture 族 | P1 |
| K-07 | timing graph | 部分完成（repeat/easing/trigger bounded profile） | trigger closure/keyframe/media timing 与 closure 编辑 | P1 |
| K-08 | layout/occlusion review | 部分完成（review + authored grid/flow/anchor repeat + 有限箭头 visual bounds） | 真实字体测量、mask 轮廓、solver，再做显式 apply | P1 |
| F-02 | master/layout/theme | 部分完成（direct background + owner-local master/layout/slide placeholder frame/text/有限 rotation-flip source-bound profile；slide→layout→master 有界 effective frame 与 slide-owner 物化已交付；直接嵌入 master/layout image background 的 crop/opacity 与单 owner 替换、关系/媒体闭包已交付；owner-local rotation/flip 的存在、显式零/false、删除三态已闭合） | 多继承、无法唯一匹配的 inherited placeholder、复杂 inherited/non-text placeholder content profile 和 theme cascade | P0 |
| F-03 | fields/WordArt/完整文本 | 部分完成（rich text 子集、静态 typed field display source-bound 编辑、bounded automatic `slidenum` cache refresh 与 `author`/`datetime`/`datetimeFigureOut`/`datetime1`/`datetime2`/`datetime3`/`datetime4`/`datetime5`/`datetime6`/`datetime7`/`datetime8`/`datetime9`/`datetime10`/`datetime11`/`datetime12`/`datetime13`/`uaqdatetime1`/`uaqdatetime2`/`uaqdatetime3`/`uaqdatetime4`/`uaqdatetime5`/`uaqdatetime6`/`uaqdatetime7` host-refreshable cache、复杂脚本 direct `a:cs` owner、formal `text.language` owner、source-free authored text glow/inner shadow/reflection/soft edge、direct rich-text run 及 paragraph `defaultText` glow-radius/color/theme-color/opacity/shadow-blur/distance/direction/alignment/color/theme-color/opacity/rotate-with-shape/inner-shadow-blur-distance-direction-color-theme-color-alpha/reflection-blur-distance-start-opacity-end-opacity-direction/soft-edge radius source-bound leaves 已交付） | 未覆盖的自动字段求值、其它 automatic field、WordArt、defaultText 其它 effect leaf、复杂 effect 与完整语言/脚本字体回退组合 | P0 |
| F-04 | full geometry/group transform | 部分完成（literal custom-path source-bound profile、recognized preset-shape 完整/partial literal sibling `presetGeometryAdjustment` leaves、普通 group 外层 frame leaf、bounded childFrame leaves、连接线对象锚点及有符号端点已交付） | guide/handle 分级、多路径/闭合语义、preset partial/formula/extension graph 的整体编辑、child-space descendant 联动/rescale；K-01 literal/points freeCurve bounded residual proof 已闭合 | P0 |
| F-05 | full image/effect graph | 部分完成（literal custom-mask、picture preset identity/complete adjustments、ordinary image 上单个非 inverse preset 或 bounded custom `compositing.clipStack` 到 picture mask 的 authored profile、picture border/shadow、ordinary-shape shadow、source-free authored shape/image/line/picture `glow`、`innerShadow`、`reflection`、`softEdge`、imported/source-bound 普通 shape/line/picture direct glow、inner shadow、reflection 与 soft edge、solid background opacity、direct embedded image background crop/opacity 与单 owner replacement 的 source-bound profile 已交付，覆盖 slide/master/layout owner 的关系/媒体闭包；简单 partial/formula mask graph 的 literal sibling leaf 也已交付） | compositing native matrix、multi-level/超出 custom codec 的 clip/effect closure、shape/image 3-D、非简单 partial/formula/unknown mask adjustments、复杂 mask/effect graph | P0 |
| F-06 | complete table semantics | 部分完成（固定矩形表格 + `setTableStyle` 五 flag + `setTableGeometry` 固定列/行宽高 + direct cell fill/border + 单一直接嵌入 cell image replacement + 同一 SlidePart 内共享 image relationship 的 copy-on-write + 固定拓扑多 run 文本替换 + 跨段落样式一致多 run text-style + 固定段落/直接纯文本、field 或 break mixed-run text body + 嵌入式 picture bullet 及其 direct font/color/size source-bound profile + structured `text.style` bounded direct bodyPr layout，含 rotation/overflow/upright/normalAutoFit 百分比） | 段落/列表/复杂字段/效果级 inheritance、显式删除、外部/跨 owner image relationship、picture bullet 未知/效果化 marker 子图、自动 reflow/自动换行/extension profile | P1 |
| F-07 | full ChartML | 部分完成 | 已补齐普通/combo ChartPart 的 legend placement/overlay、有界 legend `c:spPr` fill、标题 `titlePlacement`、plot-area `c:spPr/a:ln`、bounded `styleIndex` 与柱/组合图 `c:overlap`；标题的 `none`/`aboveChart`/`centeredOverlay` 已接到 PPJ/schema/wire/native ChartML，并有 authored → 去嵌入 → source-bound → 二次投影证据；plot-area line 以直接 RGB/宽度/预置虚线的 bounded profile 与 plot-area fill 共存，style index 以直接 `c:style/@val` 承载 1–48 内建样式且 source-bound 只改目标 ChartPart；仍缺 3D/extension/复杂 combo、共享/外链公式与通用 workbook 同步；本地安全公式及 opaque native chart 双 footprint（category `c:strRef` 直接 inline/string cell、bar/line/area/pie/doughnut/radar 以及 bounded column/line/area combo 的 value、scatter/bubble 的 X/Y/size）都只是受控窄路径；opaque native chart 的直接 `c:strLit`/`c:numLit` literal cache 另有 ChartPart-only leaf 编辑，保留无关 `externalData` 而不触碰 workbook | P0 |
| F-08 | full DiagramML | 部分完成（8 类布局 profile + OfficeKit 自有 picture 节点 asset 替换已交付） | operator/profile 扩展；picture asset 的新增/删除、共享/外部关系和复杂缓存继续保持 fail-closed | P1 |
| F-09 | full timing | 部分完成（有限 timing graph/repeat/easing profile，以及 authored media `playback.trigger` 的 onClick/onSlideStart 已交付） | typed trigger closure、keyframe/motion、imported media timing | P1 |
| F-10 | media editing/playback | 部分完成（基础/clone profile 与 authored `playback.trigger` 已交付） | payload/timing/captions、复杂/imported playback boundary | P2 |
| F-11 | OLE/3D/Ink/custom XML | 少数 profile | 按 content type 单独立项 | P2 |
| F-12 | notes/comments/custom shows | 有界完成（modern comments authored + source-bound root/direct-reply） | master/handout、复杂 thread/action topology | P1 |
| F-13 | actions/interactivity | 部分完成（typed click/hover action authored/projection，以及 authored media `playback.trigger` 的 onClick/onSlideStart） | media/macro 的复杂/imported trigger、完整 action graph | P1 |
| F-14 | accessibility/reading order | 部分完成（explicit page/group readingOrder + machine review + six typed-owner/source-bound media accessibility） | SmartArt/表格宿主 semantics、Checker 等价验证；安全 direct-element 和普通 group-child z-order 回写已完成；七类 owner（含保持 opaque 的 imported media）的 `setAccessibility` 已闭合 PPJ capability→SlidePart 写回→二次投影，剩余为 SmartArt 内部朗读顺序、宿主语义与 Checker 等价 | P1 |
| F-15 | theme/effect/font scheme | 部分完成（bounded grammar tint/shade + authored `design.theme.fontScheme.major/minor/majorEastAsia/minorEastAsia/majorComplexScript/minorComplexScript` + `design.theme.accentColors`/`accentTransforms`（含 `tint`/`shade`/`lumMod`/`lumOff`/`alphaMod`/`alphaOff`/`satMod`/`satOff`/`redMod`/`redOff`/`greenMod`/`greenOff`/`blueMod`/`blueOff`/`hueMod`/`hueOff`/`gray`/`comp`/`inv`/`gamma`/`invGamma`）/`colorRoles`，含有限 alpha，以及复杂脚本 direct `a:cs` 字体 owner） | transform/inheritance/fallback/完整语言脚本字体回退 | P1 |
| F-16 | auto layout/occlusion | 部分完成（review + authored component grid/flow/anchor + 有限箭头 visual bounds） | 真实字体测量、mask 轮廓、solver 与显式 layoutApply | P1 |
| F-17 | document/security settings | 未建模/只读 | 按真实需求拆分 | P2 |
| F-18 | Windows host evidence（独立证据，不计入 gap） | 未验收 | 明确环境后另开验收 lane；不从 PPJ 完成度扣分 | P2 |

## 7.1 本轮最小验收记录（2026-09-03）

本轮后续增量（2026-09-06）：`textField.automatic` 已进入 PPJ/schema/wire，并在当前有界 profile 中支持 `type: "slidenum"`、`type: "author"`、默认 `type: "datetime"`、`type: "datetimeFigureOut"`、`type: "datetime1"`、`type: "datetime2"`、`type: "datetime3"`、`type: "datetime4"`、`type: "datetime5"`、`type: "datetime6"`、`type: "datetime7"`、`type: "datetime8"`、`type: "datetime9"`、`type: "datetime10"`、`type: "datetime11"`、`type: "datetime12"` 与 `type: "datetime13"`、`type: "uaqdatetime1"`、`type: "uaqdatetime2"`、`type: "uaqdatetime3"`、`type: "uaqdatetime4"`、`type: "uaqdatetime5"`、`type: "uaqdatetime6"`、`type: "uaqdatetime7"`。`slidenum` 的 authoring/source-bound 缓存按所属 slide 序号刷新；`author`、`datetime`、`datetimeFigureOut`、`datetime1`、`datetime2`、`datetime3`、`datetime4`、`datetime5`、`datetime6`、`datetime7`、`datetime8`、`datetime9`、`datetime10`、`datetime11`、`datetime12`、`datetime13`、`uaqdatetime1`、`uaqdatetime2`、`uaqdatetime3`、`uaqdatetime4`、`uaqdatetime5`、`uaqdatetime6` 与 `uaqdatetime7` 写入并回读 host-refreshable 标记及缓存文本，其中 `datetimeFigureOut` 请求宿主使用 MM/DD/YYYY，`datetime1` 请求宿主使用 MM/DD/YYYY，`datetime2` 请求宿主使用 Day, Month DD, YYYY，`datetime3` 请求宿主使用 DD Month YYYY，`datetime4` 请求宿主使用 Month DD, YYYY，`datetime5` 请求宿主使用 DD-Mon-YY，`datetime6` 请求宿主使用 Month YY，`datetime7` 请求宿主使用 Mon-YY，`datetime8` 请求宿主使用 MM/DD/YYYY hh:mm AM/PM，`datetime9` 请求宿主使用 MM/DD/YYYY hh:mm:ss AM/PM，`datetime10` 请求宿主使用 hh:mm，`datetime11` 请求宿主使用 hh:mm:ss，`datetime12` 请求宿主使用 hh:mm AM/PM，`datetime13` 请求宿主使用 hh:mm:ss AM/PM，`uaqdatetime1` 请求宿主使用 DD/MM/YYYY（Umm al‑Qura 日历），`uaqdatetime2` 请求宿主使用 Day, DD Month, YYYY（Umm al‑Qura 日历），`uaqdatetime3` 请求宿主使用 DD Month, YYYY（Umm al‑Qura 日历），`uaqdatetime4` 请求宿主使用 DD/MM/YY（Umm al‑Qura 日历），`uaqdatetime5` 请求宿主使用 YYYY-DD-MM（Umm al‑Qura 日历），`uaqdatetime6` 请求宿主使用 DD-Month-YY（Umm al‑Qura 日历），`uaqdatetime7` 请求宿主使用 DD Month YYYY（Umm al‑Qura 日历）；OfficeKit 不伪造本机时间或文档作者。PPJ 二次投影恢复 `automatic: true`；其它未覆盖的自动字段和完整字段指令图继续 fail closed。

本轮增量补充：opaque native chart 的直接 ChartML `c:strLit`/`c:numLit` cache 已纳入 `chartDataCategory`/数值 native leaf。它只在所属 ChartPart 做 token-splice，保留可能存在的 `externalData` 关系，不绑定或改写 embedded workbook；二次投影会恢复 literal 值。含共享字符串、公式、外链或不规则 workbook 的引用路径仍保持 source-owned/fail closed。随后新增 OfficeKit 自有 picture SmartArt 节点 asset 替换回归，focused 回归从 40 项增至 41 项；Windows PowerPoint 仍是独立未验证证据，不进入计数。

本轮后续增量（2026-09-07）：Kimi `ChartConfig.titlePlacement` 已接入 PPJ/schema/wire/native ChartML。`none` 表示无标题，`aboveChart` 表示无 `c:overlay` 或显式 false，`centeredOverlay` 写入 `c:title/c:overlay val="1"`；普通/combo ChartPart 完成 authored → 去嵌入 → source-bound 从 centeredOverlay 改为 aboveChart → 二次投影，且只改目标 ChartPart。无标题统一回投影为 `none`；公式、超出 bounded ChartML 标题拓扑和不支持的图表族仍 source-owned/fail closed。

本轮后续增量（2026-09-05）：`PpjSourceBoundImageShadowColorEditsLeafAndReprojects` 覆盖图片 direct outer-shadow RGB `srgbClr/@val` 的 source-bound native leaf，证明单字段 token splice、SlidePart-only changed-part、其它 shadow 属性和非目标 ZIP part 保留、Open XML 校验及二次投影。

本轮后续增量（2026-09-05）：`textBodyColumnGapEmu` 绑定文本 shape 的直接 `p:sp/p:txBody/a:bodyPr/@spcCol`，在已有 `textBoxStyle.columnGap` 语义之外提供独立 native leaf；`PpjSourceBoundTextBodyColumnGapEditsLeafAndReprojects` 验证 `25400` → `50800` 的单属性 token splice、文本与其它 body 内容保留、SlidePart-only changed-part、Open XML 和二次投影。

本轮后续增量（2026-09-05）：`textBodyRotationDegrees` 绑定文本 shape 的直接 `p:sp/p:txBody/a:bodyPr/@rot`，在已有 `textBoxStyle.rotation` 语义之外提供独立 native leaf；`PpjSourceBoundTextBodyRotationEditsLeafAndReprojects` 验证 `5400000` → `8100000` 的单属性 token splice、文本与其它 ZIP part 保留、SlidePart-only changed-part、Open XML 和二次投影。

本轮后续增量（2026-09-05）：`textBodyVerticalOverflow` 绑定文本 shape 的直接 `p:sp/p:txBody/a:bodyPr/@vertOverflow`，在已有 `textBoxStyle.verticalOverflow` 语义之外提供独立 native leaf；`PpjSourceBoundTextBodyVerticalOverflowEditsLeafAndReprojects` 验证 `ellipsis` → `clip` 的单属性 token splice、文本与其它 ZIP part 保留、SlidePart-only changed-part、Open XML 和二次投影。

本轮后续增量（2026-09-05）：`textBodyHorizontalOverflow` 绑定文本 shape 的直接 `p:sp/p:txBody/a:bodyPr/@horzOverflow`，在已有 `textBoxStyle.horizontalOverflow` 语义之外提供独立 native leaf；`PpjSourceBoundTextBodyHorizontalOverflowEditsLeafAndReprojects` 验证 `overflow` → `clip` 的单属性 token splice、文本与其它 ZIP part 保留、SlidePart-only changed-part、Open XML 和二次投影。

本轮后续增量（2026-09-05）：`textBodyUpright` 绑定文本 shape 的直接 `p:sp/p:txBody/a:bodyPr/@upright`，在已有 `textBoxStyle.upright` 语义之外提供独立 native leaf；`PpjSourceBoundTextBodyUprightEditsLeafAndReprojects` 验证 `1` → `0` 的单属性 token splice、文本与其它 ZIP part 保留、SlidePart-only changed-part、Open XML 和二次投影。
本轮后续增量（2026-09-05）：`textBodyAnchorCenter` 绑定文本 shape 的直接 `p:sp/p:txBody/a:bodyPr/@anchorCtr`，在已有 `textBoxStyle.anchorCenter` 语义之外提供独立 native leaf；`PpjSourceBoundTextBodyAnchorCenterEditsLeafAndReprojects` 验证 authored → 去嵌入投影 → `1` → `0` 单属性 token splice、文本与其它 ZIP part 保留、SlidePart-only changed-part、Open XML 和二次投影，缺失属性不发 leaf。

本轮后续增量（2026-09-05）：`textBodyForceAntiAlias` 绑定文本 shape 的直接 `p:sp/p:txBody/a:bodyPr/@forceAA`，新增 additive `PresentationTextBodyProperties.force_anti_alias`/`textBoxStyle.forceAntiAlias` 语义和独立 native leaf；`PpjSourceBoundTextBodyForceAntiAliasEditsLeafAndReprojects` 验证 authored → 去嵌入投影 → 单 `@forceAA` token 修改 → SlidePart-only changed-part → Open XML → 二次投影，缺失属性不发 leaf。
本轮后续增量（2026-09-05）：`textBodySpaceFirstLastParagraph` 绑定文本 shape 的直接 `p:sp/p:txBody/a:bodyPr/@spcFirstLastPara`，新增 additive `PresentationTextBodyProperties.space_first_last_paragraph`/`textBoxStyle.spaceFirstLastParagraph` 语义和独立 native leaf；`PpjSourceBoundTextBodySpaceFirstLastParagraphEditsLeafAndReprojects` 验证 authored → 去嵌入投影 → 单 `@spcFirstLastPara` token 修改 → SlidePart-only changed-part → Open XML → 二次投影，缺失属性不发 leaf。
本轮后续增量（2026-09-05）：`textBodyCompatibleLineSpacing` 绑定文本 shape 的直接 `p:sp/p:txBody/a:bodyPr/@compatLnSpc`，新增 additive `PresentationTextBodyProperties.compatible_line_spacing`/`textBoxStyle.compatibleLineSpacing` 语义和独立 native leaf；`PpjSourceBoundTextBodyCompatibleLineSpacingEditsLeafAndReprojects` 验证 authored → 去嵌入投影 → 单 `@compatLnSpc` token 修改 → SlidePart-only changed-part → Open XML → 二次投影，缺失属性不发 leaf。
本轮后续增量（2026-09-05）：`textBodyFromWordArt` 绑定文本 shape 的直接 `p:sp/p:txBody/a:bodyPr/@fromWordArt`，新增 additive `PresentationTextBodyProperties.from_word_art`/`textBoxStyle.fromWordArt` 语义和独立 native leaf；`PpjSourceBoundTextBodyFromWordArtEditsLeafAndReprojects` 验证 authored → 去嵌入投影 → 单 `@fromWordArt` token 修改 → SlidePart-only changed-part → Open XML → 二次投影，缺失属性不发 leaf。
本轮后续增量（2026-09-06）：`textBodyFlatTextZ` 绑定文本 shape 的直接 `p:sp/p:txBody/a:bodyPr/a:flatTx/@z`，新增 additive `PresentationTextBodyProperties.flat_text_z`/`textBoxStyle.flatTextZ` 语义和独立 native leaf；`PpjSourceBoundTextFlatTextZEditsLeafAndReprojects` 验证 authored → 去嵌入投影 → 单 `@z` 数字 token 修改 → SlidePart-only changed-part → Open XML → 二次投影，缺失、重复、额外属性、子节点、非规范和越界 `flatTx` 不发 leaf；周围 3D scene 仍保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dContourRgb` 绑定普通 shape 的严格直接 `p:sp/p:spPr/a:sp3d/a:contourClr/a:srgbClr/@val`，新增 additive `PresentationShape.shape_3d_contour_rgb`/`shape3dContourRgb` native leaf；`PpjSourceBoundShape3dContourRgbLeafEditsAndReprojects` 验证 `#336699` → `#6699CC` 的单 RGB token splice、SlidePart-only changed-part、Open XML 和二次投影，bevel/scene/transform/扩展与其它 color model 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dContourRgb` 同一叶子扩展到图片的严格直接 `p:pic/p:spPr/a:sp3d/a:contourClr/a:srgbClr/@val`，新增 additive `PresentationImage.shape_3d_contour_rgb`；最小 source-bound 实验验证 `#336699` → `#6699CC` 只替换图片 `srgbClr/@val`、只改变 owning SlidePart，并保留图片关系、crop/mask/effect、其它 3-D 状态与其它 ZIP part，复杂/不明确颜色 owner 继续 source-owned。
本轮后续增量（2026-09-06）：`shape3dContourColorScheme` 绑定普通 shape 与 picture 的严格直接 `p:sp|p:pic/p:spPr/a:sp3d/a:contourClr/a:schemeClr/@val`；新增 additive `PresentationShape.shape_3d_contour_color_scheme` 与 `PresentationImage.shape_3d_contour_color_scheme`/`shape3dContourColorScheme` native leaf，最小 source-bound 实验验证 `accent1` → `accent2` 只替换 contour theme token、只改变 owning SlidePart，并保留 3-D root 状态、图片关系/crop/mask/effect 与其它 ZIP part；带 transforms、额外子节点、未知 theme token 或复杂 3-D topology 仍 source-owned。
本轮后续增量（2026-09-06）：`shape3dExtrusionColorScheme` 绑定普通 shape 与 picture 的严格直接 `p:sp|p:pic/p:spPr/a:sp3d/a:extrusionClr/a:schemeClr/@val`；新增 additive `PresentationShape.shape_3d_extrusion_color_scheme` 与 `PresentationImage.shape_3d_extrusion_color_scheme`/`shape3dExtrusionColorScheme` native leaf，最小 source-bound 实验验证 `accent1` → `accent2` 只替换 extrusion theme token、只改变 owning SlidePart，并保留 3-D root 状态、图片关系/crop/mask/effect 与其它 ZIP part；带 transforms、额外子节点、未知 theme token 或复杂 3-D topology 仍 source-owned。
本轮后续增量（2026-09-06）：`shape3dExtrusionRgb` 绑定普通 shape 的严格直接 `p:sp/p:spPr/a:sp3d/a:extrusionClr/a:srgbClr/@val`，新增 additive `PresentationShape.shape_3d_extrusion_rgb`/`shape3dExtrusionRgb` native leaf；`PpjSourceBoundShape3dExtrusionRgbLeafEditsAndReprojects` 验证 `#663399` → `#9966CC` 的单 RGB token splice、SlidePart-only changed-part、Open XML 和二次投影，bevel/scene/transform/扩展与其它 color model 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dExtrusionRgb` 同一叶子扩展到图片的严格直接 `p:pic/p:spPr/a:sp3d/a:extrusionClr/a:srgbClr/@val`，复用 additive `PresentationImage.shape_3d_extrusion_rgb`；最小 source-bound 实验验证 `#663399` → `#9966CC` 只替换图片 `srgbClr/@val`、只改变 owning SlidePart，并保留图片关系、crop/mask/effect、其它 3-D 状态与其它 ZIP part，复杂/不明确颜色 owner 继续 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneCameraPreset` 绑定普通 shape 的严格两子节点 `p:sp/p:spPr/a:scene3d/a:camera/@prst`，要求裸 camera/lightRig pair，新增 additive `PresentationShape.shape_3d_scene_camera_preset`/`shape3dSceneCameraPreset` native leaf；`PpjSourceBoundShape3dSceneCameraPresetLeafEditsAndReprojects` 验证 `orthographicFront` → `perspectiveFront` 的单 camera token splice、SlidePart-only changed-part、Open XML 和二次投影，camera rotation/FOV/zoom、backdrop、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneCameraZoomThousandthPercent` 绑定同一严格两子节点 scene 的显式 `p:sp/p:spPr/a:scene3d/a:camera/@zoom`，新增 additive `PresentationShape.shape_3d_scene_camera_zoom_thousandth_percent`/`shape3dSceneCameraZoomThousandthPercent` native leaf；`PpjSourceBoundShape3dSceneCameraZoomLeafEditsAndReprojects` 验证 `100000` → `50000` 的单 camera zoom token splice、SlidePart-only changed-part、Open XML 和二次投影，省略/default zoom、FOV/rotation、light rig rotation、backdrop、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneCameraFov60000` 绑定同一严格两子节点 scene 的显式 `p:sp/p:spPr/a:scene3d/a:camera/@fov`，新增 additive `PresentationShape.shape_3d_scene_camera_fov_60000`/`shape3dSceneCameraFov60000` native leaf；`PpjSourceBoundShape3dSceneCameraFovLeafEditsAndReprojects` 验证 `2700000` → `5400000` 的单 camera FOV token splice、SlidePart-only changed-part、Open XML 和二次投影，省略/非法 FOV、rotation、light rig rotation、backdrop、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneCameraRotationLatitude60000` 绑定带完整 `a:camera/a:rot` 的严格 scene owner 的 `p:sp/p:spPr/a:scene3d/a:camera/a:rot/@lat`，新增 additive `PresentationShape.shape_3d_scene_camera_rotation_latitude_60000`/`shape3dSceneCameraRotationLatitude60000` native leaf；`PpjSourceBoundShape3dSceneCameraRotationLatitudeLeafEditsAndReprojects` 验证 `0` → `1800000` 的单 camera rotation latitude token splice、SlidePart-only changed-part、Open XML 和二次投影，longitude/revolution、camera zoom/FOV、light state、backdrop、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneCameraRotationLongitude60000` 绑定同一完整 `a:camera/a:rot` 严格 scene owner 的 `p:sp/p:spPr/a:scene3d/a:camera/a:rot/@lon`，新增 additive `PresentationShape.shape_3d_scene_camera_rotation_longitude_60000`/`shape3dSceneCameraRotationLongitude60000` native leaf；`PpjSourceBoundShape3dSceneCameraRotationLongitudeLeafEditsAndReprojects` 验证 `0` → `1800000` 的单 camera rotation longitude token splice、SlidePart-only changed-part、Open XML 和二次投影，latitude/revolution、camera zoom/FOV、light state、backdrop、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneCameraRotationRevolution60000` 绑定同一完整 `a:camera/a:rot` 严格 scene owner 的 `p:sp/p:spPr/a:scene3d/a:camera/a:rot/@rev`，新增 additive `PresentationShape.shape_3d_scene_camera_rotation_revolution_60000`/`shape3dSceneCameraRotationRevolution60000` native leaf；`PpjSourceBoundShape3dSceneCameraRotationRevolutionLeafEditsAndReprojects` 验证 `0` → `1800000` 的单 camera rotation revolution token splice、SlidePart-only changed-part、Open XML 和二次投影，latitude/longitude、camera zoom/FOV、light state、backdrop、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropAnchorXEmu` 绑定完整三子节点 scene owner 的 `p:sp/p:spPr/a:scene3d/a:backdrop/a:anchor/@x`，要求裸 camera/lightRig 与完整 literal `anchor`/`norm`/`up` 三维坐标，新增 additive `PresentationShape.shape_3d_scene_backdrop_anchor_x_emu`/`shape3dSceneBackdropAnchorXEmu` native leaf；`PpjSourceBoundShape3dSceneBackdropAnchorXLeafEditsAndReprojects` 验证 `100` → `200` 的单 anchor X token splice、SlidePart-only changed-part、Open XML 和二次投影，anchor Y/Z、normal/up、camera/light transform、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropAnchorYEmu` 绑定同一完整三子节点 scene owner 的 `p:sp/p:spPr/a:scene3d/a:backdrop/a:anchor/@y`，新增 additive `PresentationShape.shape_3d_scene_backdrop_anchor_y_emu`/`shape3dSceneBackdropAnchorYEmu` native leaf；`PpjSourceBoundShape3dSceneBackdropAnchorYLeafEditsAndReprojects` 验证 `200` → `400` 的单 anchor Y token splice、SlidePart-only changed-part、Open XML 和二次投影，anchor X/Z、normal/up、camera/light transform、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropAnchorZEmu` 绑定同一完整三子节点 scene owner 的 `p:sp/p:spPr/a:scene3d/a:backdrop/a:anchor/@z`，新增 additive `PresentationShape.shape_3d_scene_backdrop_anchor_z_emu`/`shape3dSceneBackdropAnchorZEmu` native leaf；`PpjSourceBoundShape3dSceneBackdropAnchorZLeafEditsAndReprojects` 验证 `300` → `600` 的单 anchor Z token splice、SlidePart-only changed-part、Open XML 和二次投影，anchor X/Y、normal/up、camera/light transform、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropNormalDxEmu` 绑定同一完整三子节点 scene owner 的 `p:sp/p:spPr/a:scene3d/a:backdrop/a:norm/@dx`，新增 additive `PresentationShape.shape_3d_scene_backdrop_normal_dx_emu`/`shape3dSceneBackdropNormalDxEmu` native leaf；`PpjSourceBoundShape3dSceneBackdropNormalDxLeafEditsAndReprojects` 验证 `0` → `500` 的单 normal X token splice、SlidePart-only changed-part、Open XML 和二次投影，anchor、normal Y/Z、up、camera/light transform、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropNormalDyEmu` 绑定同一完整三子节点 scene owner 的 `p:sp/p:spPr/a:scene3d/a:backdrop/a:norm/@dy`，新增 additive `PresentationShape.shape_3d_scene_backdrop_normal_dy_emu`/`shape3dSceneBackdropNormalDyEmu` native leaf；`PpjSourceBoundShape3dSceneBackdropNormalDyLeafEditsAndReprojects` 验证 `0` → `500` 的单 normal Y token splice、SlidePart-only changed-part、Open XML 和二次投影，anchor、normal X/Z、up、camera/light transform、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropNormalDzEmu` 绑定同一完整三子节点 scene owner 的 `p:sp/p:spPr/a:scene3d/a:backdrop/a:norm/@dz`，新增 additive `PresentationShape.shape_3d_scene_backdrop_normal_dz_emu`/`shape3dSceneBackdropNormalDzEmu` native leaf；`PpjSourceBoundShape3dSceneBackdropNormalDzLeafEditsAndReprojects` 验证 `1` → `2` 的单 normal Z token splice、SlidePart-only changed-part、Open XML 和二次投影，anchor、normal X/Y、up、camera/light transform、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropUpDxEmu` 绑定同一完整三子节点 scene owner 的 `p:sp/p:spPr/a:scene3d/a:backdrop/a:up/@dx`，新增 additive `PresentationShape.shape_3d_scene_backdrop_up_dx_emu`/`shape3dSceneBackdropUpDxEmu` native leaf；`PpjSourceBoundShape3dSceneBackdropUpDxLeafEditsAndReprojects` 验证 `0` → `500` 的单 up-vector X token splice、SlidePart-only changed-part、Open XML 和二次投影，anchor、normal、up Y/Z、camera/light transform、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropUpDyEmu` 绑定同一完整三子节点 scene owner 的 `p:sp/p:spPr/a:scene3d/a:backdrop/a:up/@dy`，新增 additive `PresentationShape.shape_3d_scene_backdrop_up_dy_emu`/`shape3dSceneBackdropUpDyEmu` native leaf；`PpjSourceBoundShape3dSceneBackdropUpDyLeafEditsAndReprojects` 验证 `1` → `500` 的单 up-vector Y token splice、SlidePart-only changed-part、Open XML 和二次投影，anchor、normal、up X/Z、camera/light transform、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropUpDzEmu` 绑定同一完整三子节点 scene owner 的 `p:sp/p:spPr/a:scene3d/a:backdrop/a:up/@dz`，新增 additive `PresentationShape.shape_3d_scene_backdrop_up_dz_emu`/`shape3dSceneBackdropUpDzEmu` native leaf；`PpjSourceBoundShape3dSceneBackdropUpDzLeafEditsAndReprojects` 验证 `0` → `500` 的单 up-vector Z token splice、SlidePart-only changed-part、Open XML 和二次投影，anchor、normal、up X/Y、camera/light transform、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneLightRigPreset` 绑定同一严格两子节点 scene 的 `p:sp/p:spPr/a:scene3d/a:lightRig/@rig`，新增 additive `PresentationShape.shape_3d_scene_light_rig_preset`/`shape3dSceneLightRigPreset` native leaf；`PpjSourceBoundShape3dSceneLightRigPresetLeafEditsAndReprojects` 验证 `threePt` → `soft` 的单 light-rig token splice、SlidePart-only changed-part、Open XML 和二次投影，camera/light direction/rotation、backdrop、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneLightRigDirection` 绑定同一严格两子节点 scene 的 `p:sp/p:spPr/a:scene3d/a:lightRig/@dir`，新增 additive `PresentationShape.shape_3d_scene_light_rig_direction`/`shape3dSceneLightRigDirection` native leaf；`PpjSourceBoundShape3dSceneLightRigDirectionLeafEditsAndReprojects` 验证 `t` → `br` 的单 light-rig direction token splice、SlidePart-only changed-part、Open XML 和二次投影，camera/light preset/rotation、backdrop、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneLightRigRotationLatitude60000` 绑定带完整 `a:lightRig/a:rot` 的严格 scene owner 的 `p:sp/p:spPr/a:scene3d/a:lightRig/a:rot/@lat`，新增 additive `PresentationShape.shape_3d_scene_light_rig_rotation_latitude_60000`/`shape3dSceneLightRigRotationLatitude60000` native leaf；`PpjSourceBoundShape3dSceneLightRigRotationLatitudeLeafEditsAndReprojects` 验证 `0` → `1800000` 的单 light-rig rotation latitude token splice、SlidePart-only changed-part、Open XML 和二次投影，longitude/revolution、camera/light state、backdrop、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneLightRigRotationLongitude60000` 绑定同一完整 `a:lightRig/a:rot` 严格 scene owner 的 `p:sp/p:spPr/a:scene3d/a:lightRig/a:rot/@lon`，新增 additive `PresentationShape.shape_3d_scene_light_rig_rotation_longitude_60000`/`shape3dSceneLightRigRotationLongitude60000` native leaf；`PpjSourceBoundShape3dSceneLightRigRotationLongitudeLeafEditsAndReprojects` 验证 `0` → `1800000` 的单 light-rig rotation longitude token splice、SlidePart-only changed-part、Open XML 和二次投影，latitude/revolution、camera/light state、backdrop、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneLightRigRotationRevolution60000` 绑定同一完整 `a:lightRig/a:rot` 严格 scene owner 的 `p:sp/p:spPr/a:scene3d/a:lightRig/a:rot/@rev`，新增 additive `PresentationShape.shape_3d_scene_light_rig_rotation_revolution_60000`/`shape3dSceneLightRigRotationRevolution60000` native leaf；`PpjSourceBoundShape3dSceneLightRigRotationRevolutionLeafEditsAndReprojects` 验证 `0` → `1800000` 的单 light-rig rotation revolution token splice、SlidePart-only changed-part、Open XML 和二次投影，latitude/longitude、camera/light state、backdrop、extension 与其它 scene state 保持 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneCameraPreset` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:camera/@prst`；新增 additive `PresentationImage.shape_3d_scene_camera_preset`/`shape3dSceneCameraPreset` native leaf，`PpjSourceBoundPictureShape3dSceneCameraPresetLeafEditsAndReprojects` 验证单 camera-preset token splice、图片关系/mask/effect、light rig 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；复杂 scene 仍 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneCameraZoomThousandthPercent` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:camera/@zoom`；新增 additive `PresentationImage.shape_3d_scene_camera_zoom_thousandth_percent`/`shape3dSceneCameraZoomThousandthPercent` native leaf，`PpjSourceBoundPictureShape3dSceneCameraZoomLeafEditsAndReprojects` 验证 `100000` → `50000` 的单 camera zoom token splice、图片关系/crop/mask/effect、camera preset/light rig 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；省略/default zoom 与复杂 scene 仍 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneCameraFov60000` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:camera/@fov`；新增 additive `PresentationImage.shape_3d_scene_camera_fov_60000`/`shape3dSceneCameraFov60000` native leaf，`PpjSourceBoundPictureShape3dSceneCameraFovLeafEditsAndReprojects` 验证 `2700000` → `5400000` 的单 camera FOV token splice、图片关系/crop/mask/effect、camera preset/zoom、light rig 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；省略/非法 FOV 与复杂 scene 仍 source-owned。
  本轮后续增量（2026-09-06）：`shape3dSceneCameraRotationLatitude60000` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:camera/a:rot/@lat`；新增 additive `PresentationImage.shape_3d_scene_camera_rotation_latitude_60000`/`shape3dSceneCameraRotationLatitude60000` native leaf，`PpjSourceBoundPictureShape3dSceneCameraRotationLatitudeLeafEditsAndReprojects` 验证 `0` → `1800000` 的单 camera rotation latitude token splice、完整 longitude/revolution、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 rotation 与复杂 scene 仍 source-owned。
  本轮后续增量（2026-09-06）：`shape3dSceneCameraRotationLongitude60000` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:camera/a:rot/@lon`；新增 additive `PresentationImage.shape_3d_scene_camera_rotation_longitude_60000`/`shape3dSceneCameraRotationLongitude60000` native leaf，`PpjSourceBoundPictureShape3dSceneCameraRotationLongitudeLeafEditsAndReprojects` 验证 `0` → `1800000` 的单 camera rotation longitude token splice、完整 latitude/revolution、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 rotation 与复杂 scene 仍 source-owned。
  本轮后续增量（2026-09-06）：`shape3dSceneCameraRotationRevolution60000` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:camera/a:rot/@rev`；新增 additive `PresentationImage.shape_3d_scene_camera_rotation_revolution_60000`/`shape3dSceneCameraRotationRevolution60000` native leaf，`PpjSourceBoundPictureShape3dSceneCameraRotationRevolutionLeafEditsAndReprojects` 验证 `6000000` → `7200000` 的单 camera rotation revolution token splice、完整 latitude/longitude、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 rotation 与复杂 scene 仍 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneLightRigPreset` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:lightRig/@rig`；新增 additive `PresentationImage.shape_3d_scene_light_rig_preset`/`shape3dSceneLightRigPreset` native leaf，`PpjSourceBoundPictureShape3dSceneLightRigPresetLeafEditsAndReprojects` 验证 `threePt` → `soft` 的单 light-rig preset token splice、camera/direction、图片关系/crop/mask/effect、其它 3-D state 与其它 ZIP part 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 scene 与复杂 graph 仍 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneLightRigDirection` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:lightRig/@dir`；新增 additive `PresentationImage.shape_3d_scene_light_rig_direction`/`shape3dSceneLightRigDirection` native leaf，`PpjSourceBoundPictureShape3dSceneLightRigDirectionLeafEditsAndReprojects` 验证 `t` → `br` 的单 light-rig direction token splice、camera/preset、图片关系/crop/mask/effect、其它 3-D state 与其它 ZIP part 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 scene 与复杂 graph 仍 source-owned。
	本轮后续增量（2026-09-06）：`shape3dSceneLightRigRotationLatitude60000` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:lightRig/a:rot/@lat`；新增 additive `PresentationImage.shape_3d_scene_light_rig_rotation_latitude_60000`/`shape3dSceneLightRigRotationLatitude60000` native leaf，`PpjSourceBoundPictureShape3dSceneLightRigRotationLatitudeLeafEditsAndReprojects` 验证 `0` → `1800000` 的单 light-rig rotation latitude token splice、完整 longitude/revolution、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 rotation 与复杂 scene 仍 source-owned。
	本轮后续增量（2026-09-06）：`shape3dSceneLightRigRotationLongitude60000` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:lightRig/a:rot/@lon`；新增 additive `PresentationImage.shape_3d_scene_light_rig_rotation_longitude_60000`/`shape3dSceneLightRigRotationLongitude60000` native leaf，`PpjSourceBoundPictureShape3dSceneLightRigRotationLongitudeLeafEditsAndReprojects` 验证 `0` → `1800000` 的单 light-rig rotation longitude token splice、完整 latitude/revolution、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 rotation 与复杂 scene 仍 source-owned。
	本轮后续增量（2026-09-06）：`shape3dSceneLightRigRotationRevolution60000` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:lightRig/a:rot/@rev`；新增 additive `PresentationImage.shape_3d_scene_light_rig_rotation_revolution_60000`/`shape3dSceneLightRigRotationRevolution60000` native leaf，`PpjSourceBoundPictureShape3dSceneLightRigRotationRevolutionLeafEditsAndReprojects` 验证 `6000000` → `7200000` 的单 light-rig rotation revolution token splice、完整 latitude/longitude、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 rotation 与复杂 scene 仍 source-owned。
	本轮后续增量（2026-09-06）：`shape3dSceneBackdropAnchorXEmu` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:backdrop/a:anchor/@x`；新增 additive `PresentationImage.shape_3d_scene_backdrop_anchor_x_emu`/`shape3dSceneBackdropAnchorXEmu` native leaf，`PpjSourceBoundPictureShape3dSceneBackdropAnchorXLeafEditsAndReprojects` 验证 `0` → `500` 的单 backdrop anchor X token splice、完整 anchor Y/Z、normal/up、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 backdrop 与复杂 scene 仍 source-owned。
	本轮后续增量（2026-09-06）：`shape3dSceneBackdropAnchorYEmu` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:backdrop/a:anchor/@y`；新增 additive `PresentationImage.shape_3d_scene_backdrop_anchor_y_emu`/`shape3dSceneBackdropAnchorYEmu` native leaf，`PpjSourceBoundPictureShape3dSceneBackdropAnchorYLeafEditsAndReprojects` 验证 `100` → `500` 的单 backdrop anchor Y token splice、完整 anchor X/Z、normal/up、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 backdrop 与复杂 scene 仍 source-owned。
	本轮后续增量（2026-09-06）：`shape3dSceneBackdropAnchorZEmu` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:backdrop/a:anchor/@z`；新增 additive `PresentationImage.shape_3d_scene_backdrop_anchor_z_emu`/`shape3dSceneBackdropAnchorZEmu` native leaf，`PpjSourceBoundPictureShape3dSceneBackdropAnchorZLeafEditsAndReprojects` 验证 `200` → `500` 的单 backdrop anchor Z token splice、完整 anchor X/Y、normal/up、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 backdrop 与复杂 scene 仍 source-owned。
	本轮后续增量（2026-09-06）：`shape3dSceneBackdropNormalDxEmu` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:backdrop/a:norm/@dx`；新增 additive `PresentationImage.shape_3d_scene_backdrop_normal_dx_emu`/`shape3dSceneBackdropNormalDxEmu` native leaf，`PpjSourceBoundPictureShape3dSceneBackdropNormalDxLeafEditsAndReprojects` 验证 `1` → `2` 的单 backdrop normal X token splice、完整 anchor、normal Y/Z、up、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 backdrop 与复杂 scene 仍 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropNormalDyEmu` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:backdrop/a:norm/@dy`；新增 additive `PresentationImage.shape_3d_scene_backdrop_normal_dy_emu`/`shape3dSceneBackdropNormalDyEmu` native leaf，`PpjSourceBoundPictureShape3dSceneBackdropNormalDyLeafEditsAndReprojects` 验证 `2` → `3` 的单 backdrop normal Y token splice、完整 anchor、normal X/Z、up、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 backdrop 与复杂 scene 仍 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropNormalDzEmu` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:backdrop/a:norm/@dz`；新增 additive `PresentationImage.shape_3d_scene_backdrop_normal_dz_emu`/`shape3dSceneBackdropNormalDzEmu` native leaf，`PpjSourceBoundPictureShape3dSceneBackdropNormalDzLeafEditsAndReprojects` 验证 `3` → `4` 的单 backdrop normal Z token splice、完整 anchor、normal X/Y、up、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 backdrop 与复杂 scene 仍 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropUpDxEmu` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:backdrop/a:up/@dx`；新增 additive `PresentationImage.shape_3d_scene_backdrop_up_dx_emu`/`shape3dSceneBackdropUpDxEmu` native leaf，`PpjSourceBoundPictureShape3dSceneBackdropUpDxLeafEditsAndReprojects` 验证 `4` → `5` 的单 backdrop up-vector X token splice、完整 anchor、normal、up Y/Z、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 backdrop 与复杂 scene 仍 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropUpDyEmu` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:backdrop/a:up/@dy`；新增 additive `PresentationImage.shape_3d_scene_backdrop_up_dy_emu`/`shape3dSceneBackdropUpDyEmu` native leaf，`PpjSourceBoundPictureShape3dSceneBackdropUpDyLeafEditsAndReprojects` 验证 `5` → `6` 的单 backdrop up-vector Y token splice、完整 anchor、normal、up X/Z、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 backdrop 与复杂 scene 仍 source-owned。
本轮后续增量（2026-09-06）：`shape3dSceneBackdropUpDzEmu` 现在也覆盖严格图片 owner 的 `p:pic/p:spPr/a:scene3d/a:backdrop/a:up/@dz`；新增 additive `PresentationImage.shape_3d_scene_backdrop_up_dz_emu`/`shape3dSceneBackdropUpDzEmu` native leaf，`PpjSourceBoundPictureShape3dSceneBackdropUpDzLeafEditsAndReprojects` 验证 `6` → `7` 的单 backdrop up-vector Z token splice、完整 anchor、normal、up X/Y、图片关系/crop/mask/effect、camera/light state 与其它 3-D state 保留、SlidePart-only changed-part、Open XML 和二次投影；partial/非法 backdrop 与复杂 scene 仍 source-owned。
本轮后续增量（2026-09-05）：`tableBandedRows` 绑定矩形表格的直接 `p:graphicFrame/a:graphicData/a:tbl/a:tblPr/@bandRow`，在已有 `table.style.bandedRows` 语义之外提供独立 native leaf；`PpjSourceBoundTableBandedRowsEditsLeafAndReprojects` 验证 `1` → `0` 的单属性 token splice、表格文本与其它 ZIP part 保留、SlidePart-only changed-part、Open XML 和二次投影。
本轮后续增量（2026-09-05）：`tableBandedColumns` 绑定矩形表格的直接 `p:graphicFrame/a:graphicData/a:tbl/a:tblPr/@bandCol`，在已有 `table.style.bandedColumns` 语义之外提供独立 native leaf；`PpjSourceBoundTableBandedColumnsEditsLeafAndReprojects` 验证 `1` → `0` 的单属性 token splice、表格文本与其它 ZIP part 保留、SlidePart-only changed-part、Open XML 和二次投影。
本轮后续增量（2026-09-05）：`tableHeaderRows` 绑定矩形表格的直接 `p:graphicFrame/a:graphicData/a:tbl/a:tblPr/@firstRow`，在已有 `table.style.headerRows` 语义之外提供独立 native leaf；`PpjSourceBoundTableHeaderRowsEditsLeafAndReprojects` 验证 `1` → `0` 的单属性 token splice、表格文本与其它 ZIP part 保留、SlidePart-only changed-part、Open XML 和二次投影。
本轮后续增量（2026-09-05）：`tableFirstColumnEmphasis` 绑定矩形表格的直接 `p:graphicFrame/a:graphicData/a:tbl/a:tblPr/@firstCol`，在已有 `table.style.firstColumnEmphasis` 语义之外提供独立 native leaf；`PpjSourceBoundTableFirstColumnEmphasisEditsLeafAndReprojects` 验证 `1` → `0` 的单属性 token splice、表格文本与其它 ZIP part 保留、SlidePart-only changed-part、Open XML 和二次投影。
本轮后续增量（2026-09-05）：`tableLastColumnEmphasis` 绑定矩形表格的直接 `p:graphicFrame/a:graphicData/a:tbl/a:tblPr/@lastCol`，在已有 `table.style.lastColumnEmphasis` 语义之外提供独立 native leaf；`PpjSourceBoundTableLastColumnEmphasisEditsLeafAndReprojects` 验证 `1` → `0` 的单属性 token splice、表格文本与其它 ZIP part 保留、SlidePart-only changed-part、Open XML 和二次投影。

本轮后续增量（2026-09-05）：`tableLastRow` 绑定矩形表格的直接 `p:graphicFrame/a:graphicData/a:tbl/a:tblPr/@lastRow`，新增 additive `PresentationTable.last_row`/`table.style.lastRow` 语义和独立 native leaf；`PpjSourceBoundTableLastRowEditsLeafAndReprojects` 验证 authored → 去嵌入投影 → 单 `@lastRow` token 修改 → SlidePart-only changed-part → Open XML → 二次投影，缺失属性不发 leaf。

本轮后续增量（2026-09-05）：`PpjSourceBoundLineOpacityLeafEditsShapeAndConnector` 覆盖 shape 的 direct RGB 与 connector 的 direct theme line alpha native leaf；`lineOpacityThousandthPercent` 只改已有 `a:alpha/@val`，验证 authored → 去嵌入投影 → source-bound 单 token 修改 → SlidePart-only changed-part → Open XML → 二次投影，缺失 alpha 不发 leaf。

本轮后续增量（2026-09-05）：`PpjSourceBoundLineArrowSizeLeavesEditShapeAndConnector` 覆盖 shape direct RGB 与 connector direct theme 的 head/tail arrow width/length native leaves；四个 leaf 只改已有 `a:headEnd/@w|@len` 或 `a:tailEnd/@w|@len`，验证 authored → 去嵌入投影 → source-bound 单属性修改 → SlidePart-only changed-part → Open XML → 二次投影，缺失尺寸不发 leaf。

本轮后续增量（2026-09-05）：K-01 的 compact smooth points 回读边界从 24 扩到 32；`PpjKimiSmoothLineWithTwentyFivePointsPreservesCompactProjection` 用 25 点 authored 高阶 Bézier 验证去嵌入 PPJ 后仍回读为 `points/viewBox/curve: smooth`。高阶控制点反解改用有界 Householder QR，保留原有量化容差、公式证明和超过边界/任意 cubic 的 typed-path fail-closed 行为。

本轮后续增量（2026-09-05）：K-01 的 compact smooth points 回读边界从 32 扩到 48；`PpjKimiSmoothLineWithThirtyThreePointsPreservesCompactProjection` 用 33 点 authored 高阶 Bézier 验证去嵌入 PPJ 后仍回读为 `points/viewBox/curve: smooth`，继续沿用有界 Householder QR 和原有 fail-closed 证明。

- `PpjGapProfilesCompileAndReproject` 新增 authored/source-bound 普通轴标题、普通轴和 chart/radar data-label `numberFormat`、plot/series data-label 显示标志与位置 grammar token 的正向回投影，以及错误 kind 的 fail-closed 断言；目标 changed-part 仍限定为 ChartPart。

- `PpjSourceBoundMasterAndLayoutBackgroundEditsAndReprojects` 覆盖去嵌入 PPJ 后 master/layout 的 `nativeRef.setBackground` capability、direct `p:bg` 增改、仅 `ppt/slideMasters/slideMaster1.xml` 与 `ppt/slideLayouts/slideLayout1.xml` changed-part、Open XML SDK 验证和二次投影颜色恢复；复杂母版继承、inherited/non-text placeholder 图与 theme cascade 仍未开放。

- `PpjSourceBoundMasterAndLayoutPlaceholderEditsAndReprojects` 覆盖直接 `a:xfrm`/固定文本拓扑 placeholder 的占位符级 `nativeRef.setFrame` 与 `replaceText` capability、master/layout owner-only changed-part、Open XML SDK 验证和二次投影的坐标/文字恢复；不暴露无 direct frame、非文本或不规则 placeholder 图。

- `PpjSourceBoundMasterAndLayoutImageBackgroundCropOpacityEditAndReprojects` 覆盖直接嵌入 master/layout 图片背景的 crop/opacity source-bound 回写；两个 owner 各自只改对应 `slideMaster`/`slideLayout` XML，不改 slide、关系或媒体，并通过 Open XML 校验和二次投影恢复。
- `PpjSourceBoundMasterAndLayoutImageBackgroundReplacementClosesRelationshipsAndReprojects` 覆盖 master/layout 各自替换直接嵌入图片背景；每个 owner 只改自己的 XML、`.rels` 和新媒体，旧图片关系/媒体被清理，Open XML 校验和二次投影恢复替换资产与 opacity。

- `PpjSourceBoundDirectSlidePlaceholderFrameEditOnlyChangesSlidePart` 覆盖 slide placeholder 的完整 owner-local `a:xfrm`，`setFrame` 可在固定 x/y/width/height 之外写回有限 rotation/flip，验证只改 `ppt/slides/slide1.xml`、Open XML SDK 和二次投影；`PpjInheritedSlidePlaceholderProjectsEffectiveFrameAndMaterializesOnSlideEdit` 覆盖 slide→layout→master 的唯一继承 frame，首次编辑只在 `ppt/slides/slide1.xml` 物化 direct `a:xfrm` 并同时写回有限 rotation/flip，再二次投影恢复；`PpjSourceBoundMasterAndLayoutPlaceholderEditsAndReprojects` 另外覆盖 master placeholder 的 rotation/flip 三态：保留显式 `flipH=false`，删除 `rotation`/`flipV`，只改 `ppt/slideMasters/slideMaster1.xml`，并在二次投影中保留属性存在性；复杂 inherited transform 和跨 owner presence 仍不开放。

- `PpjSourceBoundPresetGeometryAdjustmentLeafEditsAndReprojects` 覆盖 recognized `roundRect` 的完整 preset `a:avLst/a:gd fmla="val N"`：逐槽 native leaf 只 token-splice 对应公式，保持 `prst`、guide 顺序和其他 XML，只改 `ppt/slides/slide1.xml`，通过 Office 2021 Open XML 校验并二次投影恢复调整值；`PpjSourceBoundPartialPresetGeometryLiteralSiblingEditsAndReprojects` 补充 partial list：公共 geometry 不伪造不完整数组，只为 literal sibling 发 leaf，修改后保留其他 guide 并二次投影恢复；formula/extension/child-bearing/unknown preset guide 图的整体编辑继续 source-owned。

- `PpjImageOpacityGrammarTokenAuthorsAndReprojects` 现同时覆盖 `design.styles.image` + `image.styleRef` 的 fit/opacity 字段级 precedence：named image style 可按声明顺序胜过 inline/direct 值，authored 去嵌入投影恢复生效 paint；source-bound 只对生效值触发既有 `setImageFit`/`setOpacity`，改动只保留图片所属 SlidePart 和既有媒体关系闭包，并由二次投影恢复值。复杂 mask、主题/master cascade 和跨 owner 样式关系仍不开放。

- 本轮新增的窄回归也已通过：`PpjSourceFreeModernCommentsAuthorAndReproject` 覆盖 PPJ 确定性生成 Office person、根评论、直接回复、textRange 锚点和状态并去嵌入后二次投影；`PpjModernCommentsProjectAndReprojectTextAndStatus` 保持既有 source-bound modern comment 编辑；`PpjSourceBoundTranslucentSolidBackgroundEditsAndReprojects` 覆盖直接数值/grammar opacity 写回 `p:bg`、单一 SlidePart changed-part 和二次投影恢复；`SourceBoundImageBackgroundCropAndOpacityEditOnlySlideAndReprojects` 覆盖直接嵌入图片背景的 crop/opacity 写回、`a:blipFill` 校验、单一 SlidePart changed-part 和二次投影恢复；`SourceBoundImageBackgroundReplacementClosesRelationshipsAndReprojects` 覆盖单 owner 背景图片替换、SlidePart `.rels`/新媒体增加、旧媒体清理和二次投影恢复；`PpjSourceBoundMasterAndLayoutImageBackgroundReplacementClosesRelationshipsAndReprojects` 覆盖两个 design owner 的图片替换、各自 `.rels`/新媒体增加、旧媒体清理和二次投影恢复；`NativeChartDataLeavesCoverCategoricalComboPlots` 覆盖三条 bounded column/line/area combo plot 的 `c:val/c:numRef` cache 与独立 embedded worksheet 列绑定、两点双 footprint 编辑、Open XML 校验和二次投影恢复；`NativeChartDataLeavesCoverCategoryCacheAndInlineWorksheetText` 覆盖 opaque native chart 的 `c:cat/c:strRef` 与直接 inline/string worksheet cell 类别缓存、`chartDataCategory` 双 footprint 编辑和二次投影恢复；`PpjDatasetEncodingAuthorsAllKimiSeriesFamilies` 另外覆盖 `seriesDefaults.dataLabels.textStyle` 的递归嵌套继承/局部覆盖；`SourceBoundPictureSmartArtCanReplaceAnExistingNodeAsset` 覆盖 OfficeKit 自有 `picture` SmartArt 的节点 asset 读回、`setSmartArtImage` 能力、嵌套 drawing `.rels`/media 关系闭包、源保持编译、Open XML 校验和二次投影。

- `PpjComponentWeightedStackAuthorsAndReprojectsDeterministicFrames` 覆盖 authored component horizontal/vertical `layout.weights` 的 `[1,2,1]` slot 比例、gap 保留、去嵌入后二次投影普通 shape frame，以及零权重和数量不匹配的 fail-closed 校验；完整 solver、跨对象约束和 source-bound 自动重排仍不在范围内。
- `PpjChartTitlePlacementAuthorAndEditSourceChart` 覆盖 combo ChartPart 标题的 `centeredOverlay` authored 写入、去嵌入回投影、source-bound 改为 `aboveChart` 的单 ChartPart changed-part 以及二次投影恢复；无标题的 `none`、公式/复杂标题拓扑和非 native ChartPart 家族继续保持有界或 fail-closed。
- NativeAOT 测试项目成功编译；仅保留现有 4 条 nullable warning，无编译错误。
- Kimi/PPJ 与完整 PPT 有界能力簇共 49 个 focused xUnit 用例全部通过，覆盖：曲线 points/smooth（含任意多段 Bézier 保持 typed path）、chart frame image fill、dataset/encoding 13 类族（含 scatter/bubble canonical numeric rows）、formula chart reference、native chart data（bar/line/area/pie/doughnut/radar category/value、bounded combo value、scatter/bubble X/Y/size，以及保留 `externalData` 的 `c:strLit`/`c:numLit` ChartPart-only literal cache）、grammar token（含表格布尔样式标志、source-bound 图表文字样式和 common plot 标量）、图表轴位置、图表标题位置、图表 plot-area line、图表 style index、稀疏点 data-label 的 fill/line/text、表格 cell style/image relationship copy-on-write、固定拓扑 mixed-run table text body、表格 field 缓存文本、固定拓扑 table `run.break`、普通文本框/带文本形状/占位符 direct bodyPr style（含 canonical `normalAutoFit` 百分比 native leaves）、PPJ 段落 tab stop authored/source-bound 回投影、master/layout、多 run field、页面和 group reading order/action 及普通 group 外层 frame 与 `childFrame/chOff` leaves、component repeat/image policy（含 weighted stack）、literal custom geometry、recognized preset geometry adjustment leaf、有限 source-bound timing graph、OfficeKit 自有 picture SmartArt 节点 asset 替换；动画用例同时证明 `timing.nodes[]` 和 compact `animations[]` 的 trigger 规范化及错误 trigger fail-closed。
- `git diff --check`、PPJ review smoke、Claude marketplace smoke 和 PPJ schema 解析通过。
- `node test/reference-skills.mjs` 通过（`reference skill plugins smoke ok`）；它只验证 Skill/引用插件 smoke，不等同于 Windows PowerPoint 宿主验收。
- 本轮没有执行完整 `npm test`、Windows PowerPoint、完整发布或 npm 发布；这些均不是当前 PPJ gap 分母。

- Windows PowerPoint 的“未验收”只保留在 F-18/宿主证据栏；它不改变 K/F 任一语义条目的状态，不计入完成度分子或分母，也不作为本轮待实现项。

**本轮后续增量（2026-09-07）：** Kimi `ChartConfig.plotAreaLine` 已接入 PPJ `style.plotAreaLine`/schema/wire/native ChartML。普通/combo ChartPart 以 `c:plotArea/c:spPr/a:ln` 承载有界的直接 RGB、宽度和预置虚线，并与已有 `plotAreaFill` 共用 shape-properties 容器；authored 与 source-bound 回投影恢复该轮廓，source-bound 只改目标 ChartPart。主题/效果线图、复杂 shape-properties、vector fallback 和不支持的图表族继续 source-owned/fail closed。`PpjChartPlotAreaLineAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮后续增量（2026-09-07）：** Kimi `ChartConfig.styleIndex` 已接入 PPJ `styleIndex`/schema/wire/native ChartML。普通/combo ChartPart 以直接 `c:style/@val` 承载 OfficeKit 有界的 1–48 内建样式索引；authored 与 source-bound 回投影恢复该值，source-bound 只改目标 ChartPart。vector-lowered、advanced chart、非法样式拓扑和其他不支持内容继续 source-owned/fail closed。`PpjChartStyleIndexAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮后续增量（2026-09-07）：** Kimi `ChartConfig.dataTable` 已接入 PPJ `dataTable`/schema/wire/native ChartML。普通/combo ChartPart 以直接 `c:plotArea/c:dTable` 承载四个保留 presence 的边框/图例键开关、直接 none/solid/gradient 填充和直接 RGB/预置虚线；authored 与 source-bound 回投影恢复该状态，source-bound 只改目标 ChartPart。文本属性/扩展子节点、主题/效果 paint、非法拓扑和不支持的图表族继续 source-owned/fail closed。`PpjChartDataTableAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮后续增量（2026-09-07）：** F-07 的原生 radar 变体已接入 PPJ `style.radarStyle`/schema/wire/native ChartML。普通 radar ChartPart 以直接 `c:radarChart/c:radarStyle` 承载 `standard`、`marker`、`filled` 三种有界模式；authored 与 source-bound 回投影恢复该值，source-bound 只改目标 ChartPart。3-D、扩展/效果化 ChartML、复杂组合、vector fallback 和不规则 radar 拓扑继续 source-owned/fail closed。`PpjRadarStyleAuthorProjectAndEditOnlyChartPart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮后续增量（2026-09-07）：** Kimi `Chart.fontFamily` 已闭合为 PPJ 顶层 `chart.fontFamily`。普通/combo 原生 ChartPart 读写 `c:chartSpace/c:txPr` 下的直接 Latin typeface；去嵌入投影恢复该字段，source-bound 修改只改目标 ChartPart 并可二次投影。East Asian/complex-script 全局 fallback、额外全局文字属性、vector-lowered 图表和不支持的图表族保持 source-owned/fail closed。`PpjChartFontFamilyAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 修改 → 二次投影。

**本轮后续增量（2026-09-07）：** Kimi `ChartSeriesConfig.explosion` 已闭合为 PPJ `data.series[].explosion`。普通 pie/doughnut ChartPart 读写直接 `c:ser/c:explosion`，保留 0–400 范围内的显式零；去嵌入投影恢复 series 字段，`setChartSeriesStyle` source-bound 修改只改目标 ChartPart，并通过二次投影确认结果。point-level `dPt` explosion 仍归 `pointStyles`；combo、3D、扩展和不规则 ChartML 继续 source-owned/fail closed。`PpjChartSeriesExplosionAuthorAndEditSourceChart` 覆盖 authored → 去嵌入 → source-bound 改为 0 → 二次投影。

## 8. 当前结论

PPJ 目前已经不是“只有少量 shape 的 PPT JSON”。它在 Kimi 的常规图片、背景、文本、表格和图表范围内有较宽的 typed surface，并且在动画、SmartArt、source-bound 保真和部分高级图表上已经超过 Kimi `pptd.md` 的公开描述。本轮已经把 K-01 的 bounded literal/points line profile（含 source-bound changed-part/二次投影证据）以及 K-02/K-03 的主要 authored 语言缺口接到 schema、wire、NativeAOT 编译/投影和 focused round-trip；K-02 的 ChartPart image frame 现在覆盖直接关系、crop、stretch/tile、opacity、单 owner image 替换时的 `.rels`/media 清理与二次投影。K-03 现在还把安全 opaque native chart 数据点收敛为 `nativeRef.leaves[].kind=chartDataValue/chartDataXValue/chartDataYValue/chartDataBubbleSize`，可证明地按通道同步 ChartPart cache 与 embedded worksheet cell，且 category profile 覆盖 bar/line/area/pie/doughnut/radar、bounded column/line/area combo、scatter/bubble 覆盖 X/Y/size。新增的 literal custom geometry source-bound path edit、静态 typed field display edit，以及 F-06 direct cell fill/border/单一直接嵌入 cell image replacement/固定拓扑多 run 文本替换/跨段落样式一致多 run text-style/mixed-run text-body source-bound edit 也已有最小 changed-part/二次投影证据；本轮又闭合 OfficeKit 自有 `picture` SmartArt 的节点图片 asset 投影和已有节点替换，关系源路径覆盖缓存 drawing 的嵌套 `.rels`/media 闭包，仍不宣称任意第三方 picture SmartArt 可编辑。K-02 的 gradient/image frame、K-04 的 normal opacity、`compositing.opacity`、image/line opacity token profiles、K-05 的 grammar evaluator、color tint/shade 以及文字/形状/chart 标题基本属性 token authored lowering、K-06 的 imagePolicy 槽位 profile、schema-v3 replacement plan 和显式 PPJ apply transaction、K-07 的 repeat/autoReverse/easing/trigger profile、K-08 的 authored grid/flow/anchor repeat 和旋转/阴影 visual-bounds review 也已接入。F-13 的有限 shape click/hover action、F-14 的 explicit readingOrder/machine accessibility review、F-15 的有限 color transform 也有窄验收证据，但都还不是完整 PowerPoint parity。

本轮 K-05 又补上 source-bound 图表标题/图例/数据标签/坐标轴文字样式的 kind-checked token 写回；它仍只覆盖明确的 ChartPart 文字样式 owner，不等于完整 theme/master cascade。K-08/F-16 的只读 review 又补上有限 line/connector 箭头头型 visual bounds 与稳定人工建议；mask 轮廓、真实字体测量、solver 和显式 apply 仍未完成。

但仍不能说已经追平 Kimi 的全部原语，更不能说已经接近完整 PowerPoint：

- 对 Kimi 的直接差距现在集中为“已实现有限 profile、仍缺完整闭环”：通用 chart dataset/encode 的更广泛跨 family 组合和 workbook footprint、可写回的 design grammar，以及 line 的 authored 曲率 sugar 无损保留与完整 geometry；K-01 literal/points line、K-02 solid/gradient/image chart frame（单 owner image replacement 的 relationship/media closure 已闭合）、K-03 高层本地公式窄路径和 opaque native chart category/value 双 footprint 窄路径（category `c:strRef` 直接 inline/string cell、bar/line/area/pie/doughnut/radar 以及 bounded column/line/area combo value、scatter/bubble X/Y/size）已有分别闭合的残差证据，但还不是通用公式/工作簿同步。
- 对两者共同的产品缺口是 blend/isolation、真实文本测量、通用自动排版和遮挡求解；normal opacity、authored component grid、显式 reading order 和机器 accessibility review 已有边界实现，但 review 不提供自动修复，也不替代宿主 Checker；
- 对完整 PowerPoint 的主要差距是任意 OOXML/关系拓扑、完整 ChartML/DiagramML/timing、媒体/OLE/3D/宏、主题效果继承和 source-bound 交互关系编辑；真实宿主行为只在 F-18 单独记录，不计入 PPJ gap；
- 当前正确的路线仍然是“有界语义 + 可证明编辑 + 不支持时保留/失败”，而不是把所有 XML 标签投影成看似可编辑的普通 PPJ 字段。

当前工作树是本地未提交状态；已完成的描述表示代码和 focused evidence 已存在，不表示已经合入远端或完成发布。Windows PowerPoint lane 按本轮范围保持“未启动”，且不计入 PPJ gap 或完成度分母。
