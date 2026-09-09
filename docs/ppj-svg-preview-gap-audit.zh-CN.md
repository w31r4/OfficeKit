# OfficeKit 本地 PPT 预览渲染器差距审计

审计日期：2026-09-09。状态：**预览链路已打通，功能覆盖和可靠性尚未完成**。

后续实施进度（2026-09-09）：下文保留原始审计快照；本表及对应章节的“修复进度”优先于原始缺陷描述。

| 差距 | 已实施并验证 | 尚未完成 |
| --- | --- | --- |
| G-12 输出安全与证据 | 新目录独占发布、文件 hash/清单、pending/final 生命周期、失败保留、源/候选/资产身份；轻量故障测试、真实 authored/source-bound codec+sharp 测试及 gate-policy 通过 | 本轮执行专项检查，未运行完整 fast/slow gate；视觉正确性仍需独立验证 |
| G-05、G-11 的相关部分 | 资产复用已加载字节；空资产诊断；发布失败不会自报完整成功 | 图片视觉语义、字段级支持等级及其他诊断缺口仍未完成 |
| 其余 G 编号 | 本轮未改变绘制语义 | 按原审计要求继续逐项补齐 |

G-12 使用方法与清单字段见 [预览输出说明](ppj-preview-output.md)，实施清单见 [ppj-preview-output-evidence](../openspec/changes/ppj-preview-output-evidence/tasks.md)。以上不是全部 gap 的完成声明。

本文面向 OfficeKit 维护者和后续接手开发的 Agent，记录本地 SVG/PNG 预览距离“简单、覆盖现有功能、能可靠辅助结构与视觉 review”的实际差距。文档记录问题和验收要求，不代表这些问题已经修复，也不代替实施规格。

## 1. 结论、范围和证据口径

### 1.1 当前结论

`officekit ppj preview` 已能读取 PPJ，调用现有编译器，然后生成 SVG、PNG 和 `render.json`。这证明了一条可运行的本地预览链路，但不能证明预览忠实表达了输入。

目前存在会改变内容含义的错误：带文字的形状丢失几何本体、饼图不表达数值比例、连接线不依据端点关系、复杂图表读取错误的数据字段。这些问题不能用“只是视觉近似”解释。部分错误还没有诊断，甚至被标为 `supported`。

因此，当前不应宣称：

- 已全面覆盖 OfficeKit 现有演示文稿功能；
- 已具备可靠的独立视觉验收能力；
- 全量 smoke 通过就代表图表、图片或 source-bound 内容显示正确；
- 能输出预览就等于 PPTX 导出、编辑保真或宿主行为通过验收。

本次不提供完成百分比：没有经过字段级核对的分母，也没有足够的语义与视觉断言来定义分子。

### 1.2 要实现的目标

目标是本地个人使用的 Skill 和 CLI 运行时，不引入前后端服务架构。预览应复用 OfficeKit 的语义和布局能力，以较轻的绘制后端输出静态视觉检查结果。

“覆盖现有功能”应落实为两层要求：

1. **覆盖核对完整**：每个现有元素、视觉字段和 source-bound 边界都有明确归属、实现状态、诊断和回归案例，不能静默丢失。
2. **实际绘制完整**：OfficeKit 已支持且属于静态页面显示范围的功能，应按其真实语义绘制。把所有未实现项改成占位，只能完成诚实报告，不能算渲染器目标完成。

对于未知原生内容，允许保留源预览或有说明的占位；不能据此宣称它可编辑。动画、音视频播放、宏和 PowerPoint 宿主交互不由一张静态 PNG 验证，但其存在、静态呈现和检查范围必须明确。

不要求与 PowerPoint 像素级一致。允许字体栅格化等合理差异，但不允许改变数据比例、方向、层级、缺失值含义或对象可见性。

### 1.3 本次快照

| 项目 | 本次记录 |
| --- | --- |
| HEAD | `2f338d23ae35eba88cddc8f2e78e84ce8dcf190b` |
| Node.js | `v23.10.0` |
| `src/ppj/svg-preview.mjs` SHA-256 | `2f1d1c2c072efa1f1e2255fa1e912b361b21c735dff0e863f072f3444a2b6421` |
| canonical fixture SHA-256 | `6b3e6c50a67da620494df75aac7a10bb343b0db88e6d97c9f803437b8e5ae048` |
| 工作区情况 | 存在对数轴 `logBase` 的未提交实现和 OpenSpec 文档；本审计不修改这些在途工作 |
| 本次实际执行 | 两项预览专项测试、全量 PPJ preview smoke、canonical fixture 输出检查、schema/声明/矩阵比对 |
| 未执行 | 完整 `npm test`、NativeAOT 重建、对数轴验收、全量图片人工盲评、宿主 PowerPoint 验收、性能基准 |

注意：测试使用当前运行时能加载的 codec，未重建 NativeAOT，不能将本次预览测试视为工作区 C# 改动已经生效的证据。本文也不声称已固定所有字体、原生二进制和资源环境。

证据分为三类：**运行确认**表示本次实际执行所得；**代码确认**表示当前实现可以直接确定的行为；**待验证**表示尚缺专门案例，不能推断成功或失败。下面的完成条件都是后续要求，而非本次完成声明。

本文术语：`authored` 指从结构化输入创建文稿；`source-bound` 指编辑仍绑定原始 PPTX、原生对象和所有权证据；`opaque` 指无法完整、安全建模的原生内容；`canonical JSON` 指规范化程序文本，不承诺已展开布局；`fixture` 指可重复运行的固定测试输入；`smoke` 仅检查基本路径能否跑通。

写作期间有其他工作流继续更新对数轴、registry 和能力矩阵。本文保留测试发生时的观测，并在第 6、8 节注明复核变化；这些并行更新不算本审计完成的修复。

## 2. 当前执行链路及能力边界

### 2.1 三个入口不能混为一谈

| 入口 | 当前路径 | 能提供的证据 |
| --- | --- | --- |
| `officekit ppj preview` | PPJ → workspace → C# 编译 → 返回的 `programJson` → JS 拼接 SVG → 可选 sharp 栅格化 | OfficeKit 自有预览；目前仅部分表达页面 |
| `officekit ppj render` | PPJ → PPTX → LibreOffice → PDF → Poppler → PNG | 导出文件经外部 Office 软件处理后的静态显示证据；仍需视觉审查 |
| `officekit ppj review` | PPJ → PPTX → `reviewArtifact` | 结构等检查；当前调用显式传入 `visualReview: unavailable`，不等于已检查预览图片 |

入口见 [CLI](../src/ppj/cli.mjs)、[自有预览](../src/ppj/svg-preview.mjs)、[外部渲染和 review](../src/ppj/render-review.mjs)。当前 Presentations Skill 的交付示例仍使用 `ppj render`，新增 `preview` 不会自动替换它。

### 2.2 已经具备的基础

- 复用 `loadPpjWorkspace` 和 `compilePpjWorkspace`，不是另行解析 PPTX XML 的第二套文件 codec。
- CLI 对自有预览实现使用动态导入；`sharp` 也在需要写输出时才加载。
- SVG 保留 `data-officekit-id`，方便对应 PPJ 元素。
- 按页面元素数组顺序输出，能够保持这一层的绘制顺序。
- 已有部分缺失资源、占位、未知元素及顶层越界诊断。
- `render.json` 记录编译输入程序和编译输出的 hash；有 basic SVG/PNG smoke。

这些基础可以保留，但不能扩展解释为字段保真、编辑保真或最终视觉结果正确。

### 2.3 G-01：预览没有拿到完整的已解析绘制结果

`renderPpjToSvg` 读取 `compiled.programJson`，而 authored 编译器返回的是 `validation.CanonicalJson`。编译器另外持有 expansion/build plan，返回的 JSON 并不等于已经展开、排好布局、解析完样式的绘制场景。

本次运行 canonical fixture 后，返回的 JSON 中仍有 **1 个 `component`**。预览端随后自行读取 label/value、安排位置；它没有复用该组件实际编译出来的元素树。

影响包括组件 repeat、dataset 编码、styleRef、主题和 grammar token 等高层表达：PPTX 编译可能处理正确，预览却绕过这些处理重新猜测显示。

来源：[workspace 编译转发](../src/ppj/workspace.mjs)、[authored 编译回执](../native/OfficeKit/src/OfficeKit.Codec/PpjAuthoredPresentationCompiler.cs)、[预览入口](../src/ppj/svg-preview.mjs)。

完成条件：明确一个共用的解析/布局输出边界，预览消费编译器同义的结果；组件和 dataset 的高层写法与等价展开写法应得到等价预览。具体采用共用中间表示还是暴露现有布局结果，需要另行设计，本文不指定新增大框架。

## 3. 元素与页面视觉差距

下表对齐 [PPJ schema](../src/ppj/ppj-v1.schema.json) 中的 16 类元素。声明状态来自 [preview capabilities](../src/ppj/svg-preview-capabilities.json)，不是本审计认可的完成状态。

| 元素 | 声明状态 | 实际行为和未覆盖范围 |
| --- | --- | --- |
| `text` | supported | 输出纯文本和固定行距；未完整处理 runs、字体、字号、段落、换行、AutoFit、列表及样式继承 |
| `shape` | supported | 无文字时基本画矩形；有文字时提前进入文字分支，几何和填充丢失 |
| `line` | 未列入声明 | 无独立绘制分支，通常进入未知元素占位；路径和自由曲线未实现 |
| `icon` | 未列入声明 | 无图标轮廓绘制分支 |
| `image` | supported | 使用固定 contain 式 `<image>`；裁切、焦点、mask、边框和 effects 未完整映射 |
| `chart` | 按 chartType 声明 | 各类型差距见第 4 节，不能整体认定支持 |
| `table` | supported | 均分行列；未按显式尺寸、合并单元格和样式绘制 |
| `connector` | supported | 固定水平线，不解析 from/to、anchor、路线和箭头 |
| `group` | supported | 递归输出子元素，没有 childFrame 与父 frame 的坐标转换 |
| `media` | 未列入声明 | 无专门的静态海报/封面显示或播放能力说明 |
| `placeholder` | 同时列入 supported、partial | 实现为虚线框加文字，并报告 partial；声明自相矛盾 |
| `smartArt` | 未列入声明 | 没有消费 SmartArt 实际布局，落入通用兜底 |
| `ole` | 未按实际类型列入声明 | 声明用了 `embeddedOle`；没有使用 OLE `previewAsset` 的专门分支 |
| `opaque` | opaque | 有 `previewAsset` 时可以显示图片并报告 partial，否则通用虚线框；缺完整源预览和诊断回归 |
| `component` | partial | 仅抽取 label/value 并竖排，不按组件定义、variant 和真实布局展开 |
| `slot` | 未列入声明 | 没有独立解析/绘制分支，须与模板或组件展开共同处理 |

`nativeRef` 是源绑定字段，不是与 `shape` 平级的实际元素类型。将它列入类型声明不能证明 source-bound 各 owner 已覆盖。即使某个类型有通用占位，也不代表该类型的所有字段和诊断已覆盖。

### 3.1 G-02：文字和形状内容缺失

代码中的 `if (e.type === "text" || e.text)` 位于 shape 分支之前。任何带真值 `text` 的形状都会只输出文字。

本次 canonical fixture 的决策节点 `decision-flow-gate` 本来是 `flowChartDecision`，有填充、白色文字和居中设置；实际 SVG 为：

```xml
<g data-officekit-id="decision-flow-gate"><text x="806" y="302" font-family="Arial, sans-serif" font-size="18" fill="#172033"><tspan x="806" dy="0">Pass?</tspan></text></g>
```

该组没有菱形或背景。此节点也没有出现在 diagnostics 中。

其他代码确认的差距：

- `textValue` 把每个 run 用换行连接；同一段内“普通字 + 加粗字”会被错误拆行。
- 字号读取 `textStyle.fontSize`，没有完整消费实际 `defaultText.size`、run style 和命名样式。
- 文字首行偏移固定为 18、后续行偏移固定为 22；没有完整排版、边距、对齐和溢出处理。
- 所有普通 shape 都输出 `<rect>`，忽略 `geometry.preset` 和 preset adjustments。
- 自定义路径诊断检测 `geometry.customPaths/path`，而 schema 的自定义几何是 `kind: custom`、`viewBox`、`paths`；正常的 schema 字段可能被静默漏报。

完成条件：文字与形状作为同一元素的两个显示部分保留；按真实几何绘制；富文本 run 不改变段落边界；未支持的文本效果必须定位到具体字段并报告。

### 3.2 G-03：变换、可见性与组件布局

`frame.rotation/flipH/flipV` 没有用于 SVG 变换，group 也没有处理 `childFrame` 映射。元素 `hidden` 没有参与绘制过滤。页面隐藏状态没有形成明确的选择策略和输出证据。

组件分支自行平均分配竖向槽位，没有处理组件定义中的实际元素、variant、horizontal/grid/flow、权重与锚点。该分支用 `||` 取值，还可能把数值 0 当作空值省略。

完成条件：嵌套组变换、镜像、旋转和可见性与编译语义一致；组件使用实际展开结果；隐藏页是否包含在预览中必须是明确策略，而非偶然输出。

### 3.3 G-04：样式、主题和页面背景

预览的颜色函数主要接受字符串或少数 hex/value 字段，没有完整解析 PPJ 颜色 token、样式引用、继承及透明度。`page.background.color` 若是 token 对象，会回落成白色；shape 常回落到固定蓝灰色。

当前没有完整应用 styleRef、designGrammar 优先级、主题/master/layout、gradient/image fill、stroke dash/opacity、shadow/glow/reflection/softEdge、compositing 等显示状态。

完成条件：复用样式解析结果；实际支持的填充、轮廓、透明度与 effects 有对应绘制或明确边界。不能一律用默认色替换，再报告 supported。

### 3.4 G-05：图片和资源一致性

图片固定 `preserveAspectRatio="xMidYMid meet"`，没有按 `cover/contain/stretch/tile/none` 区分处理。crop、focus、mask、border、shadow 等字段也没有完整绘制。

这会使图片主体范围、边缘和阴影与导出结果不同。原始 PNG 的透明像素可能由栅格后端正常显示，但这不等于整套图片语义已经支持，也没有证据证明任意输入的主体边界可靠。

资源加载另有两个问题：

1. workspace 已读取资产字节，预览却又按 URI 读一次磁盘，未复用同一份已加载数据；编译和绘图之间存在资源变化窗口。
2. 第二次读取失败会被吞掉并生成空 `href`。Map 中仍存在 asset 对象，因此普通 image 的“资产不存在”检查不一定能发现这个错误。

这不表示所有缺失资源都能绕过 workspace；很多错误会更早失败。问题在于预览自己的二次读取和错误状态不完整。

完成条件：消费同一份验证过的资产；裁切与 mask 使用已知参数，不猜主体；无法确定的信息保守显示并说明；空字节/解码失败不能伪装为成功。

### 3.5 G-06：表格与连接关系

表格直接用元素宽高除以行列数量。canonical fixture 的列宽为 116/158，但预览画成两列各 137；行高和单元格合并、填充、边框、文字排版也没有完整应用。

connector 直接画 frame 中线，没有读取 `from/to.element`、anchor、connectorType 和箭头。本次 canonical fixture 连接线输出只有水平 `<line>`，没有 `endArrow: triangle` 对应的箭头。

完成条件：表格几何来自真实尺寸与 span；连接线随实际端点和锚点变化，不能用固定方向代替关系；独立 `line` 和 `connector` 应分清职责。

### 3.6 G-07：opaque、源绑定和非静态能力

有 `opaque.previewAsset` 的源图像显示路径已经存在，不应误记为“完全没有 opaque 支持”。但仍缺：

- 无预览时，把 summary、nativeKind、visibleText 和未支持原因清楚呈现给 reviewer；目前通常只有无说明的虚线框。
- 对 OLE 等其他实际类型的源预览处理。
- 对 imported/nativeRef 修改后，预览显示的内容是否确实对应修改后状态的验证。
- 第三方输入的原包保留、修改范围、重新导入与稳定 ID 证据。
- 对动画、媒体、交互的静态检查范围说明，避免一张图被误当作播放/行为验收。

预览使用一张源快照，不等于把可编辑对象写成图片；两者应分开描述。不能为生成预览而破坏原包或把 opaque 拓扑重写成猜测结构。

## 4. 图表差距

### 4.1 G-08：数据归一化和公共坐标系统缺失

预览直接读 `data.categories` 和 `data.series[].values`，未完整处理 `dataset/encoding/dataFilter/seriesDefaults`。坐标范围基本统一使用 `Math.max(1, ...values)`，没有实现负数范围、明确 min/max、双轴、数值 X 轴、轴反向及其他缩放方式。

以下字段也没有完整应用：title、legend、轴标题与刻度、网格、numberFormat、dataLabels、marker、pointStyles、trendlines、errorBars、chart frame 和 plotArea 样式。

这些字段既可能影响外观，也可能影响事实理解。例如把不同单位的主副轴压到同一尺度，会改变读者对两条序列变化的判断。

完成条件：绘图消费归一化数据；每个 mark 可追溯到 series/point 和正确坐标轴；数据表、标签、图例与图形一致；纯样式差异和数据关系错误分开报告。

### 4.2 G-09：逐类型差距

下表覆盖当前 schema 的 16 种 `chartType`，另列两种通过字段表达的变体。具体实现见 [chart 分支](../src/ppj/svg-preview.mjs)。

本表为代码与 schema 对照结果，不表示本次对每一种类型都执行了正负案例；最后一列的验证案例是待补测试。实际运行范围见第 6.1 节。

| 类型/变体 | 当前绘制 | 关键差距与最小验证案例 |
| --- | --- | --- |
| bar | 与 column 共用竖矩形逻辑 | 没有横向条形布局；仅设置元素 chartType 而省略 series.chartType 时可能没有数据 mark；验证普通横向两系列 |
| column | 固定宽度竖矩形 | 各系列重叠在相同位置；未正确处理分组/堆叠/负值，null 被转 0；验证两系列正负值与缺失点 |
| line | 按有限数值拆 polyline 段 | series 类型继承不完整；单点段被丢弃、marker 未画；验证 `[1, null, 3]` 保留两个独立观测 |
| area | 与 line 相同的无填充 polyline | 缺面积、基线和正确堆叠；验证多系列面积含缺失点 |
| combo | 拼接 column/line/area 分支 | 共用最大值，忽略主副轴、单位与部分样式；验证柱线双轴不同量纲 |
| scatter | 按数组序号或假定的对象 x/y 画点 | schema 的 values 是 number/null，真实 X 在 `xValues`；当前忽略该通道；null 还可能在访问 `v.y` 时抛错，需专门回归确认 |
| bubble | 无专门气泡映射 | 没有同时表达 X/Y/size；验证非均匀 X 和不同 bubbleSizes |
| pie | 一个固定半径的单色圆加类别文字 | 没有按比例分扇区；`[1,9]` 与 `[9,1]` 不会改变扇区；这是事实表达错误 |
| doughnut | 无专门环形映射 | 缺比例扇区、holeSize 和切片布局；验证两种明显不同的比例 |
| radar | 无专门雷达映射 | 缺 spoke/value axis 和闭合多边形；验证多系列、缺失点和标签 |
| heatmap | 按当前数值 min/max 画颜色格 | 缺声明色域、色板和样式语义；缺失点已有虚线格，但没有整体正确性断言 |
| treemap | 把 values 画成一行宽度比例条块 | 忽略 `series.parents/levels`，没有表达层级；验证至少两层父子关系 |
| sunburst | 按 values 画单层扇区 | 忽略 `parents/levels`，没有同心层级；验证多层而非单层 pie |
| waterfall | 把每个值持续累加 | 忽略 `pointRoles: delta/total`，尺度也未按累计范围计算；验证中间 total 和负增量 |
| candlestick | 假定 values 每项含 open/high/low/close | 实际 schema 使用 `openValues/highValues/lowValues` 与 values 通道；当前可产生只有背景的图，需真实 OHLC 案例 |
| sankey | 读取 `data.nodes/links`，画节点框和流量条数 | 这不是当前标准数据字段；实际使用 `series.sources/targets/values`；代码也没有画 links；验证多源汇流与分流 |
| pictograph | 根据最大值归一化成最多约 12 个圆 | 忽略 `symbol.unit/iconName/preset`，四舍五入改变数量表达；它是 symbol 变体，不是合法独立 chartType |
| stream stacking | 用错开的粗 polyline 模拟带状图 | 不是实际堆叠/居中面积，缺失点会经 num 变 0；`streamgraph` 本身也不是 schema 的独立 chartType |

这里没有把 treemap/Sankey 的输入误写成通用 nodes/links：应以当前 schema 和编译器真实字段为准，而不是沿用渲染器自己假定的数据模型。

### 4.3 G-10：缺失值规则尚未统一

普通 line 分支用 `Number.isFinite` 拆段，是已有的有限保护；但单点段被过滤，其他图表还使用 `Number(value)` 或 `num`，而 `Number(null) === 0`。

此外，PPJ 已允许显式 `displayBlanksAs: zero/gap/span` 和 series `nullHandling: zero/gap/connect`。这与“未经授权不得把缺失当 0 或连线”的可靠性规则需要明确衔接：

- 未声明转换策略时不能自行补 0、连线或推断观测。
- 用户显式要求 zero/connect 时，应记录这是显示策略，不是原始观测值；若当前可靠性验收禁止这种显示，应明确拒绝或标记人工复核。
- 不能静默忽略输入策略，也不能为了让图好看而选择另一种策略。

完成条件：各图表通道共享缺失值语义；有 all-null、首尾缺失、连续缺失、单点段、真实 0 与缺失并存、显式策略的回归。

## 5. 诊断、CLI 和输出证据差距

### 5.1 G-11：支持等级与真实输出不一致

当前存在四类不一致：

1. 类型声明称 supported，但有重要字段被忽略，例如 shape/image/group/connector。
2. heatmap、treemap、sunburst、waterfall、candlestick 在能力表是 partial，具体分支却推入 `status: supported`。
3. 总状态由“是否存在任意 diagnostic”计算：即使只有 supported 记录，总状态也会变成 partial。
4. 缺 sharp 的 unavailable 记录在总状态计算之后加入；总状态可能仍是 supported。

当前 diagnostics 通常只有元素 id、status、reason，缺页面、字段路径、具体被忽略的值及处置建议。顶层 bounds 检查只看原 frame，不检查嵌套子元素、旋转、阴影、文字溢出或 mask 后实际范围。

完成条件：按实际元素/字段支持情况聚合状态；supported、partial、opaque、unavailable 的含义一致。部分支持必须列出限制，错误事实不得以普通视觉 warning 掩盖。

### 5.2 G-12：输出安全和失败处理

修复进度：下面列出的初始输出问题已有实现与专项回归。新增 `src/ppj/preview-output.mjs` 独占创建目录和文件，在失败时保留实际产物；清单记录 hash、字节数、原始页 ID、输入/源/候选/资产身份以及 Node/栅格后端信息。源码版本摘要也进入清单。SVG 和 PNG 仅在写入成功后列出，缺少 sharp 返回不完整失败而保留 SVG；pending/final 清单与错误码区分失败阶段。

验证命令：`npm run test:ppj-preview-output`、`node test/ppj-svg-preview.mjs`、`node test/ppj-preview-capability-coverage.mjs` 均已通过。真实 smoke 同时验证 authored 和原生重新投影后的 source-bound 输入，原始 PPJ/PPTX 字节保持不变；模拟故障测试覆盖目标冲突、并发、路径映射、资源变化窗口、页面栅格失败及清单写入失败。

集成修复进度：新轻量回归已加入 fast/slow gate，slow 的 presentation 分段同时执行 SVG 预览与覆盖声明检查。`node test/gate-policy.mjs` 已改为校验当前 PPJ 入口、已退役入口不再出现及分段连续性，专项运行通过；本轮未运行完整 fast/slow gate。原始缺陷记录如下：

- `mkdir(..., recursive: true)` 加普通 `writeFile` 会复用目录并覆盖同名 SVG/PNG/render.json，没有继承外部 render 的独占输出保护。
- 写入是逐文件进行，栅格失败可能留下部分输出；缺失败清单、完成标记或原子交付策略。
- 缺少 sharp 时，render.json 仍为每页列出 PNG 文件名，实际文件可能不存在。
- 已记录 program/output hash，但缺每张预览图的实际 hash、sourceBound、输入/输出定位、渲染版本和依赖信息等完整关联证据。

完成条件：不覆盖用户输入和已有证据；失败结果能区分已产生/未产生的文件；输出清单只声明真实产物；预览必须绑定实际输入和资产。路径安全还需独立负向测试，本审计不推断未经验证的路径攻击场景。

### 5.3 G-13：CLI 和交付流程未完整接入

`preview` 复用了 render 参数解析，接受 `--pages`，但 handler 没有把页选择传入实现，实际仍渲染所有页。preview 返回值也没有专用 command/摘要分支，非 `--json` 时会回落成完整 JSON，包含 SVG 内容。

自有 preview 未自动进入 `check → build → render/review → imported re-import` 的交付证据链。已有 review 能力仍然有用，但不能把 preview 的顶层边界检查当作完整 review，也不能把外部 render 的证据标签套在 SVG preview 上。

完成条件：参数要么实现、要么明确拒绝；统一可用的 CLI 返回和错误码；在 Skill/文档中说明两条渲染路线及其证据边界；结构、视觉、编辑保真检查有独立状态。

## 6. 测试、覆盖台账和性能差距

### 6.1 本次实际测试结果

| 检查 | 结果 | 真正证明的范围 |
| --- | --- | --- |
| `node test/ppj-svg-preview.mjs` | 通过 | 单个 canonical fixture 能生成两页、稳定 ID 字符串存在、PNG 大于 1000 bytes、manifest 页数正确 |
| `node test/ppj-preview-capability-coverage.mjs` | 通过，30 个声明项 | 有部分类型名称声明；不代表 30 种功能绘制正确 |
| `node test/ppj-preview-smoke.mjs` | 42 输入，4 rendered，38 compilerRejected，0 rendererFailed | 4 个调用完成预览；失败数量按脚本错误文本分类，未独立计量各阶段；没有进行语义/视觉等价验收 |
| canonical 输出检查 | 确认组件未展开、决策形状丢失、箭头未画 | 这些缺陷在现有测试通过的同时真实存在 |
| 能力矩阵比对 | 初次比对不一致；写作期间再次读取已一致 | 确实观察到过漂移，随后被其他工作流更新；不再把它记成当前仍不一致 |

4 个生成预览的输入及程序返回状态：

| 输入 | 返回状态 | diagnostics 数 |
| --- | --- | --- |
| `examples/ppj/minimum.ppj` | supported | 0 |
| `test/fixtures/presentation/evidence-ledger-canonical.ppj` | partial | 2 |
| `skills/presentation-template-library/skills/artifact-template-aqua-impact-story/assets/references/reference.ppj` | partial | 1 |
| `skills/presentation-template-library/skills/artifact-template-simple-dark-mode/assets/references/reference.ppj` | supported | 0 |

这四项的状态是实现自报，不能替代本审计结论。canonical 的两个诊断仅为图表 supported 和组件 partial，没有报告决策形状缺失。

38 个被脚本归类为 `compilerRejected` 的失败按错误分组如下。这些错误指向 schema/source/compile 边界，但脚本没有独立的阶段计量，不能把分类名当成完整调用跟踪：

| 错误 | 数量 | 下一步 |
| --- | --- | --- |
| `ppj.source.staleProjection` | 30 | 对照原始 PPTX、当前 projector 和实际运行的 codec 版本追查；不能只删掉 source 绑定使之通过 |
| `ppj.schema.enum` | 6 | 定位具体字段和值，对齐合法 schema；本次没有完成逐文件根因分析 |
| `unsupported_ppj_compile_feature`，gapWidth 用于不适用图表 | 2 | 修正输入声明或核对适用边界，保留原案例及修复记录 |

不能仅凭错误分组断言“都是老模板的问题”。本次没有重建并核对所有 NativeAOT 版本，也没有重新冻结这些 fixtures。

上表和第 3.1 节 SVG 是本次工具输出的审计摘录。完整原始 stdout、临时 PNG/SVG/render.json 没有归档为仓库固定证据；本文不声称这些精确运行结果已具备跨环境复现包。后续正式验收应保存原始记录，而不仅保留汇总表。

### 6.2 G-14：测试覆盖不足

当前预览测试缺少“画对”的断言。饼图只有圆、形状缺失、双轴错误等均可能通过图片存在性检查。

覆盖测试的问题：

- 图表名称来自硬编码列表，再用 feature/path 的字符串包含关系匹配；不是遍历实际 schema 类型和字段。
- 把 supported/partial/opaque 合成一个 Set，只检查“声明过”，不检查等级正确。
- fixture 只检查 page 的顶层 elements，不递归 group/component；chart 被类型检查直接放行。
- 当前 `test/fixtures` 下只有一个 `.ppj` 文件；这不代表仓库没有其他测试生成 PPJ，只是该扫描并未检查它们。
- 三个 preview 专项脚本虽在 package.json 注册，但未列入 `scripts/run-test-gate.mjs` 的常规 gate。

smoke 还通过错误字符串正则区分 compilerRejected/rendererFailed，而非实际失败阶段；且仅当 rendererFailed 非空才失败。即使所有输入都在编译阶段被拒绝，也可能返回成功退出码。

完成条件：结构、数据语义、视觉、编辑保真分别有断言；正例必须确实进入渲染；预期拒绝独立列负例，不得用它们抵消正例失败。事实错误设硬门槛，不能被外观评分抵消。

### 6.3 G-15：能力台账尚未防止后续漂移

当前生成矩阵记录了 181 个 schema definitions、46 个 authored boundaries、278 个 native leaf kinds、151 个 Help API 等计数。这些是不同粒度的集合，不是可相加的“渲染功能总数”。

矩阵生成器的 `visualStatus` 根据编译 boundary 的 behavior 推导为 unreviewed 或 partial-or-opaque，并不是渲染验证结果。它还将 programJson 标为复用入口，却没有指出 canonical JSON 与已展开布局的区别。

写作期间矩阵中的 previewCapabilities 已与声明同步。这消除了本次观察到的快照差异，但没有修复类型声明本身的矛盾，也没有补上字段级绘制验证和自动防漂移检查。

完成条件：以 schema/能力 registry 为来源，建立“字段 → 语义 owner → 预览映射或明确边界 → 正/负 fixture → 断言 → 文档”对应关系。新增视觉字段必须触发维护检查；生成矩阵有 `--check` 或等效漂移 gate。不能仅要求开发者记住同步维护几份清单。

### 6.4 G-16：高性能和运行稳定性尚无验收

当前能确认 SVG→PNG 使用 sharp，不能据此宣称整个渲染器高性能。preview 每次先进行 PPTX 编译，随后又读资产、拼接整页 SVG，并在结果中保留所有页面 SVG；写 PNG 逐页进行。大量 base64 图片也可能在多个页面输出中重复。

尚缺：冷启动/热运行耗时、编译与绘制分阶段耗时、峰值内存、字体环境、长文和大图、复杂图表、多页大文稿、失败回收等实测。

完成条件：在固定环境下报告分阶段耗时和内存，以小型、常用和压力级文稿对比；先测量再决定缓存、批处理或并发。不在缺数据时承诺毫秒级速度或任意规模支持，也不为基准测试引入常驻服务。

## 7. 修复顺序与验收要求

以下是建议顺序，不代表已获得独立实施决策。P0 指会造成误判或破坏证据的当前问题，P1 指完成静态功能覆盖，P2 指性能与交付完善；P2 不等于最终目标可以不做。

| 阶段 | 关联缺口 | 交付结果 |
| --- | --- | --- |
| P0：先停止误报 | G-02、G-06、G-09～G-14 | 修复事实表达错误；未修部分准确报告；输出不覆盖、文件清单真实；失败分类正确 |
| P0：统一语义输入 | G-01、G-04、G-08、G-10 | 明确复用 compiler expansion/样式/数据归一化的边界，避免继续增加第二套猜测逻辑 |
| P1：基础元素 | G-02～G-07 | 文字、几何、图片、表格、线、组、组件、源预览的静态显示有真实回归 |
| P1：图表 | G-08～G-10 | 按当前 schema 全部图表及变体逐项核对比例、尺度、层级、方向和缺失点 |
| P1：维护门禁 | G-11、G-13～G-15 | 字段映射与测试自动对齐，Skill 交付区分证据类型，矩阵持续可验证 |
| P2：性能与打包 | G-12、G-13、G-16 | 依赖缺失与失败恢复可靠；命令行为明确；基准可重复 |

### 7.1 每项功能的最小回归组合

每个新增/修复功能至少提供：

1. 一个正常 authored 输入，确认编译成功且预览真正绘制目标内容。
2. 一个能揭露该功能风险的边界输入，例如缺失数据、负值、复杂关系、嵌套变换或透明边缘。
3. 对支持 source-bound 编辑的功能，从原始源字节重新投影，完成修改/删除/二次投影，并检查非目标内容与包修改范围。
4. 一项几何/数据语义断言和一项可审阅的渲染结果；不能仅检查输出文件存在。
5. 对明确不支持的输入，检查诊断、可见占位/源预览及符合既定策略的状态。允许保留的 opaque/partial 与必须拒绝的错误输入分别验收；后者必须失败。

测试可以小而明确，不要求每个字段都跑庞大的宿主验收。字段只影响数值格式，就验证该格式的真实输出；字段影响拓扑，就必须验证关系，而非图片文件大小。

### 7.2 总体完成标准

- [ ] 16 类元素和全部当前静态视觉字段已逐项归属，现有支持范围实现真实绘制。
- [ ] 所有当前 chartType 及 symbol/stream 等变体按数据语义绘制，或仅在源内容本就不透明的边界保留 opaque。
- [ ] 缺失数据、比例、方向、层级、主副轴、可见性均有硬门槛回归。
- [ ] source-bound 内容不被扁平化，实际修改后的预览与导出/二次投影证据对应。
- [ ] 图片范围、透明度和 effects 的不确定性被明确处理，不能静默改变主体。
- [ ] 支持状态、错误码、产物清单和文件 hash 与实际运行一致。
- [ ] 所有正向 fixture 都进入真实绘制；负向 fixture 的拒绝原因单独验收。
- [ ] 结构检查、渲染检查、编辑保真分别报告结果；没有任一项被外观分或编译通过替代。
- [ ] 字段级映射和生成文档检查进入维护 gate，新功能不会静默落后于渲染器。
- [ ] 本地性能、依赖缺失和失败恢复完成实测，达到另行确定的使用目标。

## 8. 与接口补齐、Skill 实验的关系

### 8.1 对数轴接口：在途实现，不能混算渲染完成

`ppj-chart-axis-log-base` 已有 proposal/design/specs/tasks，工作区包含 schema、proto/generated JS、C# axis codec、编译、投影和校验等改动，写作期间还出现了专项测试及 reference/registry 更新。任务清单从初次读取的 0/4 变为复核时的 3/4，说明另一路工作仍在推进。

本次审计没有执行其 NativeAOT 构建和专项验收，也没有核验清单勾选对应的原始测试日志，所以应记为“已有实现，完成情况由独立变更验收；未由本审计确认”，而非“只有提案”。当前进度以该变更的最新任务和实际测试证据为准，不能只凭勾选判断整个接口或渲染器完成。

目标是 value axis 的 logBase 创建、读取、修改和删除，含 numeric X 与 combo 副轴、token、非法值及所有权保护。具体跟踪见 [对数轴任务](../openspec/changes/ppj-chart-axis-log-base/tasks.md)。

该接口即使通过 codec 回归，当前 preview 也未消费其缩放语义。它是“接口更新后渲染需要同步”的具体例子，不应以接口完成代替渲染完成。本文不修改或验收其他工作流正在编辑的文件。

### 8.2 更广泛 PPJ/PowerPoint 差距

跨 family 图表、完整 ChartPart/workbook 所有权、复杂主题继承、SmartArt、动画/交互、布局求解器及宿主行为等更大范围，继续由 [PPJ/Kimi/PowerPoint 差距台账](ppj-kimi-pptd-full-ppt-gap-backlog.zh-CN.md) 跟踪。

该台账的历史基线和 bounded 完成声明不能直接当作本次最新全仓验收。本文核实的是预览实现与当前输入模型之间的差距，不宣称已经逐项重新验收所有 C# native leaf、Office Live 适配器或完整 PowerPoint 功能。

### 8.3 Skill 路由实验仍是独立验收

本次未重新冻结案例、准备六个外部 PPTX fixture、执行四路线作者任务、两轮盲评或人类校准。因此本文不能作为 Kimi 路由默认启用或原 Skill 实验完成的依据。

恢复实验时仍需核对：各路线 1→10 编辑任务、每场景的缺失数据/复杂关系/source-bound 输入、事实与视觉分开评分、硬门槛失败处理、报告与 tasks.md 一致。具体完成状态需在实验工作区重新确认；本文不把以前的计划当作已执行事实。

本机确认存在独立评测工作树 `/home/zenfun/mywork/OfficeKit-ablation`，跟踪入口为该工作树内的 `openspec/changes/presentation-skill-ablation/tasks.md` 和 `evals/presentation-skill-ablation/`。该绝对路径仅是本次定位信息，不是运行时依赖；在其他机器可用 `git worktree list` 查找相应工作树，再按仓库内相对路径访问。

该目录还存在 `report.v1.md`、`report.tri-route.v1.md`、`report.multi-route.v1.md` 和多个版本的案例/证据文件。文件存在不证明研究完成；恢复时应先确定本轮采用哪个冻结版本，不能合并不同版本的统计作为同一轮结果。

本次补读该 tasks.md，执行研究与分析报告部分仍有未勾选项；其作者任务文字仍写 Shared/Kimi 两臂，不能直接当作后来要求的四路线执行清单。这进一步说明需要先对齐任务、最新 cases/rubric 和实际 evidence，再恢复实验。本文未全面检查该工作树内的实验产物。

## 9. 复核入口

### 9.1 可重复执行的检查

从仓库根目录运行：

```sh
node test/ppj-svg-preview.mjs
node test/ppj-preview-capability-coverage.mjs
node test/ppj-preview-smoke.mjs
```

测试会在操作系统临时目录写预览；smoke JSON 打到 stdout。保存输出时应使用新的证据位置，不覆盖输入。若出现 `spawnSync rg EPERM`，应先解决运行环境权限，不要将环境错误记录成绘制语义错误。

检查结果应同时保留输入数量、真正渲染数量、前置拒绝原因、渲染错误、实际产物与诊断；还应补记具体运行的 codec、字体和依赖版本。上述三条命令不是总体验收的替代品。

### 9.2 主要代码与文档

| 文件 | 核对用途 |
| --- | --- |
| [svg-preview.mjs](../src/ppj/svg-preview.mjs) | `renderElement` 的真实绘制分支、状态汇总、输出和资源处理 |
| [svg-preview-capabilities.json](../src/ppj/svg-preview-capabilities.json) | 当前支持等级声明及其矛盾 |
| [ppj-v1.schema.json](../src/ppj/ppj-v1.schema.json) | 实际元素、图表数据和样式字段；不能根据渲染器字段反推 schema |
| [workspace.mjs](../src/ppj/workspace.mjs) | 资源读取、安全检查及 compile 转发 |
| [PpjAuthoredPresentationCompiler.cs](../native/OfficeKit/src/OfficeKit.Codec/PpjAuthoredPresentationCompiler.cs) | CanonicalJson、expansion/build plan 和编译回执的区别 |
| [PpjPresentationCompiler.cs](../native/OfficeKit/src/OfficeKit.Codec/PpjPresentationCompiler.cs) | source-bound 编译路径和回执 |
| [render-review.mjs](../src/ppj/render-review.mjs) | LibreOffice/Poppler 路线、独占输出和结构 review 边界 |
| [cli.mjs](../src/ppj/cli.mjs) | preview 参数、handler 与格式化返回 |
| [canonical fixture](../test/fixtures/presentation/evidence-ledger-canonical.ppj) | 当前测试通过但仍显示错误的具体输入 |
| [预览 smoke 测试](../test/ppj-svg-preview.mjs) | 现有文件/页数/稳定 ID 断言 |
| [覆盖声明测试](../test/ppj-preview-capability-coverage.mjs) | 声明检查范围及递归/字段缺口 |
| [全量 smoke](../test/ppj-preview-smoke.mjs) | 42 个输入的运行路径与错误分类 |
| [常规 test gate](../scripts/run-test-gate.mjs) | 预览专项检查是否进入持续验证 |
| [能力矩阵生成器](../scripts/generate-presentation-capability-matrix.mjs) | 计数和 visualStatus 的来源 |
| [生成的能力矩阵](presentation-capability-matrix.json) | 快照数据，不等同于已验证覆盖 |
| [Presentation 开发规范](../skills/presentations/AGENTS.md) | 功能、文档、预览、保真和测试的共同完成要求 |

维护本文时，应逐项更新证据和状态。修好某个 G 编号不代表相关类型的所有字段完成；只有达到第 7 节要求，才可以收回“整体未完成”的结论。
