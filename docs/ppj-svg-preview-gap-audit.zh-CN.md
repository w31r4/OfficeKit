# OfficeKit 本地 PPT 预览渲染器差距审计

SmartArt 所有权复核（2026-09-10）：已按源码确认的独占整图替换契约修正前轮过严的部件预期，改以实际关系目标验证图外文件、同页兄弟对象及无关关系保留。最终 `tmp/officekit-native-scene-paint-2MsoJ4/integration.json` passed，取代前轮部件断言失败；导入缓存缺样式/连接线仍未修复，不能宣称源图视觉通过。详见[当前差距 G-07](ppj-preview-current-gaps.zh-CN.md#g-07source-boundopaque-和静态检查边界)。

SmartArt 缓存进展（2026-09-10）：内部 authored verified drawing 已实际绘制；当前导入缓存缺完整样式/连接线，已明确 unavailable。真实源文字修改的严格部件保留断言仍失败，最终报告 `tmp/officekit-native-scene-paint-XyIRJd/integration.json` 为 failed，不能沿用前轮全通过结论。具体成功范围、失败和待审计所有权见[当前差距 G-07](ppj-preview-current-gaps.zh-CN.md#g-07source-boundopaque-和静态检查边界)。

形状轮廓进展（2026-09-10）：内部形状/路径已消费虚线、端帽和连接角，真实创建/源修改/像素/重新投影通过，报告 `tmp/officekit-native-scene-paint-chiuBR/integration.json`。精确 Office 虚线节距、主题和生产入口仍开放，详见[当前差距 G-04](ppj-preview-current-gaps.zh-CN.md#g-04样式主题和效果)。

段落边距进展（2026-09-10）：内部整段左边距及左对齐首行缩进已消费 native 坐标，悬挂和清零的真实像素/源编辑/重新投影通过；报告 `tmp/officekit-native-scene-paint-PHbQnD/integration.json`。列表、非左对齐首行缩进及生产路由仍开放，详见[当前差距 G-02](ppj-preview-current-gaps.zh-CN.md#g-02文字和形状)。

point 段落间距进展（2026-09-10）：内部行距、段前/段后已消费明确 point 值，真实 40→55pt 和前后清零的三行像素、源保留及重新投影通过，报告 `tmp/officekit-native-scene-paint-d2vSYk/integration.json`。百分比间距、完整排版和生产入口仍开放，范围见[当前差距 G-02](ppj-preview-current-gaps.zh-CN.md#g-02文字和形状)。

字间距进展（2026-09-10）：内部 signed fontSpacingPoints 已映射 SVG，继承与零覆盖、真实字形宽度、源编辑及等值 authored 像素对照通过；最终报告 `tmp/officekit-native-scene-paint-b252Er/integration.json`。仍是既有 495daafc 包与新 JS 的限定验证；kerning、排版及生产入口开放，详见[当前差距 G-02](ppj-preview-current-gaps.zh-CN.md#g-02文字和形状)。

文字格式进展（2026-09-10）：内部单下划线、单删除线和有符号基线偏移已通过合成及指定 495daafc 包的真实创建/源编辑/像素/重新投影检查，报告 `tmp/officekit-native-scene-paint-MUHC4L/integration.json` 为 passed。具体行为、首次拒绝原因和剩余边界已整合进[当前差距 G-02](ppj-preview-current-gaps.zh-CN.md#g-02文字和形状)；完整排版、其他装饰和生产入口仍未完成。

当前状态阅读入口：[当前能力与剩余差距](ppj-preview-current-gaps.zh-CN.md)（2026-09-10，`ec9a2ab4` 加已有工作区修改）。该文按功能整合本页后续实施进展，区分正式 CLI、内部 painter 和旧运行时证据。本页保留多轮审计与历史日志，未标明同一快照的“当前”描述不可合并为最新覆盖结论。

最新实施补记（2026-09-10，G-05 图片边框）：内部图片边框已消费 direct RGB、width、opacity、style、cap/join；使用已绘制遮罩轮廓，边框位于图片 clip 外层，保留外半边线。无 mask 时沿 frame；custom mask 遵守各路径 stroke=false。合成验证椭圆边框、半透明及显式零宽/零 alpha；未解析主题色保留失败诊断，虚线精确长度继续标注近似。真实创建和源编辑的 magenta 边框像素、内部绿色图片、fresh projection width=4、原源及非目标 ZIP 保留通过，整套报告 `tmp/officekit-native-scene-paint-Hyrhgw/integration.json` 为 passed。使用既有 495daafc 包及当前 JS，presentation 4/4 通过。主题颜色解析、效果、tile 及生产入口仍开放。

最新实施补记（2026-09-10，G-02/G-05 roundRect）：普通形状和图片遮罩共用圆角矩形绘制，默认 adjustment 从仓库 preset-geometry-profiles.json 读取。半径遵循该文件引用的 pinned docx4j preset 定义：短边乘以 pin(0, adjustment, 50000)/100000。合成测试验证非正方形、默认/0/25000/50000、负值和超上限的几何；真实 mask 创建、源编辑为 50000、再次编辑为 0、角落像素和 fresh projection 均通过。最终完整集成报告 `tmp/officekit-native-scene-paint-M9UI1y/integration.json` 为 passed，仍使用 495daafc 包和当前 JS。presentation 4/4、gate-policy、差异检查通过。默认 /tmp 因 ENOSPC 导致首轮发布和 codec 失败，改用仓库 TMPDIR 后重跑成功；未删除旧产物。其他预设、文字布局、效果和正式入口仍开放。

最新实施补记（2026-09-10，G-05 图片遮罩）：内部 painter 已消费 rect/ellipse/diamond 及 literal customMaskPaths，使用 frame 坐标的 SVG clipPath，和 source crop、alpha、外层变换组合。每个遮罩使用独立且可重复生成的 ID；未知 preset、未支持的 adjustments、冲突几何及无法解析的路径保留失败级诊断。椭圆、菱形、自定义菱形三种真实创建和源编辑的中心/角落像素、fresh projection、原文件及非目标 ZIP 保留均通过，裁切后的绿色内容留在遮罩内，角落透出蓝色底图。最终集成报告 `/tmp/officekit-native-scene-paint-QMZO1r/integration.json` 为 passed，使用既有 495daafc 包及本次 JS；合成、presentation 4/4、差异空白检查通过。其余 preset/adjustments、图片 tile/effects 和正式 CLI 接入仍开放，G-01 保持 7/15。

最新实施补记（2026-09-10，G-05 / G-01 3.3）：内部图片绘制已消费 native crop 的四边有符号千分之一百分比，正值裁切、负值透明留白，嵌套 SVG viewport 限制图片在 frame 内；原有外层旋转/镜像和 alpha 保留。非法边值或空 source rectangle 明确失败，crop+tile 仍显示未实现。合成覆盖左/上裁切、负值、零值、非法值及输入不变。真实 `495daafc` NativeAOT 包配合本次 JS 完成创建、原源 no-op、正裁切改负留白、fresh projection、像素与 ZIP 保留检查：蓝色底图从留白处可见，红绿内容映射正确，仅 slide1.xml 改变。最终报告 `/tmp/officekit-native-scene-paint-Pxd6KU/integration.json` 为 passed；presentation 4/4、gate-policy 通过。中间透明测试因底图误用弧形而失败，改为矩形后重跑通过。此进展取代下文“内部按 frame stretch、尚未消费 crop”的旧描述；mask、tile、效果和生产 CLI 接入仍开放。

最新实施补记（2026-09-10，G-06 / G-01 3.4）：内部 SVG 已消费 elbow 的 literal `bendAdjustment`，按有向水平跨度计算折点，保留 0、负值及超过 100000 的值。真实 NativeAOT 创建 25000、source no-op、源编辑至 75000 的 SVG 路径、折点 PNG 像素和 fresh projection 均通过；源字节及非目标 ZIP 成员不变，仅 `ppt/slides/slide1.xml` 改变。报告为 `/tmp/officekit-native-scene-paint-90ZBTJ/integration.json`，整套集成 `passed`。运行使用既有 `495daafc` 包（PPJ SHA-256 `d43dcf4ce8000aa50165277aabfa0ed906680184aa2797a16e9f6c6e13b37c17`）和本次 JS，未重建后续 C# 变更。下文“非默认 bend 未绘制”作为此前状态被本条限定进展取代；curved、完整导入变换、自动避障、生产入口仍待完成，任务勾选保持 7/15。

审计起始日期：2026-09-09；最新文档核对日期：2026-09-10。当前分支：`main`/`origin/main` 均为 HEAD `231467d8`，已包含对象锚点、连接线编辑、箭头宽度/长度、可编辑 bend 调整、非负百分比堆叠、literal `arcTo`、custom-geometry text rectangle、ordered guides、adjustment formulas、custom-path extrusion eligibility、connection sites 和 adjustment handles。工作区另有 `ppj-custom-path-references` 的未提交源码/schema/OpenSpec 变更。最后一个可复现验证快照为 `495daafc`（构建包及集成报告见第 1.3 节）；其后的 `190dcfec`、`01c1b28d`、`231467d8` 代码和该工作区变更均未计入该包或覆盖结论，旧二进制结果仍单独保留，不能混用。

注：上文的 `495daafc` 只代表从干净源树构建并实际加载过的单次验证包；并发变更使后续双构建 reproducibility 比较失效，因此不把它记作 reproducibility 通过。

**结论：本地预览可以运行，当前 NativeAOT → native scene view → 内部 SVG/PNG 的限定集成已通过（对象锚点 2/2、literal `arcTo`、表格、连接线、line、pie/doughnut、column/bar 及 source-bound 回归均通过）；但全面静态功能覆盖和正式 CLI 场景接入仍未完成。生产入口仍走旧 `programJson` painter，G-01 任务仍为 7/15，因此不能宣布整体渲染器完成。**

本文面向维护者和后续开发者，记录距离“简单、覆盖现有静态功能、可靠辅助结构与视觉 review”的差距。只盘点本地 PPT 预览；不将 Word、Excel、PDF 或 Live 宿主的能力计入渲染完成度，也不要求 PowerPoint 像素级复刻。

阅读入口：

- 判断现状：差距总表、第 1.3 节最新运行结果。
- 找到具体问题：第 2 节区分正式入口与内部 painter；第 3～6 节说明行为、影响、根因和完成条件。
- 安排开发：第 7 节的修复顺序、单项验收与持续维护要求；第 7.4 节列出可直接拆分的剩余工作和风险案例。
- 复跑和定位：第 9 节命令、代码及任务清单。
- 查旧记录：[历史验证附录](ppj-svg-preview-gap-audit-history.zh-CN.md)。历史绿色结果不覆盖当前失败。

目前最需要明确的事实：

| 范围 | 已经做到 | 仍然欠缺 |
| --- | --- | --- |
| 正式 CLI | `ppj preview` 能编译并生成 SVG/PNG/清单，已有事实诊断和非覆盖发布保护 | 仍读取 canonical JSON，没有接入实际编译场景；旧绘制错误仍存在 |
| 共用场景 G-01 | 编译采集、候选导入、归属、传输和适配已勾选 7/15 | 真实绘制完整性、生产切换、场景诊断/发布绑定及完整验收尚未完成；7/15 不是覆盖率 |
| 内部基础 painter | 有限几何、富文本、图片、组、Bezier/`arcTo` 路径、表格、显式坐标连接线和普通 line；当前 NativeAOT 集成已验证对象锚点 2/2 | 文字排版、主题/效果、图片裁切、更多类型及生产路由仍缺；不等于完整静态覆盖 |
| 内部 pie / doughnut | 单系列比例、角度、孔径、逐点颜色、径向分离及源编辑/删除有真实证据；全缺失正例已修正 | 多环、负值、完整标签/图例/继承仍缺；分离只提供显式 radial-review 几何，不承诺 Office 精确间距 |
| 内部 column / bar | 普通分组、普通堆叠及非负百分比堆叠有局部证据；新增两种百分比方向的创建/no-op/源数值修改、比例像素与重新投影通过 | 负值百分比、对数基线、混合/副轴、完整标签/继承尚缺；缺失或零总量不伪造百分比；仍未接入正式 CLI |
| G-11 / G-12 | 已完成各自限定的诊断与发布契约，任务分别 9/9、8/8 | 不意味着图形已正确；新场景的完整对接仍待 G-01 4.x |
| 验证状态 | `495daafc` 可复现快照包构建成功（7 files，105,957,056 bytes；PPJ SHA-256 `d43dcf4ce8000aa50165277aabfa0ed906680184aa2797a16e9f6c6e13b37c17`）；NativeAOT 场景集成 `passed`，`relationFailures=[]`、`customArcPath=true`、锚点 2/2；该系列复核的场景/分段检查与仓库 `tmp/` smoke 为 4 rendered/38 compilerRejected/0 rendererFailed | `190dcfec`、`01c1b28d`、`231467d8` 未进入该包/未重验；smoke 中 38 个输入仍在编译阶段拒绝；生产场景路由、第三方 PPTX/workbook、宿主、人类校准及性能未完成 |

G-12 使用方法与清单字段见 [预览输出说明](ppj-preview-output.md)，实施清单见 [ppj-preview-output-evidence](../openspec/changes/ppj-preview-output-evidence/tasks.md)。以上不是全部 gap 的完成声明。

G-11 后续实施记录（2026-09-09，9/9 任务完成）：[支持诊断变更](../openspec/changes/ppj-preview-support-diagnostics/tasks.md) 已从独立模块推进到真实绘制和发布。能力摘要和生成矩阵同源生成，按真实 schema 校验 16 类元素、16 类图表，另列变体和源绑定字段；不再把 shape/image/connector 等整个类型宣称为完整支持。当前类型声明见第 3 节，具体实现与验收见第 5.1 节。

诊断测试已接入 fast/slow；circular 全缺失 fixture 已在前轮修正，最近一次柱形/条形实施中完整重跑 `slow/presentation`，4/4 通过，具体结果见第 1.3 节。正式路线 authored 清单仍为 `partial / failed`，简单 source-bound 清单为 `opaque / requires-review`，两者都成功发布文件。这恰好说明检查通过、文件生成与画面正确是不同结论。G-14 的完整视觉覆盖和 G-15 的字段维护闭环仍开放。

### 差距总表

“未完成”表示仍有本编号的核心要求未满足，不表示整项完全没有实现。G-12 的限定契约和 G-11 的检查契约已完成，其余 14 类绘制、交付、完整覆盖及性能工作仍开放。这不是按功能数量或工作量计算的完成率。

| 编号 | 当前状态 | 主要影响 | 收口时必须拿出的证据 |
| --- | --- | --- | --- |
| G-01 共用解析与布局 | 7/15；当前包已通过有限 native scene → SVG/PNG 集成，生产入口尚未切换 | 当前 CLI 的组件与样式仍可能和编译后的 PPTX 不同；等价配对、完整字段和发布对接仍缺 | 完整配对案例的几何、数据、样式及实际 SVG 一致，并完成生产检查/发布对接 |
| G-02 文字与形状 | 未完成 | 带文字形状消失，富文本错分行，几何被替换 | 几何与文字同时存在；run、段落、字号和溢出有断言 |
| G-03 变换与可见性 | 未完成 | 旋转、镜像、嵌套组及 hidden 表达错误 | 嵌套坐标、旋转边界、可见性与实际展开结果一致 |
| G-04 样式与主题 | 未完成 | token、继承、背景和 effects 被固定默认值替代 | 样式解析优先级及具体填充、透明度、轮廓回归 |
| G-05 图片 | 部分修复 | 资源快照已统一；裁切、主体范围、mask 和效果仍不可靠 | 已知裁切和透明边缘案例；未知信息的明确诊断 |
| G-06 表格与连接线 | 内部表格/显式连接线和当前 NativeAOT 对象锚点、端点、箭头/类型编辑回归已通过；bend 字段可保留并编辑 | 生产 painter 仍画固定中线/简化表格；非默认 bend 路线、自动避障、任意 connection-site、复杂路线及宿主行为仍缺 | 非等宽与 span 几何、端点移动、锚点、箭头和 bend 路线；编译场景、SVG 与 source-bound 都要核对 |
| G-07 源绑定与 opaque | 未完成 | 源快照可能不能说明编辑后状态；静态检查范围不清 | 原包保留、目标修改、重新投影及快照有效性证据 |
| G-08 图表公共语义 | 未完成 | 数据映射、尺度、双轴与标签可能误导读者 | 每个数据点可对应正确通道、尺度、轴和标签 |
| G-09 图表类型 | 内部 line/pie/doughnut/column/bar 部分实现；整体未完成 | 比例、层级、累计值、流向及 OHLC 可能错误或消失 | 第 4.2 节每类图表至少一个真实语义反例转为通过 |
| G-10 缺失值 | 未完成 | null 被补 0、跨缺失连线，或独立观测被省略 | 真实 0、连续缺失、单点段及显式显示策略回归 |
| G-11 支持状态 | 本项检查契约完成，9/9 任务通过 | 已知事实错误变为失败；partial/opaque 保留限制和图片警示 | 字段/页/全局与发布结果一致；事实与视觉分开；不代替绘制修复 |
| G-12 输出安全与证据 | 限定契约完成 | 已保护已有路径、保留失败产物并绑定实际输入 | 已有故障注入与真实 authored/source-bound 发布测试；边界见第 5.2 节 |
| G-13 CLI 与交付链 | 未完成 | `--pages` 被忽略，预览结果未闭合各类 review | 参数行为、摘要输出、结构/渲染/编辑保真独立状态 |
| G-14 测试覆盖 | 部分修复；当前内部集成及常规分段通过 | 现有正例已覆盖有限路径/表格/连接线/line/circular/bar；全元素、全图表、复杂源和正例数量门槛仍待补 | 正例实际渲染、负例独立分类、语义及视觉断言 |
| G-15 持续维护 | 部分修复 | 类型和生成摘要已防漂移，circular 与普通分组柱/条形的 registry/tasks 证据已补记；字段级绘制约束仍缺 | schema/registry 到预览、fixture、断言的自动对应检查 |
| G-16 性能与稳定性 | 未验证 | 尚不能承诺日常多页、大图、长文的速度和内存 | 固定环境下的分阶段耗时、峰值内存及压力结果 |

## 1. 结论、范围和证据口径

### 1.1 当前结论

`officekit ppj preview` 已能读取 PPJ，调用现有编译器，然后生成 SVG、PNG 和 `render.json`。这证明了一条可运行的本地预览链路，但不能证明预览忠实表达了输入。

正式 CLI 的旧 painter 仍存在会改变内容含义的绘制错误：带文字的形状丢失几何本体、饼图不表达数值比例、连接线不依据端点关系、复杂图表读取错误的数据字段。这些问题不能用“只是视觉近似”解释。G-11 已为已登记的错误生成失败级可靠性和图片警示，未知视觉字段保守报告限制；诊断没有使图形本身变正确。

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

### 1.3 审计基线与最新复核

本节先记录本次文档核对，再保留各轮实施结果。后面的折叠区和[历史验证记录](ppj-svg-preview-gap-audit-history.zh-CN.md) 保留更早轮次。历史通过或失败不自动延续到新代码，各表的执行范围不可混用。

#### 当前分支核对：`231467d8`（2026-09-10，未重建）

- `main` 与 `origin/main` 当前同为 `231467d8`。该提交新增 custom-geometry adjustment handles（此前的 connection sites 已在 `01c1b28d`）；工作区另有 `ppj-custom-path-references` 的未提交源码/schema/OpenSpec 变更。生产 `svg-preview.mjs` 仍消费 `compiled.programJson`；内部 `preview-scene-svg.mjs` 则消费 native `PresentationPreviewScene`，两条路线的支持范围不能合并。
- `190dcfec`、`01c1b28d`、`231467d8` 以及该工作区变更没有进入任何已验证的 NativeAOT 包或集成报告；必须在分支稳定后重新构建并重跑场景、smoke 和受影响的字段回归。
- 当前 registry 仍是保守等级：16 类元素中 9 类 `partial`、7 类 `opaque`；16 类 `chartType` 全部为 `partial`，没有任何类型因内部 painter 的局部回归而提升为 `supported`。这组数字是声明分布，不是渲染覆盖率。
- OpenSpec 任务总计 24/32 勾选：G-01 为 7/15，G-11 为 9/9，G-12 为 8/8；G-01 的 3.3、3.4、4.1、4.2、5.1～5.4 仍未满足完整验收条件。

#### 最后一次已实际加载的验证快照：`495daafc`（2026-09-10）

- 按 checked-in build workflow 使用 SDK 8.0.128、仓库 `tmp/` 临时目录从干净 `495daafc` 源树生成 `/home/zenfun/mywork/OfficeKit/tmp/officekit-preview-runtime-495daafc`；构建退出 0，7 个文件共 105,957,056 bytes，PPJ 可执行文件 SHA-256 为 `d43dcf4ce8000aa50165277aabfa0ed906680184aa2797a16e9f6c6e13b37c17`，Office 可执行文件 SHA-256 为 `22a1c33c67e3ed15fbf2f3224013c0f0005a053404ba1819b3e9d67b7c8f2703`，manifest SHA-256 为 `55611aac24a355df9f520cc957cfeabf950051d23846c747f203420c7493dec4`。随后尝试 `npm run verify:office-kit-build` 时分支在两次构建之间继续变化，差异来自并发源变更，不能据此宣称 `495daafc` 的 reproducibility 已通过。
- 对该精确包执行 `node test/ppj-preview-scene-native.mjs /home/zenfun/mywork/OfficeKit/tmp/officekit-preview-runtime-495daafc` 退出 0。报告 `/tmp/officekit-native-scene-paint-MTxwlx/integration.json` 为 `status: passed`、`relationFailures: []`；`internalPainting` 显示 `customArcPath: true`、`generatedBezierPaths: 6`、`directedAnchorCases: 2/2`、显式坐标连接线箭头、合并表格像素及 source-bound 表格重新投影均通过。column/bar（普通、堆叠、非负百分比堆叠）、line、pie/doughnut 的 native/SVG/像素/源编辑回归也在该报告中通过。
- 该通过只覆盖“传输 → scene view → 内部 SVG/PNG”边界，不是生产 CLI 的场景切换；报告 scope 明确排除 production scene routing 和 complete paint coverage，也没有覆盖 PowerPoint 宿主、第三方 workbook/PPTX、人类视觉校准或性能目标。`495daafc` 之后的 `190dcfec`、`01c1b28d`、`231467d8` 代码未由该包验证。
- 与该系列复核关联的 focused checks（scene SVG、scene view、capability coverage、gate-policy、两个 capability generator `--check`、`slow/presentation` 4/4、`proto:check`、strict OpenSpec）在更早的干净快照已退出 0；仓库 `tmp/` 临时目录的全量 smoke 为 42 个 `.ppj`：4 个进入渲染（均 `partial`，diagnostics 44/686/3674/1032），38 个在编译阶段拒绝，`rendererFailed` 为 0。默认系统 `/tmp` 的一次复跑曾因 `ENOSPC` 产生 1 个 renderer failure，按环境资源错误处理，未混入上述代码结果。
- 这组证据仍不能完成生产 route、完整字段绘制、任务 3.3/3.4/4.x/5.x、第三方 fixture、人类校准或 G-16 基准；G-01 仍保持 7/15，不换算总体完成率。

源码入口：[锚点求解器](../native/OfficeKit/src/OfficeKit.Codec/PpjConnectorEndpointResolver.cs)、[源绑定编辑](../native/OfficeKit/src/OfficeKit.Codec/PpjConnectorSourceBoundCompiler.cs)、[专项测试](../native/OfficeKit/tests/OfficeKit.Codec.Tests/PpjConnectorObjectAnchorTests.cs)、[独立变更及测试记录](../openspec/changes/ppj-connector-object-anchors/tasks.md)。

#### 前次实施记录：非负百分比堆叠（2026-09-10）

该轮补 native BAR 的非负 `percent-stacked`；先缩放再求和，保留原值与显示比例。缺失总量和全零总量分别显示 Incomplete / Zero total，不生成虚假百分比。负值百分比仍失败；完整功能和生产切换仍未完成。该轮使用旧临时包，当前包的复验结果以上方最新复核为准；具体边界见第 2.9 节。

- 合成专项和完整 presentation 分段通过，后者 4/4，旧生产产物 `/tmp/officekit-ppj-preview-BTDmLF`。合成覆盖两方向、比例/累计、反向轴、明确范围、极大值、真实零、缺失、零分母、下溢/精度损失与负值拒绝。
- 该轮真实集成 `/tmp/officekit-native-scene-paint-l3j50c/integration.json`：两种百分比方向的 authored/no-op/原始源数值 `4→6` 修改通过；份额从 `4/12` 变 `6/14`，native/SVG/像素/重新投影一致，仅目标 ChartPart 改变。原始源和所有非目标 ZIP 成员不变，仍是无 workbook 的本项目 literal-data 文稿。
- 该轮普通堆叠及此前独立回归也执行通过；当时旧包的两个对象锚点失败 0/2，最终退出 1。该数值保留作历史反例；当前包已由上方最新复核验证为 2/2，不应继续当作现行失败。
- 初次产物 yqBN8Z 的测试分类误沿用 Negative 名称，改为 Other 后整套重跑，最后结果以上述 l3j50c 为准；不把复跑累计为新增成功数量。
- painter、合成测试和真实集成脚本的 hash 与旧包身份只作历史追溯；当前 PPJ 包身份和 scope 以上方最新复核为准。
- 该轮未改 C#/proto、未重建当前运行时、未做全仓/全量 smoke/外部 Office/第三方 workbook/人类校准/性能验收。G-01 当时为 7/15，本轮也未满足提升条件。
- 收尾 gate-policy、两个生成器、strict OpenSpec、77 个本地链接及差异空白检查通过；已查看修正标签后的最终 bar-percent-edited PNG，仅为 Agent 审阅。

#### 前次实施：普通堆叠 column / bar（2026-09-10）

内部 `barChart()` 新增 native `grouping: stacked`，正负累计分别进行；保留点值、累计起终点及真实零。缺失分类显示 Incomplete，已知值仍在证据中，但不伪造完整累计位置。百分比堆叠仍明确失败，生产入口未切换。实现边界见第 2.8 节。

| 本轮实际执行 | 结果与范围 |
| --- | --- |
| 内部 SVG 专项 | 退出 0；两方向、正负分别累计、反向轴、显式 overlap、缺失与真实零、输入不变、溢出和精度损失拒绝；既有回归继续通过 |
| 真实 NativeAOT 集成 | `/tmp/officekit-native-scene-paint-js6Kkn/integration.json`；两方向 authored/no-op/源数值 `4→6` 的 native、累计几何、RGB 像素、重新投影通过，只改变目标 ChartPart，原源及所有非目标 ZIP 成员不变 |
| 整体集成结果 | 两个 authored 对象锚点仍失败，0/2；其余独立断言执行，最终退出 1，报告为 failed。没有吞掉失败或修改其期望 |
| 常规 presentation | 4/4，退出 0；旧生产路线产物 `/tmp/officekit-ppj-preview-xTlM9i`，不是新场景的生产验收 |
| 维护与看图 | gate-policy、两个生成器 `--check` 通过；已查看 column authored 和 bar edited PNG，不是人类校准 |
| 首次合成错误 | 反向轴 fixture 用普通对象代替 protobuf message，序列化失败；改为 generated Axis schema 创建，未放宽断言 |
| 未执行/边界 | 未改 C#/proto、未重建运行时、未跑全仓/全量 smoke/外部 Office/第三方 workbook/人类校准/性能；使用既有临时包，仅证明已有原生字段与本次 JS 的限定行为 |

本轮 SHA-256：painter `04ab1e744e51f9b812d5f8218da61e3e3001ea02d5a350cb7f75bc3514b52171`，合成测试 `eaa0486a93dda55188b1936ca7b4940fd4ea35e31a12475b2f11216a4e4d9348`，真实集成脚本 `3dd8405e7a751d49cdf73ceb34983e858fc119b5ed24cbe67005aa6506dce3f3`。PPJ 二进制仍为 `9c97c6605c36c9e2b8218fe208a380745cc25b3b94742c4a620eb4ed21a496e9`。G-01 仍为 7/15，不提升生产支持等级。

#### 前次文档复核：`1fe503cd` 与普通分组工作区修改

以下“本次”仅指堆叠实施前的文档复核，哈希和未执行范围不覆盖上方新实施。

本次重新读取生产入口、内部 painter、对象锚点编译路径、三个 OpenSpec 任务清单及留存的真实集成报告。内部 painter 与两个测试文件的 SHA-256 均与下方柱形/条形实施记录一致；生产 `svg-preview.mjs` 与发布器仍分别为 `f98c59fb95478610748e4377b906c922156fef0e64898eb1659cad2d1d7986ff`、`0d5aca1b4c89a5ed43e599250414dedc0312f7076cc928359112759d05f374df`。

| 本次动作 | 核对结果 | 证据边界 |
| --- | --- | --- |
| 重跑内部 SVG 专项 | `node test/ppj-preview-scene-svg.mjs` 退出 0 | 给定 native 场景的有限几何及失败断言通过；不包含真实编译、生产发布或人类审阅 |
| 重跑两个能力生成器 `--check` | 均退出 0 | 当前摘要和矩阵未漂移，不是字段级视觉覆盖证明 |
| 文档链接与差异检查 | 主文档 77 个、历史附录 2 个本地文件链接均有效；`git diff --check` 通过 | 只核对本地目标存在性和差异空白，不声称本次重新核验所有外部来源 |
| 读取既有 `integration.json` | `/tmp/officekit-native-scene-paint-2he9uM/integration.json` 仍为 `failed`；柱/条形 10 个修改候选及 circular 的 10 个 explosion 候选有像素/重新投影成功记录，锚点为 0/2 | 本次没有重新执行 NativeAOT 集成；只是确认原始报告与文档一致 |
| 核对任务勾选 | G-01 为 7/15，G-11 为 9/9，G-12 为 8/8 | 不改变任务状态，不换算总体完成百分比 |
| 核对实际路由 | 正式入口仍消费 `compiled.programJson`；内部 BAR 仍拒绝非普通分组及对数轴 | 堆叠/百分比堆叠尚未实现，不能把研究或后续计划写成已支持 |
| 本次未执行 | presentation 分段、真实原生集成、C# 专项、NativeAOT 构建、全仓测试、全量 smoke、第三方评测、外部 Office、人类校准和性能基准 | 下方相应结果属于之前的实施轮次，不是本次新增验收 |

临时报告当前可读，但没有因此成为可移植的仓库 fixture。后续复现须按第 9 节重新生成，不能把机器上的 `/tmp` 路径作为运行依赖。

#### 最近一次实施：普通分组 column / bar，`1fe503cd` 加工作区修改

| 检查项 | 本次结果 | 证明范围与限制 |
| --- | --- | --- |
| 实现范围 | 内部 painter 新增 `barChart()`，复用原生 BAR 方向/数据及共享填充/轮廓；两个测试、registry、本文与 G-01 tasks 同步 | 未修改 C#、proto、正式 `svg-preview.mjs`、安装包或 Skill 路由；任务仍为 7/15 |
| 内部 SVG 专项 | `node test/ppj-preview-scene-svg.mjs` 退出 0 | 普通分组方向/反向、正负值、缺失与真实零、gap/overlap、点级填充/透明/无线条、明确失败及输入不变；完整执行原有 line/circular 回归 |
| 常规 presentation | `npm run test:slow -- --segment presentation` 退出 0，4/4 | 最终产物 `/tmp/officekit-ppj-preview-dQGuM6`；仍为旧生产路线的诊断、发布、真实 SVG/PNG 和声明检查，不是新场景生产接入 |
| 真实 NativeAOT 集成 | `node test/ppj-preview-scene-native.mjs /tmp/officekit-preview-runtime-xZSfO9` 最终退出 1 | `/tmp/officekit-native-scene-paint-2he9uM/integration.json` 为 failed；两个对象锚点仍为 0/2，未降低断言或吞掉失败 |
| 新增真实通过 | column/bar 各自 authored、双轴反向、source no-op；数值、gap=0、overlap=-100、删除 gap、删除 overlap，共 10 个源修改候选 | 每候选核对 native、SVG 几何、精确 RGB 像素、缺失位置空白、再次投影；仅目标 ChartPart 改变，源字节和所有非目标 ZIP 成员不变 |
| 继续通过的独立检查 | 基础/组件/路径/文字/表格/显式坐标连接线、普通 line 和两种 circular（含 10 个 explosion 修改） | 同一个最终脚本重新执行；不能与前轮相同案例重复累计为新增覆盖 |
| 初次错误与处理 | 首个新 fixture 的 series.fill 错用了字符串，编译报 `ppj.schema.type`；按 schema 改成 solid fill 对象后重跑 | `/tmp/officekit-native-scene-paint-vI7tSw` 是前置拒绝；中间 NslolO 成功局部证据由最终 2he9uM 替代，不冒充额外通过数量 |
| 维护检查 | gate-policy、两个能力生成器 `--check`、JS 语法、strict OpenSpec 和差异空白检查通过 | 不提升生产支持等级，不把生成摘要一致当作全部字段有绘制 |
| 看图 | 已查看最终 column authored、bar reversed PNG，确认零参考线和有符号方向可见 | 是 Agent 对限定图像的检查，不是人类校准或宿主验收 |
| 未执行 | NativeAOT 重建/双构建、C# 专项、proto、全仓测试、全量 smoke、Skill 全套、外部 Office/第三方 workbook、人类校准、性能基准 | 使用旧显式临时包，仅证明它与本次 JS 在这些已有原生字段上的行为；不能证明当前 HEAD 的全部新 C# 接口 |

本轮文件 SHA-256：

| 文件 | SHA-256 |
| --- | --- |
| 内部 painter | `217aa21f1419468edf43f7c341209919aea7692ae9fc1fc12ae43a03f0f16327` |
| 内部 SVG 专项 | `33864847b73ae7e8360c909749de16f36f350c955ba11b2198260a51fb8f748b` |
| 真实集成脚本 | `e9b05e35bc4383e9c260d50178be1274fdc3bd080ee8232296255743ea5a5680` |
| 指定包 PPJ 二进制 | `9c97c6605c36c9e2b8218fe208a380745cc25b3b94742c4a620eb4ed21a496e9` |

<details>
<summary>前次文档复核及 circular 实施记录（不是本轮新测）</summary>

#### 前次文档复核：`caa7cc95`

本次只完善差距文档，保留已有代码和任务修改，不实施渲染功能。已重新核对生产入口、对象锚点编译路径、三个任务清单和下表五个 JS 文件的 SHA-256；文件摘要与 circular 实施记录一致。已读取留存的 `integration.json`，确认两类 circular 共 10 个 explosion 源修改候选通过，两个对象锚点失败，整体 `status: failed`。这是读取既有原始报告，不是重新执行 NativeAOT 集成。

本次实际重跑 `node test/ppj-preview-scene-svg.mjs` 及两个能力生成器的 `--check`，均退出 0；主文档 77 个、历史附录 2 个本地链接均有效，`git diff --check` 通过。三个任务清单分别为 G-01 7/15、G-11 9/9、G-12 8/8，未修改勾选状态。未重跑 presentation 分段、真实 NativeAOT 集成、C# 专项、构建、外部 Office 或性能测试。`caa7cc95` 新增的图表阴影 scale/skew 接口未由旧运行时验收，且尚无对应预览绘制，详见第 8.1 节。

#### 前次实施测试：`dc5e725c` 加 circular 工作区修改

下表以及正文中未另标轮次的 circular 实施、presentation 分段和原生集成结果，均指 `dc5e725c` 加 circular 修改的实施测试，不指上面的文档复核。标明 G-11、前次或历史的记录仍属于各自原轮次；本次文档复核的新执行范围仅以上一小节为准。

| 检查项 | 2026-09-09 本次实际结果 | 证明范围与限制 |
| --- | --- | --- |
| 代码快照 | HEAD `dc5e725c`；保留原有工作区修改，继续内部 painter/两个测试并同步 registry、tasks、审计 | 未修改 C#、proto 或生产 SVG 入口；不推断 origin/main、已安装包同步状态 |
| G-01 任务 | 1.1、1.2、2.1、2.2、2.3、3.1、3.2 已勾选，7/15 | 3.3、3.4、4.1、4.2、5.1～5.4 未勾选；任务数不是绘制覆盖率 |
| 正式入口 | `renderPpjToSvg` 编译时仍只传 `includeNodeMap: false`，随后解析 `compiled.programJson` | 未请求 `includePreviewScene`，不调用内部 painter；正式 CLI 缺陷见第 3～4 节 |
| 内部 SVG 合成专项 | `node test/ppj-preview-scene-svg.mjs` 退出 0，整个脚本通过 | 全缺失合法正例/非法点样式拒绝例；新增 explosion 0/25/100/400、逐点零覆盖、旋转、边界与真实零不移动；完整执行后续测试 |
| presentation 常规分段 | `npm run test:slow -- --segment presentation` 退出 0，4/4 通过 | 诊断（包含内部专项）、发布保护、正式 SVG/PNG、声明覆盖全部执行；不是全仓测试或完整绘制通过 |
| 正式路线产物 | `/tmp/officekit-ppj-preview-FSn5Hk` | 真实 codec + sharp；authored 和 CLI 样例为 complete / partial / failed，简单 source-bound 为 complete / opaque / requires-review；人工视觉状态仍 requires-human |
| 显式 NativeAOT 集成 | `node test/ppj-preview-scene-native.mjs /tmp/officekit-preview-runtime-xZSfO9` 退出 1 | 报告 `status: failed`；两个对象锚点案例均失败，0/2。实际都为 `(0,0.5) → (1,0.5)`，详见第 3.5 节 |
| 集成实际通过的基础项 | minimum/canonical、1 组组件/显式配对、6 条生成 Bézier 路径、源文字修改、显式坐标连接线、合并表格像素及源表格移动重新投影 | 有 native/SVG/像素或重新投影断言；不同项目证据不同，不等于所有案例都有完整视觉与编辑生命周期覆盖 |
| 集成实际通过的普通 line | authored、source no-op、表格移动后图表不变、源数值修改、重新投影、两系列独立缺失、marker/轮廓/边缘与缺失区间像素 | 第一系列 `[2,null,0,4,5]→[2,null,3,4,5]`，第二系列不变；literal-data 输入，不含 workbook |
| 集成实际通过的 pie / doughnut | 每类 authored、比例交换、旋转、source no-op、源数值修改和再次投影；比例颜色像素通过，doughnut 透明孔像素通过 | `[1,null,0,9]→[9,null,0,1]`；角度 0/90，环孔 60/40。两类各只改变 `ppt/slides/charts/chart1.xml`，其余 ZIP 成员和源字节不变；详见第 2.6 节 |
| circular 分离编辑 | pie/doughnut 各自 authored/no-op，以及 point-zero、point-delete、series-zero、series-delete、both-delete，共 10 个源修改候选通过 | 每次从原始源投影重建请求；native/SVG/精确 RGB 像素与再次投影一致，仅目标 ChartPart 改变；无 workbook |
| 测试客户端生命周期 | 首次扩展运行遇到独立 wire 客户端空闲退出；改为每次比较创建并释放后整条重跑 | 原生客户端默认 1000 ms 空闲回收。没有修改生产规则、伪造响应或跳过断言；首次产物 KC5nP4 不作最终通过证据 |
| 内部产物 | `/tmp/officekit-native-scene-paint-e6xHTu` | 保存 SVG、PNG、逐案例 diagnostics 和 `integration.json`；报告同时保留独立通过与整体失败。临时产物不是已归档的可移植 fixture 包 |
| 使用的二进制 | 沿用显式临时构建包，加载器检查 manifest/hash | 本次没有重建 HEAD；当前 C# 新接口不能仅凭这份旧包的成功用例视作已验收 |
| 维护检查 | gate-policy、两个能力生成器 `--check`、相关 JS 语法、strict OpenSpec 均通过 | registry/tasks 已补 circular 与 radial-review 边界，没有提升生产等级或勾选 3.3/3.4 |
| 收尾检查 | portability 255 文件、reference-sync 333 文件通过；主文档 77 个、历史附录 2 个本地链接有效；差异空白检查通过 | 这些维护检查不增加已验证的绘制功能数量；任务仍为 7/15 |
| 看图与文档 | 已查看最终 authored pie 和 point-zero doughnut PNG；历史记录保留在附录 | 看图不是宿主验收或人类校准；上一轮独立读者复核不计为本轮视觉证据 |
| 本次未执行 | 全仓测试、全量 preview smoke、C# 专项、NativeAOT 重建与双构建验证、proto、Skill 全套门禁、外部 Office、人类校准和性能基准 | 不沿用历史结果来填补；也未重新核验独立 Skill 评测工作树 |

该次 circular 实施的代码身份如下，不是当前新增柱形/条形后的文件摘要。之后文件发生变化，必须重新运行对应检查；只记录 HEAD 不足以覆盖未提交修改。

| 文件 | SHA-256 |
| --- | --- |
| 生产 `svg-preview.mjs` | `f98c59fb95478610748e4377b906c922156fef0e64898eb1659cad2d1d7986ff` |
| 内部 `preview-scene-svg.mjs` | `ffd4614d264979be819aeb672a464f8c1939c18a50af627b15fb7eb21dd2b5bd` |
| 发布 `preview-output.mjs` | `0d5aca1b4c89a5ed43e599250414dedc0312f7076cc928359112759d05f374df` |
| 内部合成测试 | `dbdb343b50954cbe927bfa09393e17558cd0a5f89db415c30c368f95930a3aa7` |
| 真实 NativeAOT 集成测试 | `68b171578fd66039875d77b82ec69088066879b8072bcd350b3f4a5d4c0e5bb9` |
| 指定包 PPJ 二进制 | `9c97c6605c36c9e2b8218fe208a380745cc25b3b94742c4a620eb4ed21a496e9` |

</details>

证据分三类：**运行确认**表示指定案例在对应轮次实际执行；**代码确认**表示可直接定位的实现；**待验证**表示没有足够的案例证明。负例按预期被拒绝可以是测试通过，但不是正向绘制成功；反之，合成 fixture 不合法造成测试失败，也不能直接推断所有合法输入都画错。

术语：`authored` 指从结构化输入创建文稿；`source-bound` 指编辑绑定原始 PPTX、原生对象与所有权证据；`opaque` 指无法完整安全建模的原生内容；`canonical JSON` 是规范化程序文本，不承诺已展开布局；`fixture` 是固定测试输入；`smoke` 只检查基本运行路径。

### 1.4 当前如何使用，以及不能如何判断完成

可以使用本地 preview 辅助定位对象、查看部分布局和发现已登记的事实错误，但必须同时读取页面/全局诊断和可靠性，不能只看图片或命令退出码。导出 PPTX 的外部渲染、结构检查、编辑保真检查仍须各自保留。

| 看到的结果 | 可以得出的结论 | 仍需检查 |
| --- | --- | --- |
| `ok: true`、`output.status: complete`、退出码 0 | 请求的预览文件发布成功 | `reliability` 是否失败；图形是否正确；结构与编辑保真是否通过 |
| `reliability: failed` | 自动检查已发现已知事实错误或生产失败 | 按 reason/path 修正；不能用外观评分抵消，也不能作为正确结果交付 |
| `reliability: requires-review` | 有 partial/opaque 或尚未证明的状态 | 人工核对具体限制；涉及源对象时检查候选文件，不只检查原始快照 |
| `reliability: passed` | 当前已评估输入通过已实现的自动规则 | 不是全功能支持，也不自动完成宿主视觉与编辑保真验收 |
| G-11/G-12 专项通过 | 诊断和限定发布契约的回归没有失败 | G-01～G-10 实际绘制、G-13 交付、G-14～G-16 覆盖维护与性能 |

尚未完成的核心工作有三组：第一组是让预览使用编译器的实际场景；第二组是补齐文字、几何、图片、表格、关系和图表的真实绘制；第三组是让测试正例、交付检查和后续功能维护能持续证明结果。它们都属于原目标，不会因提供占位或风险提示而自动完成。

### 1.5 五条可靠性要求，目前落实到哪一层

这五条是功能完成的硬约束。Skill 写出要求、检查器发现问题、渲染器正确表达、交付完成验收，是不同进度，不能合并称为“已支持”。

| 用户要求 | 当前已落实部分 | 仍然存在的差距 | 对应编号 |
| --- | --- | --- | --- |
| 缺失数据不得当作 0 连线 | 生产普通 line 有有限拆段；内部 line 已验证缺失索引、真实 0、孤立点和源编辑，pie/doughnut 正常混合缺失案例有独立证据；登记的错误会降低可靠性 | 生产单点段仍被丢弃；若干图表仍通过 Number/num 将 null 转 0；内部含缺失的 zero/span 暂明确失败，尚未实现带证据的显示策略 | G-08～G-11 |
| 图表关系必须和数据拓扑一致 | 对已知比例、通道、层级、轴等错误有事实诊断 | 正式路线的饼图比例、Sankey 边、树层级、OHLC 通道、主副轴等仍可能错误；内部圆形图的单系列比例已局部修复，不能推广到正式 CLI | G-01、G-06、G-08、G-09 |
| source-bound / opaque 不得扁平化 | 复用编译流程与源身份；native 已从实际编辑候选采集，有限归属测试验证 owner、opaque 原内容与非目标 ZIP 保留；内部源表格移动及重新投影本轮通过 | 生产 SVG 还未消费候选场景；本项目生成的简单源输入不能替代第三方复杂内容的全面验证 | G-01、G-07、G-12、G-14 |
| 图片不确定信息保守处理 | 使用同一资产字节，空资源失败；有字段限制和有限 contain 映射断言 | 解码、裁切、透明边缘、主体范围、mask 与阴影没有完整绘制/验证，不能依据未经证明的主体估计改图 | G-04、G-05、G-11 |
| 输出前结构、渲染、编辑保真分别检查 | 已有独立入口、诊断和可核对的发布清单 | preview 尚未形成三项独立验收的完整交付链；文件发布成功仍可与事实失败并存 | G-11～G-14 |

最终验收应分别保留 `结构结果`、`事实可靠性`、`视觉审阅结果`、`编辑保真结果` 和 `文件发布结果`。这是验收记录要求，不声称当前 CLI 已输出这五个统一字段。任一适用硬门槛失败都应阻止“正确交付”的结论；新建文稿没有源编辑时可将源编辑保真标为不适用并注明原因，不能伪造一次 re-import 通过。

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

#### 阅读覆盖数字前，先区分模型层次

本文出现的三组数字并非同一个覆盖分母：

- **16 类 PPJ 元素**是作者输入语言的类型，包含 component、slot、text、chart 等。
- **9 类 native content**是编译/导入后场景的内容分支。组件可展开成多个节点，文字可附着在 shape 上，部分图表会生成 group 和矢量子节点；不能把原生类型数与作者类型数一一相减。
- **16 种 chartType**是图表的语义分类，此外还有 symbol 和 stream stacking 等字段变体。传递了 chart 对象，不代表各类型及数据通道都有正确 painter。

因此，适配层“9 类全部保留”的验收说明传输与对象视图的覆盖；第 3、4 节核对的是实际绘制。未来核对新增字段时，必须沿真实编译结果找到承载它的节点，不能仅按同名类型寻找绘制分支。

### 2.3 G-01：实际预览尚未使用完整的已解析绘制结果

`renderPpjToSvg` 读取 `compiled.programJson`，而 authored 编译器返回的是 `validation.CanonicalJson`。编译器另外持有 expansion/build plan，返回的 JSON 并不等于已经展开、排好布局、解析完样式的绘制场景。

初次审计运行 canonical fixture 后，返回的 JSON 中仍有 **1 个 `component`**。当前预览端仍自行读取 label/value、安排位置；它没有复用该组件实际编译出来的元素树。

影响包括组件 repeat、dataset 编码、styleRef、主题和 grammar token 等高层表达：PPTX 编译可能处理正确，预览却绕过这些处理重新猜测显示。

来源：[workspace 编译转发](../src/ppj/workspace.mjs)、[authored 编译回执](../native/OfficeKit/src/OfficeKit.Codec/PpjAuthoredPresentationCompiler.cs)、[预览入口](../src/ppj/svg-preview.mjs)。

完成条件：明确一个共用的解析/布局输出边界，预览消费编译器同义的结果；组件和 dataset 的高层写法与等价展开写法应得到等价预览。

#### 已形成的方案与尚未实现的部分

现已建立 [G-01 提案](../openspec/changes/ppj-preview-compiler-scene/proposal.md)、[设计](../openspec/changes/ppj-preview-compiler-scene/design.md)、[规格](../openspec/changes/ppj-preview-compiler-scene/specs/ppj-preview-compiler-scene/spec.md) 和 [15 项任务](../openspec/changes/ppj-preview-compiler-scene/tasks.md)。1.1、1.2、2.1、2.2、2.3、3.1、3.2 已实施并勾选；其他 8 项未完成。wire 已有只读场景回执，authored 采集 writer 实际输入，source-bound 导入最终候选，节点保留真实 owner；JS 传输、校验和适配已有指定运行时的分阶段验证，但生产 SVG 未消费场景，旧组件 label/value 启发式仍在执行。内部 painter 已消费部分场景，不能再笼统称为“没有场景到 SVG 实现”；也不能由此推断全部字段正确。

3.1 的当前实现分布在 [轻量 wire](../src/codecs/office-kit-ppj-wire.mjs)、[native 客户端](../src/ppj/native.mjs)、[workspace](../src/ppj/workspace.mjs) 和 [场景校验模块](../src/ppj/preview-scene.mjs)。普通 build/check 不请求场景；显式请求但回执缺失、版本不符或内容身份不匹配时明确报错，不回退成旧 canonical 绘制。只读校验不进行布局，也不能替代视觉字段支持判断。

[传输回归](../test/ppj-preview-scene-transport.mjs) 使用合成字节及替代 native 调用，证明两种 wire 读法、选项转发、显式 0/false/缺失状态、完整性失败与默认惰性加载。[真实运行时回归](../test/ppj-preview-scene-native.mjs) 覆盖真实进程、跨语言场景摘要及 authored/source-bound 回执；它只替换包定位，不伪造响应。此前旧包运行曾在汇总两个锚点失败后退出 1，作为历史反例保留；当前 `1bb41206` 包的同一脚本已通过，详见第 1.3 节。两个 C# 库入口另由 3.1 轮次的 39 个原生用例验证。打包后的 Office profile 不接受 PPJ，这是现有拆包契约，不是缺少场景功能。3.1 完成不等于场景到 SVG 全部正确。

3.2 的 [只读适配层](../src/ppj/preview-scene-view.mjs) 已加入，[专项回归](../test/ppj-preview-scene-view.mjs) 已进入常规诊断套件。`native` 保留原始完整对象，`frame` 等仅为机械单位视图；资产按 native ID 找到已验证的 MIME/hash 字节，不二次读文件、不在此层编码图片。原生内容类型来自 generated descriptor，未知字段及后代保留在证据中并报告限制。所有适配结果都是 `paintAssessment: unassessed`，不能由无适配诊断推出支持绘制。未知/冲突页面身份不猜映射，零 group 子范围不回填为外框。

3.3 正在实现：[内部 SVG painter](../src/ppj/preview-scene-svg.mjs) 已绘制部分真实 native 几何、文本、图片与组，直接使用编译器生成的路径，没有读 canonical PPJ 猜组件或图表。[专项测试](../test/ppj-preview-scene-svg.mjs) 与真实运行时证据见第 1.3 节。它明确名为 `officekit-native-scene-svg-internal`，每页保留 pending-integration 警示，结果不是正式发布 receipt；未绘制类型仍有可见占位。生产入口、旧启发式的移除和完整 G-11/G-12 集成尚未完成，所以 3.3 不勾选。这是迁移的中间状态，不是将未实现功能永久降为占位来缩小目标。

| 剩余任务 | 具体还差什么 | 最小完成证据 |
| --- | --- | --- |
| 3.3 基础绘制接场景 | 内部基础映射已运行；生产 painter 仍读取 canonical PPJ | 补齐生产切换与 label/value 猜测移除所需验收；保留已验证真实组件、路径、像素及未修文本/效果限制 |
| 3.4 其他绘制接场景 | 当前包的内部 table/connector/普通 line/单系列 pie/doughnut/普通及堆叠 column/bar、literal `arcTo` 有限映射及真实回归通过；其他原生 chart 和剩余 opaque 内容仍待绘制 | 保留表格、折线、圆形图、柱形/条形与源编辑证据，继续生产接入、复杂关系和其余类型映射；不能重新猜 dataset/grammar/拓扑 |
| 4.1 场景诊断 | 现有 G-11 判断基于旧绘制输入 | 归属路径和未知字段可定位；仅以直接绘制回归撤销旧事实错误规则 |
| 4.2 发布绑定 | 现有清单尚未记录 scene 身份 | scene 版本/来源/摘要与实际候选绑定，返回与落盘一致；保留 G-12 所有保护 |
| 5.1 等价案例 | 已有一组组件/显式配对和像素断言，缺完整配对组覆盖 | component/repeat/slot、样式、dataset 和图表成对比较几何/数据/样式，每对覆盖相关风险输入 |
| 5.2 实际运行时 | 源码及合成测试不能说明已安装二进制正确 | 仓库构建流程、manifest/hash、双构建复现及指定二进制的真实集成；保留诊断与发布门禁 |
| 5.3 场景成本 | 缺场景开关的真实尺寸、时延与保留内存对照 | 组件密集和 source-bound 样例分阶段记录；关闭时不保留场景，超限明确失败 |
| 5.4 文档及维护收口 | registry、输出说明及 review 指引尚未随实际绘制切换验收 | 以真实结果更新受影响材料，生成检查和相应 Skill 检查通过；不把其余 G 编号一起勾完 |

此前基础实施轮次的验证记录：

- 请求 field 7 为默认关闭的 `include_preview_scene`，结果 field 13 复用 native `PresentationArtifact`。通用/专用原生入口均拒绝 validation-only 或 projection 携带场景选项，不改变 canonical JSON/hash。
- `PpjPreviewSceneBuilder` 按 native descriptor 复制字段，保留顺序和显式 0/false；隔离源授权、删除请求、opaque 原始 XML 和 OLE 替换请求。collector 已验证按顺序独立复制逐页输入、拒绝重复/不完整收集；这批基础测试当时尚未覆盖真实 writer 接入。
- 场景有版本、来源、程序/候选 hash、节点归属和 native 资产身份，验证资产字节并生成确定性摘要。限制包括保留值/序列化字节预算、100000 个视觉节点及 128 层 message 嵌套；超限明确失败。
- 原生专项 16/16 通过，覆盖 9 类 native content 的 descriptor 驱动复制、可选值、源载荷隔离、输入不变、归属记录、摘要及预算负例。SDK 使用新恢复的 `/tmp/officekit-preview-sdk-hfWOmf/dotnet` 8.0.128，未改 global.json。这不是重建 NativeAOT 或实际场景绘制验收。
- JS wire 回归已进入常规诊断套件；presentation 分段 4/4 通过，真实旧预览产物在 `/tmp/officekit-ppj-preview-vygd9W`。`buf lint`、严格 OpenSpec、差异检查通过；重复生成 hash 一致。`proto:check` 的最终 git-diff 检查因未提交生成绑定退出 1，没有冒充全绿，也没有暂存文件来改变结果。

2.1 已完成：`PpjPreviewSourceFreeBuildPlan` 委托原 plan 的 `MaterializeSlide` 和前页需求，在 writer 的 `RecordNativeBindings` 回调复制实际 slide；authored 编译结束后将场景放入 receipt。关闭选项时仍传原 plan。该阶段的 9 个 authored 用例验证真实编译字节不变、实际母版/布局背景和显式 grammar 文字继承、vector/native 图表、缺失索引、页面只 materialize 一次和普通/Morph 页面释放。后续原始归属路径与 repeat 区分的补充见下方 2.3。

2.2 已完成：`PpjPreviewCandidateScene` 对 source-bound 最终输出调用原生 `PptxCodec.Import`，不走会恢复私有 PPJ 的 projector。no-op 使用原始字节或原文件句柄，保持 `ReuseSourceFile` 且不强制 materialize；其他路径读取实际 edit-plan/export 输出。资产按 native ID 保留引用，按 MIME/hash 复用已有字节或补入实际候选资产。2.2 完成时的 6 个候选用例验证新导入等价、目标内容变化、源和非目标内容保持、opaque 未扁平化、文件复用及不恢复私有快照；2.3 又增加了 5 个归属专项用例。

2.3 已完成原生归属任务。source-bound 绑定已不再全部 unmapped：`PptxCodec.Import` 可按选项保留只读物理身份，projector 将它与语义 ID 对齐，`PpjPreviewCandidateBindings` 再用源部件路径与原生对象 ID 连接最终候选。普通导入默认不采集该列表。页内序号或画面位置不参与归属猜测；不同页复用相同原生 ID 时必须由 part 区分。

物理身份只是查找证据，不是编辑授权。binding 的 `NativeId` 仍是场景 IR 的字符串 ID，不等同于匹配过程中使用的数值 `cNvPr` ID。原始 OOXML、source 授权和 opaque 载荷仍不进入只读场景。已知 owner 使用请求中实际的 PPJ 路径；无法匹配时保留 unmapped，并仅给出能够确认的页面路径或根路径 `$`。

| 2.3 子范围 | 已实现并验证 | 保留的边界 |
| --- | --- | --- |
| source-bound 普通对象 | no-op、文字/位置叶编辑和语义 fill 的 semantic/page ID、路径及 z-order；嵌套组真实 readingOrder 与原始数组路径分开 | 复杂第三方文稿的实际视觉仍属 G-07/G-14，不能由身份断言代替 |
| 页面及元素生命周期 | 反转页、交换元素、跨 part 同 ID、删除后剩余对象路径；追加 overlay 保留实际 native ID/scenePath 和 unmapped | 追加没有源 owner 的对象不伪造归属；克隆和第三方复杂拓扑不据此宣称完成视觉保真 |
| 缺失或歧义身份 | 匹配器拒绝重复物理键、零 ID 和错误 part；普通导入不采集，开启采集不改变导入 Artifact | 匹配器单测不等于畸形/冲突第三方包的全链路回归 |
| authored 原始路径 | `PpjPreviewOrigins` 跟随实际 clone/slot 替换记录原始定义或实例 slots 路径；两层 repeat、嵌套组件和替换后序号变化均可定位回原输入 | 捕获默认关闭；内部预验证组件调用若未在原始展开时采集，明确报 `ppj.preview.originsRequired`，不重复展开或伪造路径 |
| 生成节点的归属类型 | 组件展开节点及 writer 生成子节点标为 Generated；保留实例/重复标识，矢量子节点定位到真实 chart owner | Generated 不是独立可编辑的 PPJ 子节点，也不证明对应图形已经画对 |
| 资产 | collector/candidate 测试已有 native ID、MIME/hash 与实际字节断言 | 资产正确不代表图片 crop、alpha、边缘或效果已经画对；后者属于 G-05 |

以上 2.1～2.3 完成的是 C# 场景生产与身份层。原始 canonical JSON、原有 node map 和候选 PPTX 均有开关前后字节不变断言；旧 `#component(...)` expansion 路径没有被改写成新的可编辑契约。当前 NativeAOT 包的真实场景传输、适配和内部 SVG/PNG 集成已通过；待完成的是生产入口消费、新场景图片警示/清单绑定、完整字段绘制和复杂第三方源验证。G-11/G-12 已有的警示与发布清单继续有效，不能因切换输入而撤销。

方案复用 C# 已有 `PresentationArtifact`，不再定义一套作者语言，也不在 JS 中重新解析 OOXML。默认关闭的只读场景选项已加入 wire，`programJson` 保留既有含义。下表说明整条方案及其要求，各环节的当前状态以上文任务表为准：

| 环节 | 为什么必须做 | 完成证据 |
| --- | --- | --- |
| 协议与场景封装 | canonical JSON 不是已展开、已解析的绘制树 | 版本化结果、默认关闭、区分字段缺省与显式 0/false；通用 `CodecProtocol` 与专用 `PpjCodecProtocol` 均校验，生成绑定同步；场景缺失/版本不匹配明确失败 |
| authored 场景采集 | 第二次布局可能与 writer 实际输出漂移 | 从 writer 同一次 `MaterializeSlide` 获取实际节点；场景开关不改变候选 PPTX 字节，不额外重复展开 |
| source-bound 候选导入 | native-leaf 快速编辑后的模型可能仍是旧状态，嵌入 PPJ 快照也不能代表最新内容 | 用现有 C# importer 读取最终候选；覆盖 no-op、叶编辑、语义导出与文件复用，不改写候选或原包 |
| 节点与资产身份 | 展开组件、矢量图表子节点及 native 资产 ID 不一定与原 PPJ 一一对应 | 分开记录原 PPJ 归属路径、场景路径、PPJ 元素 ID 与原生对象 ID；资产 hash/字节一致，未知归属明确报告 |
| JS 消费与真实绘制 | 只传场景但继续画 canonical JSON，不能解决根因 | 仅做单位与身份等机械转换；真实 SVG 使用场景几何、样式和通道，移除组件和数据拓扑猜测 |
| 诊断与发布对接 | 场景存在不等于文字、效果或原生图表已经画对 | 保留 G-11/G-12；只有直接绘制回归证明修复后才撤销旧限制；清单绑定 scene 和候选身份 |
| 完整性、成本与运行时 | 大场景不能静默截断，源码测试不能代替实际装载二进制 | 节点/字节预算超限明确失败；场景开关成本对比；构建并验证 JS 真正使用的 NativeAOT |

场景必须保留尚未绘制的原生视觉字段；为了减小返回值而仅保留矩形、文字，会再次造成语义丢失。源包和 opaque 二进制载荷不能作为第二份完整副本重复传输，其边界与身份仍须保留。

最低等价验证包括组件与显式展开、repeat/slot、命名样式与 grammar、dataset/encoding、矢量和原生图表，并包含缺失数据、复杂关系或源绑定边界。必须对照有效几何/数据/样式及实际 SVG，不能以“两边都生成文件”或“两边都是同样占位”判定成功。

G-01 即使完成，也只解决共用输入边界；文本排版、preset、效果和各图表 painter 的剩余问题仍按 G-02～G-10 验收。

### 2.4 内部 painter 的实际覆盖：哪些已画、哪些还没画

本表只描述 `paintPpjSceneSvg`，不是 `officekit ppj preview` 的对外支持承诺。当前每页都带 `preview.scene.paint.integration-pending` 警示；普通内容绘制、字段限制、事实判断、正式发布还没有组成完整生产链路。合成测试能证明“给定 native 字段如何变成 SVG”，不能证明编译器一定会从合法 PPJ 产生这些字段。

| 内容 | 已有映射及直接证据 | 仍缺的功能或验收 |
| --- | --- | --- |
| 形状与路径 | rect/textbox/process、ellipse、decision；形状和文字同时保留；literal M/L/C/Q/Z/`arcTo` 路径按 native viewport 转换，实际 Sunburst/Sankey 生成路径有运行断言，当前 NativeAOT 还验证了不等半径弧线；native 已继续暴露部分 preset/adjustment 字段 | preset/adjustment 的生产绘制、路径引用/arc 引用、复杂填充和效果仍缺；不能把少数路径成功算成整个图表类型完成 |
| 富文本 | 保留段落与内联 run，字号、字体、RGB、粗斜体、显式换行及透明度有 SVG 断言 | 字体度量、自动换行、AutoFit、溢出、完整段落/列表/继承；语言、baseline、strike、字距、kerning、highlight、shadow/glow 等新增字段未完成显示验证 |
| 图片 | 使用已验证资产字节，native frame、alpha、旋转/镜像进入 SVG；不再自行按 PPJ 路径重载资产 | 当前按 frame stretch 显示；裁切、mask、边缘和效果仍不完整；保留资产不等于保留实际主体范围 |
| 组与可见性 | childFrame 到父 frame 的缩放/平移、嵌套变换、元素 hidden 有断言；非法零子范围明确失败 | 未证明全部嵌套/连接组合的视觉边界；页 hidden 只是保留状态，缺页选择策略和 CLI 集成 |
| 表格 | native 列宽/行高按实际比例映射到 frame；横/纵合并、直接 RGB/no-fill/alpha、单元格文字、显式 0/false 有合成断言；四边线型映射中已直接核对右边框。真实 authored/no-op/移动候选的网格、文字、右边框、像素及重新投影通过；畸形/重叠 merge 与被覆盖的可见文字明确失败 | 继承样式、banding、渐变/图片填充、完整排版和合并外围边框仍缺；真实正例是简单本项目源输入，不能代表任意第三方表格 |
| 连接线 | native 有向坐标生成 straight；elbow 使用中点折线；目标/site 字段保留为证据；有限箭头方向/类型/尺寸及线样式有合成断言；当前 NativeAOT authored 对象锚点、目标移动、source-bound 箭头/类型编辑均通过；native 已暴露部分 connection-site 字段 | curved 明确 unavailable；导入 elbow 的原始旋转/adjustment 信息不足，dash 与箭头精确轮廓只做近似；任意 connection-site 的通用场景绘制/编辑、自动避障和宿主行为不是已解决范围 |
| 原生 chart | 普通分类 line；单系列 pie/doughnut；普通分组、普通堆叠及非负百分比 column/bar 有局部回归，见第 2.5～2.9 节 | 其他 native 枚举仍占位；负值百分比、折线/面积堆叠、平滑、混合/副轴、多环尚缺；explosion 仅径向 review 布局；轴、标签、图例、主题和完整显示策略未完成 |
| diagram、媒体、opaque/OLE 等剩余内容 | 适配层保留 native 内容、归属和可用的缓存绘制树；内部 painter 对未接入分支显示限制/占位 | 尚未完整消费缓存树、海报或源预览；不能把占位当成现有静态内容的完整绘制，更不能生成新的编辑授权 |
| 诊断与发布 | descriptor 驱动未消费字段检查；保留 scenePath、输入不可变性及页面警示 | 仍缺 G-11 的完整事实规则按 renderer profile 适配、G-12 scene 身份清单和故障集成；旧规则不能一次性撤掉 |

表格的源样式边界尤其需要保留：候选导入若没有直接 fill，当前报告 `table-inherited-fill`，不套用 authored 默认白色；合并区域被覆盖的物理单元格若有 fill/border，会报告 `merged-cell-style`，尚未完成合并外围边框解析。这些诊断是保守限制，不是样式已经还原。

此前内部代码声称“端点已经包含编译器锚点解析”，此前实施轮次已纠正该注释。准确结论是：painter 使用回执中的有向端点；当前包由编译器产生的两个 authored 对象锚点端点均已通过独立断言，旧包的失败只作回归历史。painter 仍不自行求解对象关系，也没有因此获得自动避障或完整 connection-site 语义。

### 2.5 内部原生折线图：新增覆盖与明确剩余项

这是 G-01 3.4 的内部实现及回归进展，不是正式 CLI 新增支持。入口为 `paintPpjSceneSvg` 的 `chart()`；本节描述 native `SpreadsheetChartType.LINE` 分支；目前另有 PIE/DOUGHNUT 和 BAR 分支，见第 2.6～2.7 节。其他 native 枚举仍返回带诊断占位。某些 PPJ 图表由编译器降低为 group/path，走第 2.4 节的几何分支，不受这个枚举分支直接代表。

| 行为 | 当前实现与证据 | 不能扩大的结论 |
| --- | --- | --- |
| 数据通道 | 读取 native `categories`、`series.values` 和 `missingValueIndexes`；数量不符、非有限值、无序/重复/越界索引或非 0 缺失占位明确失败 | 不接受把 `xValues`、`bubbleSizes` 混入分类折线，也没有在 JS 重算 PPJ dataset/encoding |
| 缺失与真实 0 | 缺失索引决定断线；占位 0 不参与自动数值范围；真实 0 保留。合成测试含单点、全缺失、首尾/连续/间隔缺失、负数和不同缺失位置的双系列；真实 `[1,null,0]` 保留两个独立观测，双系列端到端检查各自断线 | 尚未证明全部缺失组合、其他图表通道或显式转换策略共享正确规则 |
| 孤立点 | 有可映射 marker 时绘制 marker；否则用虚线空心 review 点显示独立观测，同时产生 `chart-isolated-review-point` 诊断 | review 点是审阅标记，不冒充作者指定的图形，也不等于完整 marker 保真 |
| 显式显示策略 | 默认 gap；带缺失的 zero/span 明确 `chart-semantics` 失败，避免伪造观测或偷偷忽略策略 | 这是尚未实现的边界，不能把“拒绝已有合法策略”算成该策略已支持；后续需要保留原值和显示变换证据 |
| 轴与尺度 | native minimum/maximum 保留显式 0；有限负值、轴反向及 logBase 有合成几何断言；对数轴出现真实非正值失败。分类轴等距；轴隐藏、单独隐藏轴线/刻度标签及分类标签间隔新增直接断言 | 只有简单分类 X/数值 Y；主副轴、数值 X、刻度算法、格式、交叉点、网格和完整轴样式未完成；有限可见性案例不代表全部继承组合 |
| 标准分组 | 接受 native grouping `none`；这是 codec 将 ChartML standard/clustered 归一后的值，不是 PPJ 输入字符串 | 堆叠/百分比堆叠需要累计坐标；smooth 需要插值，当前均明确失败，不能画成普通折线 |
| 线与 marker | 有限直接 RGB、线宽/透明度及基础线型；dot/circle/square/diamond/triangle 有分支，部分有直接 SVG/像素断言。新增 marker 直接轮廓的 RGB/宽度/透明度/线型/cap/join；零宽/零透明度保留，真实圆点轮廓像素通过 | 主题继承、全部线型精确外观及所有符号形状未完整验证。缺直接色/线宽时使用 review 默认并报告 `chart-inherited-paint`，不是还原主题 |
| 标题与布局 | title/titleBody 进入富文本路径；保留 series/point/category/value 证据；线段按 plot 裁切，范围内 marker 只受 chart frame 裁切。四侧边缘圆点有真实像素断言；超界点保留真实值及 outside-plot 标记，不钳制成边界观测 | 固定边距和三个 value ticks 不是 Office 布局；图例、标签格式、碰撞、完整 title/run 效果未完成。大 marker、小图框、嵌套/旋转下的全部裁切组合仍待验证 |
| 源绑定编辑 | 从原始源重新构造编辑；只改第一系列第三个值 `0→3`，候选绘图及重新投影为 `[2,null,3,4,5]`，第二系列 `[5,4,null,0,1]` 不变。只有 `ppt/slides/charts/chart1.xml` 字节改变；文件集合、所有非目标 part、原始源字节不变 | 只是一份本项目创建后去掉私有快照的 literal-data 源输入，不含 workbook；未验证任意第三方 line、workbook 联动或删除生命周期 |

还没有专门绘制趋势线、误差线、逐点样式、数据标签、plotArea/frame paint、完整主题及 chart 文字效果。当前 `chart-layout` 和每页 `integration-pending` 限制始终保留；通过这些局部测试不应提升生产能力等级，也不能撤掉旧 painter 的 G-11 事实错误规则。

前次 line 实施已在 registry 的 `previewScene.contract` 补记有限 line、marker 轮廓及边缘显示，并向 G-01 tasks 追加证据，没有提升生产支持等级或勾选整项。本轮已同步 registry/tasks 的 circular 和径向分离进展，仍未提升等级或勾选整项，字段维护边界见第 6.3 节。两个生成器检查通过只证明生成内容与 registry 一致，不能代替字段级绘制覆盖；完整维护闭环仍由 G-15 跟踪。

下一步验收仍分三层：补齐普通 line 的剩余语义与风险例；按真实 native 类型或编译生成几何逐项补其他图表；最后把场景绘制接入按 renderer profile 区分的检查及正式发布清单。G-08、G-09、G-10 和 G-01 3.4 均保持未完成。

### 2.6 内部饼图 / 环图：已实现比例，未完成整类支持

当前 `chart()` 将 native PIE（3）和 DOUGHNUT（5）交给 `circularChart()`。这解决了内部单系列正数比例绘制的问题，**没有改变正式 CLI 的单色圆兜底**，也没有完成 G-09。

#### 输入、算法和实际影响

| 项目 | 当前代码行为 | 边界 |
| --- | --- | --- |
| 数据来源 | 只读 native categories、series.values、missingValueIndexes；与 line 共用 `categorySeries()` 校验 | 不在 JS 重新解释 PPJ dataset、encoding 或原始 OOXML |
| 适用拓扑 | 单系列，无 Cartesian 轴、combo 或副轴 | 多系列/多环明确 unavailable，不能当成普通单环画掉其他系列 |
| 数值比例 | 使用已知非负观测；先按最大值缩放，再求总量，按占比生成 SVG 扇区 | 缺失不加入分母；比例只代表已知正值，不代表包含未知项的总体；负值和无法表示的极小比例明确失败 |
| 缺失与真实 0 | 缺失节点保留 missing 标记；真实 0 保留 point/value/fraction=0，不画有面积的扇区；附缺失/零值计数 | 含缺失的 zero/span 策略尚未实现，不能悄悄补 0 或删除策略 |
| 角度与孔径 | 消费 native firstSliceAngle；内部按正上方起顺时针计算。doughnut 孔径以外半径百分比计算，缺省 50，接受 10～90 | 起始角仅接受整数 0～360；pie 不接受 doughnutHoleSize；这里记录实现约定与已测角度，不是全部宿主布局保真 |
| 扇区路径 | 外弧拆成两段；环图增加反向内弧，中心不涂白 | 可保留底层对象透出；完整一周也有几何分支，所有退化/嵌套变换组合未全验 |
| 填充与轮廓 | 逐点 fill/line 优先于系列；直接 RGB、noFill、透明度及有限轮廓映射 | 无法解析的继承/渐变用带诊断的 review 样式；调色板不冒充实际主题 |
| 逐点归属 | 点样式索引必须递增、唯一、范围内且指向真实观测 | 不排序“修复”输入；不得把缺失点样式挪到邻点 |
| 分离扇区 | 0～400 的系列/逐点值进入径向 review 绘制，显式点 0 覆盖系列；保留原值、归属与偏移 | 统一缩放图半径以容纳偏移；不承诺 Office 精确间距或同心分离布局，始终保留 `chart-explosion-layout` 限制 |
| 标题与版面 | 标题进入现有文字路径；保留数据索引、数值、比例和角度属性 | 固定 plot 边距；标签、图例、碰撞、字体度量、主题和效果仍不完整 |

`chart-layout`、`chart-missing-share` 及每页 `integration-pending` 等限制仍保留。这里的正例不是 supported 整类声明。

#### 本次真实成功的范围

[test/ppj-preview-scene-native.mjs](../test/ppj-preview-scene-native.mjs) 对 pie 和 doughnut 分别执行：

1. 创建 `[1,null,0,9]`，确认 native 为 values `[1,0,0,9]` + missing index `[1]`；SVG 中缺失与真实 0 不混淆。
2. 改为 `[9,null,0,1]`，第一扇区从 10% 变 90%，用扇区内精确 RGB 像素确认画面真的变化。
3. 起始角由 0 变 90；doughnut 孔径由 60 变 40，核对 native、SVG 与采样像素。环心像素为底层矩形的 `#FFEEDD`，不是白色覆盖。
4. 从本项目生成的 PPTX 去除私有 authored 快照，重新投影为 source-bound；no-op 保持源字节完全一致。
5. 从这份原始源投影构造数值编辑，绘制候选并再次投影，得到 `[9,null,0,1]`。
6. ZIP 成员集合一致，只有 `ppt/slides/charts/chart1.xml` 改变，其他 part 和原始源字节不变。

这些步骤在此前 circular 实施及最近一次柱形/条形实施的显式临时包集成中均通过，本次文档复核仅读取最后一份报告。它们验证的是本项目产生的单系列 literal-data 文稿，**不含内嵌 workbook，也不是外部复杂 PPTX fixture**。没有验证图表/数据点删除、系列增删、多环、workbook 联动或任意源主题；已验证的 explosion 属性删除单列于下文，不能推广为对象或系列删除。报告 pie 的 `transparentHole: false` 表示该项不适用于 pie，不是环心测试失败；doughnut 的该值为 true。

#### 已修复：全缺失合成 fixture 冲突

上一轮失败位置为 [内部合成测试](../test/ppj-preview-scene-svg.mjs) 的“全缺失”案例：fixture 把 values 设为 `[0,0,0,0]`、missing indexes 设为 `[0,1,2,3]`，但保留了模板中的 index 0 和 3 的 pointStyles。于是两份样式都指向缺失点，painter 按规则拒绝整图，实际输出是 `native drawing unavailable`。测试却期待 `4 missing; 0 zero`，因此失败。

这是**测试输入与其正例意图冲突**。当前 [PpjSemanticValidator.ValidatePointStyles](../native/OfficeKit/src/OfficeKit.Codec/PpjSemanticValidator.cs) 同样禁止缺失点有 visual override，不能为了让测试绿而移除这条保护。修复已将“合法全缺失、无点样式”的正例和“缺失点仍带点样式”的拒绝例分开，并保留两条断言。两种图表均按此拆分，没有放宽 validator；完整内部专项与 presentation 4/4 已恢复通过。

历史影响：当时合成专项和分段在同一位置退出 1，后续断言未执行；独立真实混合缺失用例通过并不矛盾。本轮重跑已经执行原来被挡住的 alpha、非法输入、边界和独立进程断言，均通过。历史失败用于解释修正原因，不再代表当前状态。

收口还需：更多现有 circular 字段、多环/负值与源样式/workbook 案例、分离布局的语义和边界验证及 G-01 4.x 生产诊断和发布集成。分离布局允许明确标记的静态 review 近似，不要求 Office 像素级间距；但必须保持比例、点归属与覆盖优先级，核对标签随动、裁切及适用拓扑，不能用“近似”掩盖数据错误。本轮已补全缺失回归和 registry/tasks 的当前证据。单系列比例完成不关闭这些差距。

#### 此前 circular 实施新增：系列与逐点 explosion 的径向预览

本小节的“本轮”指此前 circular 实施轮次；相关用例在后续柱形/条形实施中再次执行。本次文档整理不计为一次新的绘制或源编辑测试。

本轮不再将所有非零 explosion 一律拒绝：系列值与逐点覆盖值按存在性取值，逐点显式 0 可以取消本点分离，删除逐点值则恢复系列值，删除系列值回到缺省 0。未知/越界值和缺失点样式仍失败。

具体几何是一个公开标记的 **radial-review 布局**：

- 偏移沿扇区角平分线，距离为环带厚度乘以 explosion/100；pie 的内半径为 0。
- 根据有效正值扇区的最大偏移统一缩小半径，使圆及偏移仍落在现有 plot 范围；扇区比例与环孔比例保持不变。
- 缺失点没有扇区；真实 0 保留数据但不移动，也不因其 explosion 放大布局预算。
- SVG 记录实际 explosion、来源（point/series/default）、偏移和内外半径。非零绘制保留 `preview.scene.paint.chart-explosion-layout` 与页面 requires-review。

公式参考 LibreOffice 的非同心径向构造：其 [OOXML 导入转换](https://raw.githubusercontent.com/LibreOffice/core/master/oox/source/drawingml/chart/typegroupconverter.cxx) 将 explosion 转为百分比，其 [PieChart 绘制](https://raw.githubusercontent.com/LibreOffice/core/master/chart2/source/view/charttypes/PieChart.cxx) 按环带厚度及角平分线平移。**这是参考实现选择，不是 Office 精确显示契约**。本实现保留 OfficeKit 的 0～400 范围，不复制该转换器的 100 上限钳制；没有声称与 LibreOffice 全部布局相同。Office 精确间距、同心分离、多环和标签随动仍待验证/实现，不由这个有限映射关闭。

| 验证层 | 本轮实际通过 | 尚不能推出 |
| --- | --- | --- |
| 合成 SVG | 0/25/100/400、逐点零覆盖、旋转、真实零不移动、缺失保留、输入不变、偏移方向/长度与外框边界 | 不是 Office 像素级等价；超大 stroke、嵌套变换、完整标签布局未全验 |
| 真实 authored / source no-op | 两种图表各有 series=100、point=25 的创建与原字节 no-op；检查 native/SVG、实际颜色及空隙/透明孔像素 | literal-data 简单输入，不是第三方复杂主题 |
| 真实 source 编辑/删除 | 每类从原始投影分别构造 point-zero、point-delete、series-zero、series-delete、both-delete，合计 10 个候选；native、SVG、像素、再次投影均通过 | 不含 workbook、系列增删、跨类型编辑或任意 source-bound 生命周期 |
| 包保留 | 每次修改仅目标 ChartPart 改变；文件集合、其他 ZIP 成员、源字节不变 | 不推断原生宿主刷新缓存或动态行为 |
| 运行环境 | 测试独立 wire 比较改为调用时创建客户端并在完成时释放；扩展用例后不再复用已空闲回收的句柄 | 没有禁用生产空闲回收，也不吞掉真正 native 崩溃 |

最终产物及身份见第 1.3 节。原生集成仍因两个对象锚点失败而退出 1，不能把这组 circular 成功记成全链路通过。

### 2.7 内部普通分组柱形 / 条形：方向、缺失位置与源编辑

两种 PPJ chartType 都由 codec 生成 native `SpreadsheetChartType.BAR`。`barDirection` 的空值或 `column` 为竖向柱形，`bar` 为横向条形；native `xAxis` 始终持有分类轴，`yAxis` 持有数值轴，横向显示不交换这两个语义 owner。painter 直接读取该模型，不从原 PPJ 重算数据或猜类型。

| 行为 | 已实现及验证 | 尚未完成 |
| --- | --- | --- |
| 普通分组 | native grouping 空值/none；多个系列各有固定位置，缺失点留空而不挤占/移动其他系列 | 普通 stacked 和非负 percent-stacked 已单独实现，见第 2.8～2.9 节；不作为普通分组绘制 |
| 数据与零值 | 共用缺失索引校验；真实零保留点身份，以明确标记的虚线审阅刻线表示零面积；未知值不进入范围或生成柱子 | 含缺失的 zero/span 显示策略仍明确失败；不将拒绝算作策略已支持 |
| 正负值与范围 | 自动范围包含零；显式 min/max 保留；正负柱从真实零基线向不同方向延伸，超范围几何由 plot 裁切并保留原值标记 | 对数柱的非零基线未映射；真实 axis crosses、网格与完整自动刻度尚缺 |
| 方向与反向轴 | 分类轴默认 column 从左到右、bar 从下到上；分类/数值 reverse 都进入实际坐标，真实两类型的双反向几何和颜色像素通过 | 不据此宣称全部轴位置、标签和嵌套变换组合已验证 |
| gap / overlap | 原生存在性保留，gap=0 与删除不同；-100/0/50/100 overlap 和 0/100/500 gap 有合成几何断言；真实零间距、负重叠及删除通过 | 默认绘制间距 150、重叠 0 属本预览布局约定；完整 Office 自动布局未核验 |
| 直接样式 | 系列和点级填充/轮廓共用 helper，点样式优先；RGB、noFill、alpha 和显式零宽/零透明有断言，真实负值点覆盖颜色通过 | 主题/渐变/图片填充等继续给出字段限制；review 调色板不冒充原主题 |
| 源修改 | 两类型各 5 个候选从原始投影单独重建，检查 native→SVG→像素→再次投影及目标 ChartPart 修改范围 | 本项目 literal-data 文稿，无 workbook/第三方源；没有对象/系列增删或任意组合编辑验收 |

间距按“单柱宽度”计算：分类带宽为 `B`、系列数为 `N` 时，柱宽为 `B / (N - (N - 1) × overlap/100 + gapWidth/100)`，相邻系列起点差为柱宽乘以 `(1 - overlap/100)`。这与 Microsoft 对 [GapWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.gapwidth) 和 [Overlap](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.overlap) 的相对宽度定义相符；固定 plot 边距及自动范围不是宿主布局复刻。

真实输入第一系列为 `[4,null,0,-4]`，第二系列为 `[8,2,null,-2]`，分别保留不同缺失位置和实际零；第一系列负值点有绿色覆盖，其余系列颜色保持。每种方向独立执行数值 `4→6`、gap `100→0`、overlap `0→-100`、删除 gap、删除 overlap。10 个候选都仅改变 `ppt/slides/charts/chart1.xml`，其他 ZIP 成员及原源字节不变；删除后的重新投影确认字段不存在，不能以默认值代替删除结果。新创建、双反向轴、source no-op 另有正例，不与源修改数量混算。

正负混合图额外显示虚线零参考线及 0 刻度，明确标记为 review guide；它不是作者指定的 gridline/axis crossing。纯零点也只有审阅刻线，不画伪造高度/长度的柱。所有内部图保留 `chart-layout` 和每页 `integration-pending`；标签、图例、效果、负值百分比、对数、混合轴及生产接入未完成，G-01 3.4 与 G-08～G-10 均不关闭。

### 2.8 内部普通堆叠：累计坐标和缺失分类

本节描述普通 `stacked`；非负 `percent-stacked` 的后续实现见第 2.9 节。line/area 的堆叠仍未实现。painter 读取已编译的 native 分类/系列/缺失索引，不计算 PPJ dataset 或解析 OOXML。

- 同一分类内，正值从前面正值之和开始，负值从前面负值之和开始；正负不互相抵消。自动范围取累计端点，显式 min/max 仍保留并裁切。
- SVG 为已定位点保留原值、`data-officekit-baseline` 和 `data-officekit-stack-end`。真实零可位于非零累计起点，仅显示审阅刻线，不制造面积。
- 任一系列在该分类缺失时，当前整个分类不画累计柱；显示 Incomplete，缺失与已知值分别保留，已知值标 `stack-position="unknown"`。自动范围只来自可定位的分类；不宣称覆盖未知总量。完整缺失显示策略仍待补，这种保守显示不等于全部缺失语义已支持。
- 累加溢出、非零项因精度丢失而使终点等于起点时，明确失败。不能先吞掉数值再输出一张看似完整的图。
- `overlap` 仍控制横向系列位置：100 时同分类对齐，0/-100 保留声明的错位或间距，不因 grouping 是 stacked 就改写显式值。缺省沿用内部 review 的 0，不声称还原宿主自动默认。其独立属性定义见 [Open XML Overlap](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.overlap?view=openxml-3.0.1)；本轮真实堆叠输入显式指定 100。

真实创建输入为第一系列 `[4,null,0,-4]`、第二系列 `[8,2,0,-2]`。正分类累积到 12，负分类到 -6，零分类保持两个真实零，未知分类不造出“2 就是总量”的柱。每方向从原始投影修改首值为 6，正分类累计到 14；候选绘制、像素、重新投影与 ZIP 修改范围都通过。没有 workbook，不能由这两例推断第三方或公式数据联动正确。

仍缺负值百分比策略、完整缺失显示策略、系列/对象增删、完整标签与主题，以及生产检查/发布接入。新增绘制与源编辑证据不关闭 G-01/G-08～G-10。

### 2.9 内部非负百分比堆叠：原值与比例分开保留

非负完整分类按各项占分类总量的份额绘制，符合 [Office 图表类型说明](https://support.microsoft.com/en-US/Excel/available-chart-types-in-office) 中的部分与整体关系。此处不据此推导负值策略或宿主自动布局。

| 情形 | 实际行为与证据 | 保留限制 |
| --- | --- | --- |
| 总量为正且完整 | 用 `(value / max) / sum(value / max)` 计算份额，避免有限原值之和溢出；逐系列累计，默认轴为 0～1 并显示百分比刻度 | 固定刻度/边距不是完整宿主布局；原生明确 min/max 仍优先 |
| 原值与显示值 | `data-officekit-value` 保留原值，`fraction`、baseline、stack-end 记录份额及累计坐标；修改原值会真正改变图形比例 | 没有把原始 PPTX 数据改写成百分数，也未把坐标标签当作数据值 |
| 缺失分类 | 显示 Incomplete；已知值和未知项各自保留，无推断百分比 | 不能用“已知项合计 100%”冒充包含未知项的总体 |
| 全零总量 | 显示 Zero total，真实零保留且标 `stack-position="zero-total"`；不生成 fraction、柱面积或假 100% | 百分比数学上未定义，不能声称零值丢失或成功算出了比例 |
| 正总量中的真实零 | fraction 为 0，在真实累计位置画明确的审阅刻线 | 刻线不代表非零面积，与全零总量不同 |
| 极端数值 | 同分类 `1e308 + 1e308` 可得到两个 50%；下溢或累计精度吞掉非零项时失败 | 不承诺任意精度或所有动态范围 |
| 负值 | 当前明确失败，不自行变成绝对值 | 负值百分比分母与符号显示仍待单独核验，整类百分比堆叠未完成 |

本轮读到的 [LibreOffice BarChart 实现](https://raw.githubusercontent.com/LibreOffice/core/master/chart2/source/view/charttypes/BarChart.cxx) 在 percent 分支对高度取绝对值后归一化。该参考不能单独证明 OfficeKit 应改变负号，因此没有照搬。后续应确定与现有输入/导出一致的有符号语义并加入真实反例，而不是以当前负值拒绝作为永久收口。

真实新增验证包含两方向、各自创建/no-op/源数值修改、像素与重新投影；两种普通堆叠回归同时保留。`nativeStacks` 报告按 `type + grouping` 区分四组，百分比两组为新增；零总量标记只出现在百分比组，普通堆叠不适用。第三方源、workbook、负值百分比、完整显示策略、图例标签和生产接入仍开放。

## 3. 元素与页面视觉差距

下表对齐 [PPJ schema](../src/ppj/ppj-v1.schema.json) 中的 16 类元素，描述**当前生产 CLI 路线**，不是内部 painter 的映射表。声明状态来自当前 [preview capabilities](../src/ppj/svg-preview-capabilities.json)，描述整个类型的保守边界；运行时还会检查实际字段与继承状态。内部有限修复见第 2.3～2.7 节，尚未切换生产入口。当前没有任何整个元素或图表类型被声明为 supported；这不表示连一个简单原语也画不出来。

| 元素 | 声明状态 | 实际行为和未覆盖范围 |
| --- | --- | --- |
| `text` | partial | 输出纯文本和固定行距；未完整处理 runs、字体、字号、段落、换行、AutoFit、列表及样式继承 |
| `shape` | partial | 无文字时基本画矩形；有文字时提前进入文字分支，几何和填充丢失 |
| `line` | opaque | 无独立绘制分支，通常进入未知元素占位；路径和自由曲线未实现 |
| `icon` | opaque | 无图标轮廓绘制分支 |
| `image` | partial | 使用固定 contain 式 `<image>`；裁切、焦点、mask、边框和 effects 未完整映射 |
| `chart` | partial；16 个 chartType 也均为 partial | 各类型差距见第 4 节，不能整体认定支持 |
| `table` | partial | 均分行列；未按显式尺寸、合并单元格和样式绘制 |
| `connector` | partial | 固定水平线，不解析 from/to、anchor、路线和箭头 |
| `group` | partial | 递归输出子元素，没有 childFrame 与父 frame 的坐标转换 |
| `media` | opaque | 无专门的静态海报/封面显示或播放能力说明 |
| `placeholder` | partial | 实现为虚线框加文字；原先同时属于两个等级的声明已修正 |
| `smartArt` | opaque | 没有消费 SmartArt 实际布局，落入通用兜底 |
| `ole` | opaque | 已纠正旧声明 `embeddedOle`；没有使用 OLE `previewAsset` 的专门分支 |
| `opaque` | opaque | 有 `previewAsset` 时可以显示图片并报告 partial，否则通用虚线框；缺完整源预览和诊断回归 |
| `component` | partial | 仅抽取 label/value 并竖排，不按组件定义、variant 和真实布局展开 |
| `slot` | opaque | 没有独立解析/绘制分支，须与模板或组件展开共同处理 |

`nativeRef` 是源绑定字段，不是与 `shape` 平级的实际元素类型，当前已单独列入 sourceBound 声明。即使某个类型有通用占位或字段限制诊断，也不代表该类型已经正确绘制。

### 3.1 G-02：文字和形状内容缺失

代码中的 `if (e.type === "text" || e.text)` 位于 shape 分支之前。任何带真值 `text` 的形状都会只输出文字。

初次审计的 canonical fixture 决策节点 `decision-flow-gate` 本来是 `flowChartDecision`，有填充、白色文字和居中设置；当时实际 SVG 摘录为：

```xml
<g data-officekit-id="decision-flow-gate"><text x="806" y="302" font-family="Arial, sans-serif" font-size="18" fill="#172033"><tspan x="806" dy="0">Pass?</tspan></text></g>
```

该组没有菱形或背景。初次审计时此节点也没有出现在 diagnostics 中；当前 G-11 已为此类几何丢失报告 `preview.fact.shape-geometry-omitted` 并使可靠性失败，但尚未补画几何。

其他代码确认的差距：

- `textValue` 把每个 run 用换行连接；同一段内“普通字 + 加粗字”会被错误拆行。
- 字号读取 `textStyle.fontSize`，没有完整消费实际 `defaultText.size`、run style 和命名样式。
- 文字首行偏移固定为 18、后续行偏移固定为 22；没有完整排版、边距、对齐和溢出处理。
- 所有普通 shape 都输出 `<rect>`，忽略 `geometry.preset` 和 preset adjustments。
- 旧绘制分支的自定义路径诊断检测 `geometry.customPaths/path`，而 schema 使用 `kind: custom`、`viewBox`、`paths`，因此该分支不能正确识别标准自定义路径。当前 G-11 的共享字段检查另以 `preview.geometry.unassessed` 等报告几何限制；旧分支漏报不再等于最终完全无诊断，但真实路径仍未绘制。

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

资源加载的以下两个问题在初次审计中存在，现已由 G-12 修复：

1. workspace 已读取资产字节，预览却又按 URI 读一次磁盘，未复用同一份已加载数据；编译和绘图之间存在资源变化窗口。
2. 第二次读取失败会被吞掉并生成空 `href`。Map 中仍存在 asset 对象，因此普通 image 的“资产不存在”检查不一定能发现这个错误。

当前实现直接使用 `workspace.assets[].data` 构造图片，普通 image 对空 href 报 unavailable，发布层将其计为失败；相应故障测试已经通过。以上历史问题不表示所有缺失资源都能绕过 workspace，很多错误会更早失败。资源身份修复不代表 crop、mask、主体判断或所有图片解码情形已经验收。

完成条件：消费同一份验证过的资产；裁切与 mask 使用已知参数，不猜主体；无法确定的信息保守显示并说明；空字节/解码失败不能伪装为成功。

### 3.5 G-06：表格与连接关系

表格直接用元素宽高除以行列数量。canonical fixture 的列宽为 116/158，但预览画成两列各 137；行高和单元格合并、填充、边框、文字排版也没有完整应用。

connector 直接画 frame 中线，没有读取 `from/to.element`、anchor、connectorType 和箭头。初次审计的 canonical fixture 连接线输出只有水平 `<line>`，没有 `endArrow: triangle` 对应的箭头；当前该绘制分支未改变。

#### G-06 对象锚点：源码修复和当前运行时闭环已验证

当前 `066a6924` 已增加独立求解器：显式 frame anchor 按目标及祖先组变换求坐标；auto 在 top/right/bottom/left 候选中按页面距离选取，平局保留固定候选顺序。结果回到连接线父坐标系，再量化为 EMU。目标缺失、不支持的目标类型、不可逆 frame 或超出非负本地坐标范围会明确拒绝，不再退回连接线自身 frame。后续 `94e3006f`、`ddc4213f`、`474e9c54` 又补齐了有符号端点、箭头和 source-bound 类型编辑。

绑定通过有界 frame-anchor 元数据保留，不冒充几何 connection-site 索引；native 已另行保留部分 connection-site 信息。源编辑包括端点更改、重定向、解除/建立绑定及目标移动依赖重算；畸形元数据保守处理。源码专项记录见第 1.3 节及独立变更任务清单。它没有解决自动避障、曲线/肘线精确路径、原生任意连接站点的通用场景绘制/编辑或宿主行为验收；这些不能从 frame-anchor 修复推导出来。

**下文调用路径和数值反例记录的是修复前源码与旧临时运行时，保留用于新构建回归，不再代表当前 HEAD 实现。**

这不是只改 JS painter 就能解决的问题。当前 [schema 的 connectorEndpoint](../src/ppj/ppj-v1.schema.json) 明确允许两种互斥输入：`{x,y}` 或 `{element,anchor}`；后者的 anchor 可以是 auto/top/right/bottom/left/center。不能要求对象锚点输入再补一组 x/y 来掩盖失败。

修复前调用路径：

1. [PpjProgramModels](../native/OfficeKit/src/OfficeKit.Codec/PpjProgramModels.cs) 将 element、anchor、x、y 解析并保留在 `PpjConnectorEndpointModel`。
2. [PpjSemanticValidator](../native/OfficeKit/src/OfficeKit.Codec/PpjSemanticValidator.cs) 检查目标 ID 是否存在于同页或组件作用域；这不是锚点坐标求解。
3. `PpjAuthoredPresentationCompiler.BuildConnector` 用 `EndpointX/EndpointY` 写 native 坐标。X 无显式值时取连接线 frame 的左/右边，Y 取 frame 中线；不读取 `ElementId/Anchor`，该分支也没有写入目标 ID/site 绑定。
4. 场景忠实采集这个 writer 输入；内部 painter 再忠实绘制坐标，并不能修正已经丢失的作者关系。SmartArt 生成连接线另有构造路径，不能用它的目标绑定证明普通 authored connector 正确。

本轮复现输入来自现有集成测试：

| 对象 | 输入或预期 |
| --- | --- |
| 起点框 | frame `(750,300,60,40)`，取 left，预期端点 `(750,320)` |
| 终点框 | frame `(450,100,60,40)`，取 right，预期端点 `(510,120)` |
| 连接线 | 自身 frame `(0,0,1,1)`；from/to 使用上述对象和锚点，straight，endArrow triangle |
| 实际 scene | `(0,0.5) → (1,0.5)`，与连接线 frame 回退逻辑完全一致 |
| 目标移动用例 | 将终点框 y 增加 40，预期终点 `(510,160)`；旧包实际为 `(1,0.5)`，第二例同样失败；当前包重新编译后两例均按预期通过 |

这条合法输入已进入旧运行时实际编译；失败是作者关系与编译场景不一致的历史功能反例。当前 C# 源码已不再采用上述回退逻辑；最后一次已验证的 `495daafc` 快照包 `/home/zenfun/mywork/OfficeKit/tmp/officekit-preview-runtime-495daafc` 的真实集成报告已记录两例端点、箭头和 bend 字段通过，但这仍不是 PowerPoint 宿主行为验收；之后的 HEAD 尚未用新包重验。

收口要求分两层：先在负责布局和对象关系的编译层修正解析，或对尚不能安全求解的拓扑明确拒绝；再验证同一端点、方向、箭头进入 SVG。明确锚点、auto 策略、目标移动、嵌套组、组件展开和非法目标各自需要适用案例，不能由 JS 自建第二套 PPJ 布局求解绕过它。显式 x/y 和 source-bound 原生连接线也要保留回归，防止修正普通 authored 路径时破坏已有行为。

范围决策已经落实为独立的 [connector 编译修复变更](../openspec/changes/ppj-connector-object-anchors/proposal.md)，并由箭头/类型编辑提交继续扩展；connection-site 字段也已进入 native 合同，但尚未由当前包验证。[G-01 设计](../openspec/changes/ppj-preview-compiler-scene/design.md) 的只读场景采集约束继续保留：修复前后输出发生的必要变化属于 connector 修复；同一版本开启/关闭场景采集仍须保持输出不变。当前包已复跑 native → SVG/PNG；剩余是生产 painter 接入、任意 connection-site 的通用绘制/编辑、自动避障、复杂路线及宿主行为边界。

内部表格/连接线的有限映射见第 2.4 节。旧包脚本曾在收集两个关系失败后继续执行独立表格/source-bound 断言；最后一次已验证的 `495daafc` 快照报告 `/tmp/officekit-native-scene-paint-MTxwlx/integration.json` 的 `relationFailures` 为空，独立表格、source-bound、line、图表和弧线路径均执行通过。该结果仍仅覆盖内部 scene painter，不代表生产 CLI 已切换，且不覆盖之后 HEAD 的新字段。

完成条件：表格几何来自真实尺寸与 span；连接线随实际端点和锚点变化，不能用固定方向代替关系；独立 `line` 和 `connector` 应分清职责。

### 3.6 G-07：opaque、源绑定和非静态能力

有 `opaque.previewAsset` 的源图像显示路径已经存在，不应误记为“完全没有 opaque 支持”。G-12 也已验证真实 source-bound 发布与源字节保留，但所用输入是本项目生成后去除私有快照的简单文稿，不是第三方复杂 PPTX。以下缺口指复杂输入和编辑后内容正确性的完整回归：

- 无预览时，把 summary、nativeKind、visibleText 和未支持原因清楚呈现给 reviewer；目前通常只有无说明的虚线框。
- 对 OLE 等其他实际类型的源预览处理。
- 对 imported/nativeRef 修改后，预览显示的内容是否确实对应修改后状态的验证。
- 第三方输入的原包保留、修改范围、重新导入与稳定 ID 证据。
- 对动画、媒体、交互的静态检查范围说明，避免一张图被误当作播放/行为验收。

预览使用一张源快照，不等于把可编辑对象写成图片；两者应分开描述。不能为生成预览而破坏原包或把 opaque 拓扑重写成猜测结构。

## 4. 图表差距

### 4.1 G-08：数据归一化和公共坐标系统缺失

本节及第 4.2、4.3 节的缺陷描述针对生产 `svg-preview.mjs`。内部 line、pie/doughnut 与 column/bar 已有第 2.5～2.7 节的局部修复，但尚未替换生产路径；不能把两套实现的支持范围合并。

预览直接读 `data.categories` 和 `data.series[].values`，未完整处理 `dataset/encoding/dataFilter/seriesDefaults`。坐标范围基本统一使用 `Math.max(1, ...values)`，没有实现负数范围、明确 min/max、双轴、数值 X 轴、轴反向及其他缩放方式。

以下字段也没有完整应用：title、legend、轴标题与刻度、网格、numberFormat、dataLabels、marker、pointStyles、trendlines、errorBars、chart frame 和 plotArea 样式。

这些字段既可能影响外观，也可能影响事实理解。例如把不同单位的主副轴压到同一尺度，会改变读者对两条序列变化的判断。

完成条件：绘图消费归一化数据；每个 mark 可追溯到 series/point 和正确坐标轴；数据表、标签、图例与图形一致；纯样式差异和数据关系错误分开报告。

### 4.2 G-09：逐类型差距

下表覆盖当前 schema 的 16 种 `chartType`，另列两种通过字段表达的变体。具体实现见生产路线的 [chart 分支](../src/ppj/svg-preview.mjs)。内部 painter 已验证一组真实 Sunburst/Sankey 生成路径、有限原生 line、单系列 pie/doughnut 和普通分组 column/bar，但没有完成全部类型、样式、边界或生产绘制验收，不能据此关闭 G-09。

本表为代码与 schema 对照结果，不表示本次对每一种类型都执行了正负案例；最后一列的验证案例是待补测试。本轮实际运行范围见第 1.3 节，历史 smoke 结果见第 6.1 节。

| 类型/变体 | 当前绘制 | 关键差距与最小验证案例 |
| --- | --- | --- |
| bar | 与 column 共用竖矩形逻辑 | 没有横向条形布局；仅设置元素 chartType 而省略 series.chartType 时可能没有数据 mark；验证普通横向两系列 |
| column | 固定宽度竖矩形 | 各系列重叠在相同位置；未正确处理分组/堆叠/负值，null 被转 0；验证两系列正负值与缺失点 |
| line | 按有限数值拆 polyline 段 | series 类型继承不完整；单点段被丢弃、marker 未画；验证 `[1, null, 3]` 保留两个独立观测 |
| area | 与 line 相同的无填充 polyline | 缺面积、基线和正确堆叠；验证多系列面积含缺失点 |
| combo | 拼接 column/line/area 分支 | 共用最大值，忽略主副轴、单位与部分样式；验证柱线双轴不同量纲 |
| scatter | 按数组序号或假定的对象 x/y 画点 | schema 的 values 是 number/null，真实 X 在 `xValues`；当前忽略该通道；null 在访问 `v.y` 时确实抛错；现有回归已确认并转为 unavailable 占位，散点缺失语义仍待修复 |
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

### 5.1 G-11：支持等级与可靠性检查已接入

初次审计的类型声明、分支自报和总状态聚合相互矛盾：存在绘制分支就称整个类型 supported，部分图表分支自行报 supported，而总状态又仅由诊断数量决定。当前这些输出判断已统一到实际输入检查；原分支信息仍保留 reason，但不能提升检查层给出的等级。

#### 已实施行为与剩余边界

| 层 | 当前实现与证据 | 不代表什么 |
| --- | --- | --- |
| registry 与生成摘要 | 校验真实元素/chartType、schema 指针、owner、测试路径、等级互斥；摘要与矩阵均可只读检查漂移 | 不是每个视觉字段都已画对或有完整正负 fixture |
| 共享诊断 | pageId/id/path/status/reason/severity/valueSummary/action；值摘要转义限长；保留 0/false/null；确定性合并 | 不能由无诊断推断未评估状态通过 |
| 实际输入检查 | 已在编译后、绘制前调用；遍历页面、组和组件，继承全局/页面/组限制；未知后代不默认支持 | 不解析布局、token 或 styles，不代替 G-01/G-04 |
| source/binary 边界 | 对 source/nativeRef 等源字段停止展开，敏感内容省略，输入保持不变 | 不证明复杂第三方 PPTX 的修改后快照有效 |
| 字段限制 | 文本、几何、变换、主题、图片、表格、连接线、图表、effects 和未解析范围均有实际路径诊断，包含 errorBars | 诊断不是对应功能的绘制实现 |
| 事实错误 | 14 类 registry 规则；每类经过实际 SVG 绘制、样式变更和成功发布回归，错误仍为 failed | 未登记或未解析的语义仍可能只报 requires-review，不能宣称发现了所有错误 |
| 页面警示 | failed 显示红色警示，partial/opaque 显示复核警示；SVG 与 PNG 同源；原 ID、画布和输入不变 | 顶部警示是审查覆盖层，不是修好原页面，也不能预知随后磁盘失败 |
| 绘制异常 | 已知 scatter/null 异常转为可定位的 unavailable 占位，保留其他页；发布清单记录 preview 失败阶段 | 不把绘制异常误报成缺依赖或资产缺失；scatter 本身仍待修复 |
| 发布证据 | 返回与落盘检查树、页/全局状态、可靠性一致；生产失败保留 G-12 错误码、产物、hash 并使可靠性失败 | output complete 不等于视觉正确或编辑保真 |

聚合顺序为 `unavailable > opaque > partial > supported`；可靠性独立分为 `failed / requires-review / passed`。已知事实错误使可靠性失败，普通未知限制要求复核。未解析轴 token 和与当前范围相同的明确 min/max 不被误当作已知数值矛盾，但仍有字段限制。实际变换边界、文字溢出、阴影范围仍未计算；检查器明确报告未解析范围。

一个已验证的有限支持状态是显式 contain PNG：检查 SVG 坐标、透明度和内嵌字节，增加 crop 后不再 supported。此正例注入 load/compile，不能推广为所有 PNG 解码、主体或透明边缘的端到端验收。真实 codec+sharp 专项另验证 canonical 和简单 source-bound 输入、原始字节、产物 hash、尺寸与警示像素。

#### 任务与验收

G-11 实施轮次记录本变更 9/9 项已完成，presentation 四步、gate-policy、生成摘要/矩阵、portability、reference-sync、链接和严格 OpenSpec 检查通过；同时真实 CLI 回归确认退出码 0 可与 failed 可靠性并存。最新复核重新运行了 scene SVG/view、能力覆盖、gate-policy、两个生成器、`proto:check`、strict OpenSpec 和 `slow/presentation` 4/4；这些只是门禁/内部回归，不等于完整全仓测试、宿主验收或性能完成。

- 1.1、1.2、2.1、2.2：同源声明、诊断基础、实际字段遍历和字段限制。
- 2.3：每个已登记事实错误在绘制与发布后仍为失败，配色不能抵消；未知限制与已知矛盾分开。
- 3.1、3.2：真实绘制/警示/发布接通；嵌套重复局部 ID、失败保留、依赖惰性和源字节保留有断言。
- 4.1、4.2：常规门禁、相关文档/Skill reference 与严格 OpenSpec 检查；最终勾选以 [任务清单](../openspec/changes/ppj-preview-support-diagnostics/tasks.md) 为准。

真实 authored 样例成功发布，但为 partial/failed；简单 source-bound 样例成功发布，但为 opaque/requires-review。清单 `ok: true` 和 CLI 退出码 0 仍只说明发布成功，调用者必须读取可靠性。failed 不得作为正确性证据；requires-review/passed 也不取代人工视觉、结构和编辑保真检查。CLI 的整体交付策略仍属 G-13。

### 5.2 G-12：输出安全和失败处理

修复进度：下面列出的初始输出问题已有实现与专项回归。新增 `src/ppj/preview-output.mjs` 独占创建目录和文件，在失败时保留实际产物；清单记录 hash、字节数、原始页 ID、输入/源/候选/资产身份以及 Node/栅格后端信息。源码版本摘要也进入清单。SVG 和 PNG 仅在写入成功后列出，缺少 sharp 返回不完整失败而保留 SVG；pending/final 清单与错误码区分失败阶段。

前次修复验收中，`npm run test:ppj-preview-output`、`node test/ppj-svg-preview.mjs`、`node test/ppj-preview-capability-coverage.mjs` 均已通过；最近 circular 实施的完整 presentation 分段包含上述三个脚本，4/4 通过；不与单项结果重复计数。真实 SVG/PNG 专项同时验证 authored 和原生重新投影后的 source-bound 输入，原始 PPJ/PPTX 字节保持不变；模拟故障测试覆盖目标冲突、并发、路径映射、资源变化窗口、页面栅格失败及清单写入失败。

集成修复进度：新轻量回归已加入 fast/slow gate，slow 的 presentation 分段同时执行诊断、输出发布、SVG 预览与覆盖声明检查。`node test/gate-policy.mjs` 校验当前 PPJ 入口、已退役入口不再出现及分段连续性。此前的合成失败已由合法/非法 fixture 拆分处理；最新复核中该分段 4/4 通过。原始缺陷记录如下：

- `mkdir(..., recursive: true)` 加普通 `writeFile` 会复用目录并覆盖同名 SVG/PNG/render.json，没有继承外部 render 的独占输出保护。
- 写入是逐文件进行，栅格失败可能留下部分输出；缺失败清单、完成标记或原子交付策略。
- 缺少 sharp 时，render.json 仍为每页列出 PNG 文件名，实际文件可能不存在。
- 已记录 program/output hash，但缺每张预览图的实际 hash、sourceBound、输入/输出定位、渲染版本和依赖信息等完整关联证据。

本项完成条件已在限定范围内满足：不覆盖用户输入和已有证据；失败结果区分已产生/未产生的文件；输出清单只声明真实产物；预览绑定实际输入和资产。路径负例已包含现有目录/文件/符号链接、危险页 ID、大小写冲突和并发竞争。测试不等于防御恶意进程替换整个目录树，也不提供断电后的 fsync 持久性保证。

仍需区分四种状态：`output.status: complete` 只说明请求产物已发布；`status/diagnostics` 是保守聚合的支持情况；`reliability` 是独立的自动可靠性门槛；`visualReview: requires-human` 说明人工视觉检查尚未完成。它们不能互相替代。

### 5.3 G-13：CLI 和交付流程未完整接入

`preview` 复用了 render 参数解析，接受 `--pages`，但 handler 没有把页选择传入实现，实际仍渲染所有页。preview 返回值也没有专用 command/摘要分支，非 `--json` 时会回落成完整 JSON，包含 SVG 内容。

自有 preview 未自动进入 `check → build → render/review → imported re-import` 的交付证据链。已有 review 能力仍然有用，但不能把 preview 的顶层边界检查当作完整 review，也不能把外部 render 的证据标签套在 SVG preview 上。

完成条件：参数要么实现、要么明确拒绝；统一可用的 CLI 返回和错误码；在 Skill/文档中说明两条渲染路线及其证据边界；结构、视觉、编辑保真检查有独立状态。

## 6. 测试、覆盖台账和性能差距

### 6.1 G-11 实施测试与历史 smoke 结果

下表以 G-11 实施轮次为主，最后一次已验证快照的检查见第 1.3 节，G-01 基础实施证据见第 2.3 节。42/4/38/0 与下方四个诊断数是该快照的本机旧生产路线结果；当前 HEAD 尚未用新包重跑，因此不能当作当前代码结果，更不是视觉正确性证明。

| 检查 | 实际结果 | 证明范围及限制 |
| --- | --- | --- |
| `node test/gate-policy.mjs` | 通过 | 活跃入口、已退役入口、分段边界及脚本存在性符合当前门禁契约 |
| `npm run test:slow -- --segment presentation` | 4/4 通过 | 顺序执行诊断、发布故障、真实 SVG/PNG、能力声明检查；不等于全部 slow 通过 |
| 共享诊断测试 | 通过 | 聚合、字段遍历、PNG/SVG 映射、事实反例及实际绘制/发布；样式和输出成功不清除事实失败 |
| 发布故障测试 | 通过 | 保护源输入和已有目标、并发、路径、部分写入、hash、资产快照、依赖惰性加载；模拟 PNG 不用于评价图像 |
| 真实 SVG/PNG 测试 | 通过 | canonical 两页、真实 source-bound 投影、源字节、候选/产物 hash、警示像素、返回/落盘检查一致；仍非逐元素绘制正确性验收 |
| 能力声明测试 | 通过，16 类元素、16 类图表 | schema 对齐、等级互斥、嵌套扫描、非法规则和摘要漂移检查；不是字段保真或视觉质量证明 |
| 矩阵生成器 `--check` | 通过 | registry 派生内容与当前生成矩阵一致，未重写矩阵 |
| Skill/文档检查 | portability 255 文件、reference-sync 333 文件、strict OpenSpec 通过 | portability 原来误要求 typed PowerPoint Live 使用 REPL，已按实际 typed CLI 纠正并保留安全断言 |
| JS 预检 | 292 文件通过 | 语法/import 检查，不是全部功能回归 |
| 全量 preview smoke（最后一次已验证快照） | 退出码 0；42 输入，4 rendered、38 compilerRejected、0 rendererFailed | 在仓库 `tmp/` 指定临时目录运行；真正进入绘制的四项均为 partial，诊断数见下文；当前 HEAD 尚未重跑，不代表视觉通过 |

当前全量 smoke 中，minimum/canonical/aqua-impact-story/simple-dark-mode 的诊断数分别为 44/686/3674/1032，四项支持状态均为 partial。这表明检查已经进入实际绘制，也提示继承诊断量需要在 G-16 测量和优化；数量多不等于语义覆盖完整。以下初次运行表仍保留为历史记录，不能与当前快照混用。

初次审计结果（保留原测试能力描述；真实专项测试后来增加了源绑定与输出证据断言）：

| 检查 | 结果 | 真正证明的范围 |
| --- | --- | --- |
| `node test/ppj-svg-preview.mjs` | 通过 | 单个 canonical fixture 能生成两页、稳定 ID 字符串存在、PNG 大于 1000 bytes、manifest 页数正确 |
| `node test/ppj-preview-capability-coverage.mjs` | 通过，30 个声明项 | 有部分类型名称声明；不代表 30 种功能绘制正确 |
| `node test/ppj-preview-smoke.mjs` | 42 输入，4 rendered，38 compilerRejected，0 rendererFailed | 4 个调用完成预览；失败数量按脚本错误文本分类，未独立计量各阶段；没有进行语义/视觉等价验收 |
| canonical 输出检查 | 确认组件未展开、决策形状丢失、箭头未画 | 这些缺陷在现有测试通过的同时真实存在 |
| 能力矩阵比对 | 初次比对不一致；写作期间再次读取已一致 | 确实观察到过漂移，随后被其他工作流更新；不再把它记成当前仍不一致 |

初次审计：4 个生成预览的输入及当时的程序返回状态（不是当前声明）：

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

上表和第 3.1 节 SVG 是此前工具输出的审计摘录，不是本轮重新测量值。完整原始 stdout、临时 PNG/SVG/render.json 没有归档为仓库固定证据；本文不声称这些精确运行结果已具备跨环境复现包。后续正式验收应保存原始记录，而不仅保留汇总表。

G-11 实施轮次的通用 Skill 模板校验器另报告已有 `name: "Presentations"` 不符合其小写命名规则；该轮只改交付 reference，未改名或变更路由。该轮仓库专用 portability/reference-sync 已通过，此工具兼容性差异不冒充通用校验成功。最近 circular 实施另重跑了仓库 portability/reference-sync；最新文档复核未重跑这三类 Skill 检查。

### 6.2 G-14：测试覆盖不足

当前已有有限图片映射和内部几何/表格/折线/圆形图的直接断言，正式路线还会识别已登记事实错误并核对警示像素，但远未覆盖全部绘制。正式 CLI 的饼图仍只有圆、带文字形状仍丢几何、双轴仍不正确；内部 pie/doughnut 局部比例已修复。两条路线不能混算覆盖，图片存在性或发布成功也不能证明画对。

覆盖测试的变化与剩余问题：

- 初次图表名称靠硬编码和字符串匹配；现在按实际 schema 对齐 16 类元素和 16 类图表，并校验 registry 指针和 owner。全量视觉字段到具体行为断言的覆盖仍未完成。
- 初次只检查等级合并后的 Set；现在校验等级互斥、非法提升为 supported 和生成摘要漂移。规则中测试文件存在不等于该测试证明了整个类型正确。
- 初次只扫顶层；现在递归扫描 fixture 对象树，并有嵌套 group/未知类型反例。独立输入检查器也测试 component 定义、slot、重复局部 ID；内部已另有一组真实组件/显式配对及路径断言，但完整布局展开与逐对象显示覆盖尚未验证，不能由扫描测试推断完成。
- 当前 `test/fixtures` 下只有一个 `.ppj` 文件；这不代表仓库没有其他测试生成 PPJ，只是该扫描并未检查它们。
- 初次审计时专项脚本未进常规 gate。当前已修复：发布故障测试进入 fast/slow，真实 SVG/PNG 与声明覆盖进入 slow/presentation；全量仓库扫描 smoke 仍是独立命令，未设正向渲染数量门槛。

smoke 还通过错误字符串正则区分 compilerRejected/rendererFailed，而非实际失败阶段；且仅当 rendererFailed 非空才失败。即使所有输入都在编译阶段被拒绝，也可能返回成功退出码。

另一个待补回归来自输出保护的新边界：smoke 用成功数量 `report.rendered.length` 命名下一例目录。如果某例在取得目录后发布失败并保留该目录，下一例仍使用同一编号，会遇到 `preview.output.exists`，干扰后续故障分类。此为当前代码可以推导的条件路径，本轮未注入该故障；后续应按输入序号或唯一案例 ID 分配目录，并验证单例失败不污染后续案例。

完成条件：结构、数据语义、视觉、编辑保真分别有断言；正例必须确实进入渲染；预期拒绝独立列负例，不得用它们抵消正例失败。事实错误设硬门槛，不能被外观评分抵消。

### 6.3 G-15：能力台账尚未防止后续漂移

此前发现内部 circular 实现与 registry/tasks 描述脱节，之后已补记比例、缺失/真实零、逐点绘制、径向分离、源编辑/删除和剩余限制。最近柱形/条形实施也已同步普通分组及 10 个源修改候选的证据；本次文档复核确认任务勾选与生产支持等级不变。两个生成器 `--check` 继续通过；这仍只证明生成内容与 registry 一致，不能证明所有字段都有绘制与行为断言。完整字段维护闭环仍待实现。

G-11 实施复核时读取的生成矩阵记录了 183 个 schema definitions、46 个 authored boundaries、278 个 native leaf kinds、151 个 Help API 等计数。本次只重新执行生成器的漂移检查，不以这组历史计数作为当前全功能盘点。这些是不同粒度的集合，不是可相加的“渲染功能总数”，也会随并行接口开发变化。

矩阵生成器的 `visualStatus` 根据编译 boundary 的 behavior 推导为 unreviewed 或 partial-or-opaque，并不是渲染验证结果。它还将 programJson 标为复用入口，却没有指出 canonical JSON 与已展开布局的区别。

G-11 已修复类型声明矛盾，并以 registry 同源生成摘要和矩阵；两个生成器均支持 `--check`，覆盖测试会拒绝新增类型未登记和摘要漂移。未知视觉后代在实际预览中不会静默认定 supported。剩余缺口是：没有证明每个新增视觉字段都具备对应绘制、正负 fixture 和行为断言。类型/摘要防漂移已经有了，字段级完整维护闭环还没有。

完成条件：以 schema/能力 registry 为来源，建立“字段 → 语义 owner → 预览映射或明确边界 → 正/负 fixture → 断言 → 文档”对应关系。新增视觉字段必须触发维护检查；生成矩阵有 `--check` 或等效漂移 gate。不能仅要求开发者记住同步维护几份清单。

### 6.4 G-16：高性能和运行稳定性尚无验收

当前能确认 SVG→PNG 使用 sharp，不能据此宣称整个渲染器高性能。preview 每次先进行 PPTX 编译，再将已加载的资产编码为 data URI、拼接整页 SVG，并在结果中保留所有页面 SVG；写 PNG 逐页进行。G-12 已消除编译后的资产路径二次读取，但并未消除 base64 和多页 SVG 的内存开销，大量图片也可能在多个页面输出中重复。

尚缺：冷启动/热运行耗时、编译与绘制分阶段耗时、峰值内存、字体环境、长文和大图、复杂图表、多页大文稿、失败回收等实测。

完成条件：在固定环境下报告分阶段耗时和内存，以小型、常用和压力级文稿对比；先测量再决定缓存、批处理或并发。不在缺数据时承诺毫秒级速度或任意规模支持，也不为基准测试引入常驻服务。

性能实施项开始时，应先约定典型文稿的页数、图片体积、文本量与图表规模，以及可接受耗时和内存，作为验收基线。本文尚未设定这些数值，不能仅凭产生一张测量表就把 G-16 标为完成。

## 7. 修复顺序与验收要求

G-11 的共享检查、绘制与发布已接通；G-01 的共用场景边界、采集、传输和适配已有实现，内部 SVG 正在逐类消费场景。第 2.6 节合成 fixture 冲突已修复、常规回归恢复通过；第 3.5 节对象锚点源码修复及当前 NativeAOT 2/2 已验证，但生产 painter 仍未切换；真实表格/source-bound 独立断言继续作为回归基线。随后仍需完成 3.3/3.4 的生产接入、4.x 检查/发布集成和 5.x 等价/性能证据。只切换输入、只补诊断或只调整测试期望都不能解决关系错误。P0 指会造成误判或破坏证据的问题，P1 指完成静态功能覆盖，P2 指性能与交付完善；P2 不等于最终目标可以不做。

表内 G-12 已完成的限定契约作为后续回归基线保留，不再列为待实现；大文稿、压力条件下的失败恢复由 G-16 继续验证。

| 阶段 | 关联缺口 | 交付结果 |
| --- | --- | --- |
| P0：先停止误报 | G-02、G-06、G-09～G-14 | 待修事实表达与 smoke 失败分类；G-11 准确报告、G-12 不覆盖与真实清单已完成，持续回归 |
| P0：统一语义输入 | G-01、G-04、G-08、G-10 | 将已有 compiler 场景真正接入 SVG，移除组件/数据猜测；保留尚未修复字段的检查与警示 |
| P1：基础元素 | G-02～G-07 | 文字、几何、图片、表格、线、组、组件、源预览的静态显示有真实回归 |
| P1：图表 | G-08～G-10 | 按当前 schema 全部图表及变体逐项核对比例、尺度、层级、方向和缺失点 |
| P1：维护门禁 | G-11、G-13～G-15 | 字段映射与测试自动对齐，Skill 交付区分证据类型，矩阵持续可验证 |
| P2：性能与打包 | G-13、G-16；G-12 为回归基线 | 待补命令行为、压力恢复与可重复基准；G-12 已验证的依赖缺失和发布故障处理持续回归 |

### 7.1 每项功能的最小回归组合

每个新增/修复功能至少提供：

1. 一个正常 authored 输入，确认编译成功且预览真正绘制目标内容。
2. 一个能揭露该功能风险的边界输入，例如缺失数据、负值、复杂关系、嵌套变换或透明边缘。
3. 对支持 source-bound 编辑的功能，从原始源字节重新投影，完成修改/删除/二次投影，并检查非目标内容与包修改范围。
4. 一项直接验证目标行为的断言，例如几何、数据语义、格式、诊断或发布状态；影响视觉内容时附可审阅渲染结果。不能仅检查输出文件存在。
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
- [ ] 字段级映射和生成文档检查进入维护 gate，渲染器不会静默落后于接口和功能更新。
- [ ] 本地性能、依赖缺失和失败恢复完成实测，达到另行确定的使用目标。

### 7.3 后续开发如何与功能更新同步

对象锚点修复已由独立变更完成，并在当前 NativeAOT 包中以 2/2 关系用例复验；后续仍要覆盖嵌套/组件/auto 的复杂边界和生产 route。不依赖它的图表和其他场景映射可分别推进，保留历史失败反例及已通过的表格、折线、圆形图、柱形/条形和源编辑验证。目前清单勾选 7/15，内部绘制已有部分实现，G-01 整项仍未完成。此前实施已向 tasks 追加局部证据；本次不因局部通过而修改 G-01 勾选。G-11 是防误判措施，不能替代后续实现。对某个具体字段可以直接复用现有结果的，不必等待一套大型新框架。

一次修改以一个明确的视觉字段或语义行为为单位，至少把以下记录放进对应变更及覆盖台账：

| 必填记录 | 内容与检查要求 |
| --- | --- |
| 输入与范围 | 写出真实 PPJ 字段路径、默认值、合法取值、适用图表/元素；注明 authored、source-bound 或两者 |
| 语义归属 | 指向当前 validator、compiler/projector 或布局解析代码；区分“接口未支持”和“接口支持但预览遗漏” |
| 绘制与降级 | 说明该字段如何改变几何/文字/像素；哪些输入仍 partial/opaque，具体诊断是什么 |
| 正例与风险例 | 固定输入和预期语义；正例必须进入绘制，风险例包含本字段相关的缺失、关系、透明度或源绑定边界 |
| 检查与产物 | 分别记录结构、渲染、编辑保真结果；保存实际 SVG/PNG、诊断及输入身份，不仅写“测试通过” |
| 同步更新 | 更新实际受影响的 schema/registry、预览映射、fixture、断言、Skill/reference 和本文状态；无关层明确不受影响即可 |

目前已实现类型声明与生成摘要的自动检查；第 6.3 节所述字段到绘制及回归的完整对应 gate 尚未实现，因此表中的完整维护要求仍需人工核对。后续自动化应检查真实字段增删与这些记录是否匹配，不能只比较能力名称是否出现。

按变更范围运行最窄检查：预览发布/绘制先跑 presentation 分段；涉及 wire 再跑 `npm run proto:check`；涉及 Skill/reference 再跑 portability 和 reference-sync 检查；涉及生成矩阵用仓库生成脚本并审查差异。C# 功能测试还须确认所用 codec 包确实包含该改动，不能拿旧二进制的预览成功代替新代码验证。

每项完成时更新对应 G 编号的“已验证行为、未覆盖字段、命令与结果、输入及版本身份”。未达到该编号全部完成条件的，只记部分修复；也不因为一项任务清单全部勾选就将第 7.2 节整体标为完成。

### 7.4 可直接拆分的剩余工作

本节是后续实施建议，不是新增支持声明，也不改变 OpenSpec 的任务范围。第 3、4 节盘点生产缺陷；下表从已有内部实现出发，指出仍需落地的工作。每行可以拆成多个小变更，完成一个字段后保留其他未完成项，不必等待整类全部完成才记录进展。

| 工作包与编号 | 需要补什么 | 最小风险案例与判定 | 主要修改边界 |
| --- | --- | --- | --- |
| 对象关系，G-06 | 源码已处理对象锚点、移动后的端点及坐标作用域；当前包已通过第 3.5 节两个反例；native 已暴露部分 connection-site 字段 | 继续补嵌套/组件/auto、任意 connection-site 的通用绘制/编辑、自动避障和复杂路线；显式坐标连接线不得退化 | 编译器、编译回归、场景/SVG 断言；生产 route 和宿主行为仍独立验收 |
| 普通堆叠，G-08～G-10 | 柱/条已补有限正负累计与风险回归，见第 2.8 节；线/面积、完整缺失策略和源生命周期仍缺 | 保留同向累计、异号不抵消、真实零与反向轴断言；新增系列/删除/复杂源和其余类型 | 消费已编译 native 字段的绘制与测试，不在 JS 重算 PPJ 数据集 |
| 百分比堆叠，G-08～G-10 | 已补非负完整分母、原值/比例对应、缺失与零总量提示；负值、完整显示策略及其他图表类型仍缺 | 保留 25%/75%、大数、缺失/全零和源编辑回归；继续加入有符号反例与复杂源 | 先确认现有 codec 的有符号语义；不自行取绝对值或补零 |
| 数值坐标与组合，G-08/G-09 | scatter 使用实际 X/Y，bubble 使用 X/Y/size；combo 保留系列类型、主副轴和单位；补对数柱基线与平滑曲线 | 非等距 X 必须改变位置；不同 size 必须改变气泡；双轴量纲不得共用一条范围；缺失分别检查通道 | 原生 chart 分支、轴与 mark 映射、真实创建/源编辑回归 |
| 层级、流向和金融图，G-09 | 核对编译生成路径及 native 分支，逐项补 radar、heatmap、treemap、sunburst、waterfall、candlestick、sankey、symbol/stream | 两层父子、汇流后分流、中间 total、OHLC 四通道、固定 symbol 单位各有直接语义断言 | 优先复用 compiler 生成结果；已有少量路径成功只作回归基线 |
| 圆形图剩余项，G-08/G-09 | 多系列/多环适用范围、负值策略、标签/图例与分离随动、源主题及 workbook | 不同环保持各自点归属；缺失不冒充总体；分离后标签不指向其他切片 | circular painter、共有文字/布局及第三方源编辑测试 |
| 文字和形状，G-02～G-04 | 字体度量、换行、段落/列表、AutoFit、溢出、剩余 preset/adjustment 的生产绘制、路径 arc 与效果 | 同段多 run 不拆段；长中文、字号变化、旋转文字与非矩形各验证实际边界 | 场景文字/几何/变换；明确字体环境与允许的栅格差异 |
| 图片与合成效果，G-04/G-05 | 按已知裁切/fit/mask/透明度处理，补边框、阴影、反射与主题颜色 | 透明边缘图、非居中裁切、半透明覆盖、阴影扩边；未知主体范围不能擅自裁切 | 已验证资产到 SVG 的映射与实际 PNG；不增加第二次路径读取 |
| 表格和剩余内容，G-06/G-07 | 合并外围边框、继承/banding、文字排版；diagram 缓存树、媒体海报、OLE/opaque 预览 | 跨行列合并与冲突样式、含可见文字的覆盖格、修改后的源预览；未知内容应可辨识 | native 表格/内容分支；source-bound 原包和非目标 part 保留 |
| 正式接入，G-01/G-11～G-13 | 将生产入口切到实际场景，同时接入 profile 对应事实规则、场景身份清单、警示及输出保护 | 同一 PPJ 的 CLI SVG 真正改变；返回/落盘证据一致；缺场景、栅格失败、写失败仍保留正确状态 | 生产路由、检查与发布；不能只换 painter 后删除全部旧错误规则 |
| 持续验收，G-14～G-16 | 固定正例/负例、修 smoke 分类及目录污染风险，补字段映射 gate、新二进制验证与性能基线 | 全部正例被拒绝时整轮必须失败；新增视觉字段不能仅靠类型名登记过关；记录实际速度/内存 | 测试、registry/生成文档、构建证据；外部 fixtures 与人类校准单独记录 |

其中堆叠缺失值尤其不能只做“跳过 null”：在普通累计中，缺失段可能使后续段的起点无法确定；在百分比图中，未知值可能使分母无法确定。普通柱/条目前对缺失分类保留值与未知位置提示，有第 2.8 节的限定回归；其他堆叠类型和完整显示策略仍待实现，不据此预先认定某种百分比近似正确。

有三项决定需要在相应工作开始时写清楚：对象锚点修复的变更范围；缺失/负值等图表显示策略与 codec 的一致性；G-16 的典型文稿规模及可接受耗时/内存。这些待决项不阻止独立的文字、图片、图表字段或测试维护工作。

当前不存在可以据此宣布整体完成的单一绿色命令。至少要同时满足：实际静态绘制覆盖、第 7.2 节可靠性要求、生产入口接入、指定新运行时的回归、字段维护检查及已约定的性能验收。第三方复杂文稿和人类校准仍需独立证据。

## 8. 与接口补齐、Skill 实验的关系

### 8.1 接口已继续更新，不能混算渲染完成

初次审计时 `ppj-chart-axis-log-base` 尚在推进，清单曾从 0/4 变为 3/4。最新复核中，对数轴实现已提交为 `9efc8ce5`，当前任务清单为 5/5 勾选；趋势线 source-bound 列表编辑也已提交为 `fbddc7e7`。本节不再把它们笼统记录为“仅有在途文件”。

本轮没有重新执行这些 C# 功能的 NativeAOT 构建和专项验收，也未核验清单勾选对应的全部原始日志。“已提交/清单已勾选”是仓库事实，“已由本次独立验收”则不是。当前生产 preview 仍未读取 `logBase`、`trendlines` 或 `errorBars` 来绘制对应内容；内部 line 已有有限 logBase 几何映射及合成测试，但趋势线和误差线仍未绘制。

此前收尾时新增的 errorBars 预览诊断目前已在实际绘制代码中：普通图表兜底分支对存在 errorBars 的系列增加 `chart-error-bars-not-rendered` 的 partial 诊断，最近 circular 实施的 presentation 分段也包含相应回归。这是诊断补充，不是误差线绘制实现；提前返回的其他图表分支仍不能据此宣称全部覆盖。

G-11 实施轮次从 `09f79e04`（literal custom error bar data）开始，另一工作流随后提交 `0e1d8c8f`（固定公式误差数据与内嵌 workbook 同步）；之前还有 `ccb93795`（source-bound scalar error bar lifecycle）。这些是可见提交事实，不等于本审计的 C# 独立验收。预览尚未绘制这些误差数据，不能用接口提交代替绘制完成。该轮随后又出现 `ppj-chart-trendline-label` 在途工作；生成矩阵的漂移检查捕获其新增 schema，已重新生成并通过检查，但没有将标签接口或绘制记为已验收。

上述对数轴变更的范围是 value axis 的 logBase 创建、读取、修改和删除，含 numeric X 与 combo 副轴、token、非法值及所有权保护。具体跟踪见 [对数轴任务](../openspec/changes/ppj-chart-axis-log-base/tasks.md)。

该接口即使通过 codec 回归，当前生产 preview 也未消费其缩放语义。内部 line 的有限进展尚不能代表 numeric X、combo 副轴等全范围同步。它是“接口更新后渲染需要同步”的具体例子，不应以接口完成代替渲染完成。本文不修改其他工作流正在编辑的文件，只核对与预览相关的证据范围。

前次文档复核补记：趋势线标签实现已提交为 `93b89e67`，不再仅是 G-11 实施轮次观察到的在途工作。预览与发布器源码 hash 没有变化；此提交不代表趋势线标签已在自有 SVG 中绘制。该次复核没有重跑该接口的 C# 专项或构建对应 NativeAOT。

前次实施快照补记：起点 HEAD 已包含 `f3a67617` 的趋势线标签手动布局及 `3eb58b3a` 的数字格式链接保留，当时工作区另有趋势线富文本改动；收尾时该功能已由另一工作流提交为 `3b6272dd`。它们继续扩大接口与预览的同步检查范围；本文只核对提交和差异归属，不将其算作已完成的趋势线视觉渲染或该次 C# 验收。

此前快照已包含 `307106dc` 的 chart text language、`ade46bd5` 的 preview bindings 归属调整、`e092c458` 的 chart text strike 与 `654cb3e5` 的有符号 baseline 偏移。当前生产 CLI 的 SVG 仍走旧绘制路线；图表文字语言、删除线、基线偏移、趋势线标签富文本、布局和格式是否进入最终画面仍要单独验证。本轮未重新执行这些接口的 C# 专项，不能因为字段已加入接口或生成矩阵，就认定渲染器已同步支持。

历史场景适配复核补记：当时读到 chart text 的 letter spacing（`f3ebbb34`）、kerning（`4782e6e5`）及独立 highlight paint（`aee69e35`）提交。`8b5b6771` 随后提交了只读场景适配层；这四项都没有改变该轮核验的 `svg-preview.mjs` 文件摘要。后续应检查这些字段在“原生保留 → 适配保留 → 实际绘制 → 行为断言”各层的状态，而不是把接口提交或适配无损当成对应文字效果已经显示正确。本轮文档复核也未对这些图表文字接口重新运行 C# 验收。

当前 HEAD 已包含 chart text outer shadow 与可选属性保留（`90202df2`）、glow/outer-shadow 组合（`5837c51e`），以及 soft edge/direct effect theme colors（`3ab96832`）。生产 painter 未绘制这些文字效果；内部 line/pie/doughnut/column/bar 虽已进入文字标题路径，也没有实现完整效果，其余 native chart 枚举仍为占位。它们是需要纳入 G-04/G-08/G-15 的具体同步项，不是新增渲染通过案例。本轮没有重建/验收这些 C# 新功能；指定临时包的摘要见第 1.3 节。

当前基线另含 inner shadow（`e75423c8`）、reflection/native angle edits（`0373a7cb`）、reflection gradient positions（`3f390e71`）和 direct chart reflection transforms（`dc5e725c`）。这些提交不会自动变成 SVG 的内阴影、反射或渐变绘制；本次只确认其存在及 painter 的未消费边界，不将旧临时 NativeAOT 包的集成通过当作这些新增接口的验收。

前次文档复核的 HEAD 为 `caa7cc95`（chart shadow scale and skew），当前为 `1fe503cd` 加柱形/条形工作区修改。`caa7cc95` 修改了 schema、wire、编译/投影、阴影 codec 和对应测试，但在该轮没有改变生产与内部 painter 的文件摘要。新增阴影缩放/斜切仍属于 G-04/G-08/G-15 的同步缺口：接口可表达不等于 SVG 已显示，至少需要原生字段保留、实际阴影变换及边界断言；本次未运行此提交的 C# 测试或重建运行时。

### 8.2 更广泛 PPJ/PowerPoint 差距

跨 family 图表、完整 ChartPart/workbook 所有权、复杂主题继承、SmartArt、动画/交互、布局求解器及宿主行为等更大范围，继续由 [PPJ/Kimi/PowerPoint 差距台账](ppj-kimi-pptd-full-ppt-gap-backlog.zh-CN.md) 跟踪。

该台账的历史基线和 bounded 完成声明不能直接当作本次最新全仓验收。本文核实的是预览实现与当前输入模型之间的差距，不宣称已经逐项重新验收所有 C# native leaf、Office Live 适配器或完整 PowerPoint 功能。

### 8.3 Skill 路由实验仍是独立验收

本次未重新冻结案例、准备六个外部 PPTX fixture、执行四路线作者任务、两轮盲评或人类校准，也没有重新核对独立评测工作树。因此本文不能作为 Kimi 路由默认启用或原 Skill 实验完成的依据。下述路径与评测清单观察来自前次审计，仅保留为恢复入口。

恢复实验时仍需核对：各路线 1→10 编辑任务、每场景的缺失数据/复杂关系/source-bound 输入、事实与视觉分开评分、硬门槛失败处理、报告与 tasks.md 一致。具体完成状态需在实验工作区重新确认；本文不把以前的计划当作已执行事实。

前次审计确认存在独立评测工作树 `/home/zenfun/mywork/OfficeKit-ablation`，跟踪入口为该工作树内的 `openspec/changes/presentation-skill-ablation/tasks.md` 和 `evals/presentation-skill-ablation/`。该绝对路径仅是历史定位信息，不是运行时依赖；在其他机器可用 `git worktree list` 查找相应工作树，再按仓库内相对路径访问。

该目录还存在 `report.v1.md`、`report.tri-route.v1.md`、`report.multi-route.v1.md` 和多个版本的案例/证据文件。文件存在不证明研究完成；恢复时应先确定本轮采用哪个冻结版本，不能合并不同版本的统计作为同一轮结果。

前次补读该 tasks.md 时，执行研究与分析报告部分仍有未勾选项；其作者任务文字仍写 Shared/Kimi 两臂，不能直接当作后来要求的四路线执行清单。恢复时需要重新确认这些观察是否仍成立，再对齐任务、最新 cases/rubric 和实际 evidence。本文未全面检查该工作树内的实验产物。

## 9. 复核入口

### 9.1 可重复执行的检查

从仓库根目录运行以下检查；命令列表不代表本次全部执行，实际结果以第 1.3 节为准。某一步失败时可以独立运行其余脚本定位，但最终仍须重跑完整分段：

```sh
node test/gate-policy.mjs
node test/ppj-preview-scene-svg.mjs
npm run test:slow -- --segment presentation
node scripts/generate-ppj-preview-capabilities.mjs --check
node scripts/generate-presentation-capability-matrix.mjs --check
node test/ppj-preview-smoke.mjs
```

单项定位时，可分别运行 `node test/ppj-preview-diagnostics.mjs`、`node test/ppj-svg-preview.mjs`、`node test/ppj-preview-capability-coverage.mjs` 或 `node test/ppj-preview-output-evidence.mjs`；这四项已包含在 presentation 分段中。此处列的是可复跑命令，本次文档复核的执行范围见第 1.3 节，历史结果见第 6.1 节；不要把列出全量 smoke 命令当作已重跑。

真实运行时集成必须显式指向仓库构建命令生成的包；脚本现包含传输/适配、内部 SVG 基础绘制、表格/对象关系、有限 line、单系列 pie/doughnut、column/bar（普通/堆叠/非负百分比）及 literal `arcTo` 的断言。最后一次已验证的 `495daafc` 快照包 `/home/zenfun/mywork/OfficeKit/tmp/officekit-preview-runtime-495daafc` 运行退出 0，报告 `/tmp/officekit-native-scene-paint-MTxwlx/integration.json` 为 `passed`，对象锚点 2/2、`relationFailures=[]`；旧临时包的锚点失败仅保留为历史反例。**这个通过只覆盖内部 scene painter，不代表生产 route 或完整字段覆盖。** 当前 HEAD 的新提交尚未重建，普通 presentation 分段不自动重建运行时，也不会自动执行这个集成脚本。准备好 `global.json` 固定的 SDK 后可单独运行：

```sh
preview_package="$(mktemp -d)"
npm run build:office-kit -- --output "$preview_package" &&
  node test/ppj-preview-scene-native.mjs "$preview_package"
```

构建成功后才运行第二条测试。测试核验包的 manifest/可执行文件 hash，输出实际 PPJ 二进制身份，并比较轻量和生成 wire 的真实回执；打包 Office profile 应拒绝 PPJ 请求。脚本还调用内部 painter/sharp，验证一组配对像素、实际生成路径及候选文字，并输出独立临时目录。C# 通用协议入口用下方原生专项验证，不能把它等同于 Office 可执行文件。以上不验证生产入口的场景切换、正式发布对接或完整功能，也不替代任务 5.2 的 `npm run verify:office-kit-build`。

测试会在操作系统临时目录写预览；smoke JSON 打到 stdout。保存输出时应使用新的证据位置，不覆盖输入。若出现 `spawnSync rg EPERM`，应先解决运行环境权限，不要将环境错误记录成绘制语义错误。

检查结果应同时保留输入数量、真正渲染数量、前置拒绝原因、渲染错误、实际产物与诊断；还应补记具体运行的 codec、字体和依赖版本。同一轮中分段和单项的结果不应重复计数。上述命令不是总体验收的替代品。

G-01 的 C# 基础与 authored 接入可单独复跑。先确认仓库根目录 `dotnet --version` 为 `global.json` 固定的 **8.0.128**；其他版本会因 `rollForward: disable` 被拒绝。可以使用自行安装的该版本 SDK 的绝对路径，不应修改 global.json 来绕过检查，也不能把此前 `/tmp` SDK 的存在当作跨机器前提。

```sh
# 协议、collector 与 envelope 基础；此前 16/16 记录对应这一测试类。
dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj \
  --filter 'FullyQualifiedName~OfficeKit.Codec.Tests.PpjPreviewSceneTests' \
  /p:SkipGetTargetFrameworkProperties=true /clp:ErrorsOnly

# authored writer 场景测试；当前包含 12 个用例。
dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj \
  --filter 'FullyQualifiedName~OfficeKit.Codec.Tests.PpjPreviewAuthoredSceneTests' \
  /p:SkipGetTargetFrameworkProperties=true /clp:ErrorsOnly

# 所有场景专项：当前源码包含 16 基础 + 12 authored + 11 candidate。
dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj \
  --filter 'FullyQualifiedName~PpjPreview' \
  /p:SkipGetTargetFrameworkProperties=true /clp:ErrorsOnly
```

测试数量可能随代码推进增加，以本次命令实际输出为准。NuGet 恢复或 SDK 失败应记为环境/构建失败，不能写成目标断言失败或通过。`495daafc` 快照包已按 G-01 任务 5.2 的 checked-in workflow 重建并由 JavaScript 集成脚本实际加载；当前 HEAD 的新增提交和未提交修改尚未重建。剩余是生产入口接入、完整绘制/等价证据和性能/宿主验收，不得把这条内部通过写成整体完成。

### 9.2 主要代码与文档

| 文件 | 核对用途 |
| --- | --- |
| [svg-preview.mjs](../src/ppj/svg-preview.mjs) | `drawElement` 绘制、`assessedDrawing` 检查合并及页面警示 |
| [内部 native painter](../src/ppj/preview-scene-svg.mjs) | 3.3 基础几何/文本/图片/组及 literal Bezier/`arcTo` 路径，3.4 部分表格/连接线/原生 line/pie/doughnut/column/bar；当前 NativeAOT 集成已跑通，但不是生产 CLI 或完整 G-11/G-12 集成 |
| [内部绘制专项](../test/ppj-preview-scene-svg.mjs) | native 字段到 SVG 的直接断言、明确限制及独立进程惰性加载 |
| [preview-output.mjs](../src/ppj/preview-output.mjs) | G-12 非覆盖发布、失败状态、文件与输入身份清单 |
| [svg-preview-capabilities.json](../src/ppj/svg-preview-capabilities.json) | 当前 registry 派生的类型级保守声明 |
| [preview-capabilities.mjs](../src/ppj/preview-capabilities.mjs) | registry/schema 对齐、声明生成与漂移检查 |
| [preview-diagnostics.mjs](../src/ppj/preview-diagnostics.mjs) | 绘制和发布共用的诊断结构与可靠性聚合 |
| [preview-input-assessment.mjs](../src/ppj/preview-input-assessment.mjs) | 实际字段、继承、metadata 与源载荷边界检查 |
| [preview-factual-errors.mjs](../src/ppj/preview-factual-errors.mjs) | 当前绘制缺陷对应的错误级规则；不是绘制修复 |
| [G-11 任务清单](../openspec/changes/ppj-preview-support-diagnostics/tasks.md) | 9/9 任务、最终门禁和明确保留的其他 gap |
| [G-01 设计](../openspec/changes/ppj-preview-compiler-scene/design.md) | 共用 native 场景、实际候选导入、身份及预算的方案；不能用规划代替实现证据 |
| [G-01 任务清单](../openspec/changes/ppj-preview-compiler-scene/tasks.md) | 7/15 已勾选；传输和适配完成，场景绘制、SVG 等价和完整运行时验收仍待完成 |
| [场景 collector](../native/OfficeKit/src/OfficeKit.Codec/PpjPreviewSceneBuilder.cs) | native 视觉字段复制、预算、摘要和载荷隔离 |
| [writer observer](../native/OfficeKit/src/OfficeKit.Codec/PpjPreviewSourceFreeBuildPlan.cs) | authored 同次 materialization 采集及节点绑定 |
| [authored 原始归属](../native/OfficeKit/src/OfficeKit.Codec/PpjPreviewOrigins.cs) | 现有 expansion/slot 替换中的只读 origin 跟踪，不改 JSON 和旧 node map |
| [authored 场景测试](../native/OfficeKit/tests/OfficeKit.Codec.Tests/PpjPreviewAuthoredSceneTests.cs) | 12 个实际编译、上下文、图表、writer 生命周期及嵌套/repeat/slot 归属案例 |
| [candidate 场景生产](../native/OfficeKit/src/OfficeKit.Codec/PpjPreviewCandidateScene.cs) | 实际候选 native 导入、资产复用、确定归属及未知身份的保守处理 |
| [candidate 身份匹配](../native/OfficeKit/src/OfficeKit.Codec/PpjPreviewCandidateBindings.cs) | 已实现物理 part/对象身份连接；重复和缺失身份不授予归属 |
| [candidate 场景测试](../native/OfficeKit/tests/OfficeKit.Codec.Tests/PpjPreviewCandidateSceneTests.cs) | 当前 11 个候选编译、文件复用、快照隔离、重排、嵌套组、删除、overlay 和歧义身份案例 |
| [ppj-v1.schema.json](../src/ppj/ppj-v1.schema.json) | 实际元素、图表数据和样式字段；不能根据渲染器字段反推 schema |
| [workspace.mjs](../src/ppj/workspace.mjs) | 资源读取、安全检查及 compile 转发 |
| [PpjAuthoredPresentationCompiler.cs](../native/OfficeKit/src/OfficeKit.Codec/PpjAuthoredPresentationCompiler.cs) | CanonicalJson、expansion/build plan 和编译回执的区别 |
| [PpjPresentationCompiler.cs](../native/OfficeKit/src/OfficeKit.Codec/PpjPresentationCompiler.cs) | source-bound 编译路径和回执 |
| [render-review.mjs](../src/ppj/render-review.mjs) | LibreOffice/Poppler 路线、独占输出和结构 review 边界 |
| [cli.mjs](../src/ppj/cli.mjs) | preview 参数、handler 与格式化返回 |
| [canonical fixture](../test/fixtures/presentation/evidence-ledger-canonical.ppj) | 当前测试通过但仍显示错误的具体输入 |
| [预览 smoke 测试](../test/ppj-svg-preview.mjs) | 文件/页数/稳定 ID、真实 source-bound 输入及 hash 断言 |
| [发布故障测试](../test/ppj-preview-output-evidence.mjs) | 非覆盖、故障恢复、资源快照和惰性依赖的负向测试 |
| [覆盖声明测试](../test/ppj-preview-capability-coverage.mjs) | schema 类型、等级互斥、嵌套扫描与生成摘要检查 |
| [共享诊断测试](../test/ppj-preview-diagnostics.mjs) | 基础聚合及字段/事实/实际绘制发布检查套件入口 |
| [输入检查测试](../test/ppj-preview-input-assessment.mjs) | 嵌套/继承、未知字段、metadata、源载荷不展开及输入不变 |
| [字段限制测试](../test/ppj-preview-field-limits.mjs) | 有限 PNG/SVG 映射正例、单字段变更和各类限制路径 |
| [事实错误测试](../test/ppj-preview-factual-errors.mjs) | 已登记错误类别及配色不抵消失败的独立回归 |
| [全量 smoke](../test/ppj-preview-smoke.mjs) | 42 个输入的运行路径与错误分类 |
| [常规 test gate](../scripts/run-test-gate.mjs) | 预览专项检查是否进入持续验证 |
| [能力矩阵生成器](../scripts/generate-presentation-capability-matrix.mjs) | 计数和 visualStatus 的来源 |
| [生成的能力矩阵](presentation-capability-matrix.json) | 快照数据，不等同于已验证覆盖 |
| [Presentation 开发规范](../skills/presentations/AGENTS.md) | 功能、文档、预览、保真和测试的共同完成要求 |

维护本文时，应逐项更新证据和状态。修好某个 G 编号不代表相关类型的所有字段完成；只有达到第 7 节要求，才可以收回“整体未完成”的结论。
