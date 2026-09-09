# OfficeKit 本地 PPT 预览渲染器差距审计

审计日期：2026-09-09。状态：**预览链路已打通，功能覆盖和可靠性尚未完成**。

最新实施复核：2026-09-09，从 HEAD `ade46bd5` 的工作区继续 G-01；核对期间另一工作流提交 chart text strike 为 `e092c458`。G-01 的任务清单为 **5/15 已勾选**：2.3 原生节点归属已完成，组件原始路径、嵌套 repeat/slot、矢量子节点 owner、源对象重排和未知身份均有回归。JavaScript 仍从 canonical JSON 绘制，尚未消费新场景，也未重建并验收它实际使用的 NativeAOT。当前检查记录见第 1.3 节，实施边界见第 2.3 节；整体渲染目标仍未完成。

前次文档复核：起点 HEAD 为 `93b89e67`；核对期间另一工作流提交 G-11 实现 `f734890e` 和 G-01 规划 `5118877b`。本文保留初次审计和 G-11 实施轮次的记录；历史测试不得算作最新实施的新测结果。提交规划不等于实现完成；此前未提交状态仅描述当时快照。

| 差距 | 已实施并验证 | 剩余工作或范围边界 |
| --- | --- | --- |
| G-12 输出安全与证据 | 已完成本项限定的发布契约：新目录独占发布、文件 hash/清单、pending/final 生命周期、失败保留、源/候选/资产身份；最新 gate-policy 与 slow/presentation 四项测试通过 | 不包含视觉正确性、断电持久性或恶意进程替换目录树的安全保证 |
| G-05 图片的有限状态 | 资产复用已加载字节；空资产诊断；独立测试验证一个显式 contain PNG 的 SVG 坐标、透明度和内嵌字节映射 | 不能推广到图片解码、裁切、主体范围、边缘、mask 和效果均正确 |
| G-11 支持诊断 | 已接通 registry、字段检查、绘制、SVG/PNG 警示及发布清单；事实错误不因配色或发布成功而清除 | 只完成可靠性检查链路，不修复 G-01～G-10 的绘制缺陷；验收见第 5.1 节 |
| G-01 编译场景 | 协议、collector、writer 同次采集、实际候选导入及原生归属完成，5/15 | JS 消费、真实 SVG 等价与所用 NativeAOT 仍待验收；未知身份继续 unmapped，不授予编辑权限 |
| 其余 G 编号 | 本轮未修改 JS 绘制语义 | 按各节完成条件继续逐项补齐 |

G-12 使用方法与清单字段见 [预览输出说明](ppj-preview-output.md)，实施清单见 [ppj-preview-output-evidence](../openspec/changes/ppj-preview-output-evidence/tasks.md)。以上不是全部 gap 的完成声明。

G-11 后续实施记录（2026-09-09，9/9 任务完成）：[支持诊断变更](../openspec/changes/ppj-preview-support-diagnostics/tasks.md) 已从独立模块推进到真实绘制和发布。能力摘要和生成矩阵同源生成，按真实 schema 校验 16 类元素、16 类图表，另列变体和源绑定字段；不再把 shape/image/connector 等整个类型宣称为完整支持。当前类型声明见第 3 节，具体实现与验收见第 5.1 节。

新增诊断测试已接入 fast/slow；本次文档复核 `slow/presentation` 重跑 4/4 通过，真实 codec+sharp 产物见第 1.3 节。`renderPpjToSvg` 已调用字段与事实检查，页面、全局及输出清单共享结果；真实 PNG 警示像素和文件 hash 均有断言。G-14 的完整视觉覆盖和 G-15 的字段维护闭环仍开放。

本文面向 OfficeKit 维护者和后续接手开发的 Agent，记录本地 SVG/PNG 预览距离“简单、覆盖现有功能、能可靠辅助结构与视觉 review”的实际差距。文档记录问题和验收要求，不代表这些问题已经修复，也不代替实施规格。

### 差距总表

“未完成”表示仍有本编号的核心要求未满足，不表示整项完全没有实现。G-12 的限定契约和 G-11 的检查契约已完成，其余 14 类绘制、交付、完整覆盖及性能工作仍开放。这不是按功能数量或工作量计算的完成率。

| 编号 | 当前状态 | 主要影响 | 收口时必须拿出的证据 |
| --- | --- | --- | --- |
| G-01 共用解析与布局 | 5/15 已勾选；真实 writer/candidate 采集及原生归属已验证，JS 尚未消费 | 高层组件与样式仍可能和编译后的 PPTX 不同 | 高层写法与等价展开写法的几何、数据及样式一致；场景必须来自实际编译结果 |
| G-02 文字与形状 | 未完成 | 带文字形状消失，富文本错分行，几何被替换 | 几何与文字同时存在；run、段落、字号和溢出有断言 |
| G-03 变换与可见性 | 未完成 | 旋转、镜像、嵌套组及 hidden 表达错误 | 嵌套坐标、旋转边界、可见性与实际展开结果一致 |
| G-04 样式与主题 | 未完成 | token、继承、背景和 effects 被固定默认值替代 | 样式解析优先级及具体填充、透明度、轮廓回归 |
| G-05 图片 | 部分修复 | 资源快照已统一；裁切、主体范围、mask 和效果仍不可靠 | 已知裁切和透明边缘案例；未知信息的明确诊断 |
| G-06 表格与连接线 | 未完成 | 列宽、合并单元格和连接方向不符合输入 | 非等宽与 span 几何、端点移动、锚点和箭头断言 |
| G-07 源绑定与 opaque | 未完成 | 源快照可能不能说明编辑后状态；静态检查范围不清 | 原包保留、目标修改、重新投影及快照有效性证据 |
| G-08 图表公共语义 | 未完成 | 数据映射、尺度、双轴与标签可能误导读者 | 每个数据点可对应正确通道、尺度、轴和标签 |
| G-09 图表类型 | 未完成 | 比例、层级、累计值、流向及 OHLC 可能错误或消失 | 第 4.2 节每类图表至少一个真实语义反例转为通过 |
| G-10 缺失值 | 未完成 | null 被补 0、跨缺失连线，或独立观测被省略 | 真实 0、连续缺失、单点段及显式显示策略回归 |
| G-11 支持状态 | 本项检查契约完成，9/9 任务通过 | 已知事实错误变为失败；partial/opaque 保留限制和图片警示 | 字段/页/全局与发布结果一致；事实与视觉分开；不代替绘制修复 |
| G-12 输出安全与证据 | 限定契约完成 | 已保护已有路径、保留失败产物并绑定实际输入 | 已有故障注入与真实 authored/source-bound 发布测试；边界见第 5.2 节 |
| G-13 CLI 与交付链 | 未完成 | `--pages` 被忽略，预览结果未闭合各类 review | 参数行为、摘要输出、结构/渲染/编辑保真独立状态 |
| G-14 测试覆盖 | 部分修复 | 已接入门禁，但仍缺“画对”断言和正例覆盖门槛 | 正例实际渲染、负例独立分类、语义及视觉断言 |
| G-15 持续维护 | 部分修复 | 类型和生成摘要已防漂移；新增视觉字段仍缺完整绘制回归约束 | schema/registry 到预览、fixture、断言的自动对应检查 |
| G-16 性能与稳定性 | 未验证 | 尚不能承诺日常多页、大图、长文的速度和内存 | 固定环境下的分阶段耗时、峰值内存及压力结果 |

## 1. 结论、范围和证据口径

### 1.1 当前结论

`officekit ppj preview` 已能读取 PPJ，调用现有编译器，然后生成 SVG、PNG 和 `render.json`。这证明了一条可运行的本地预览链路，但不能证明预览忠实表达了输入。

目前仍存在会改变内容含义的绘制错误：带文字的形状丢失几何本体、饼图不表达数值比例、连接线不依据端点关系、复杂图表读取错误的数据字段。这些问题不能用“只是视觉近似”解释。G-11 已为已登记的错误生成失败级可靠性和图片警示，未知视觉字段保守报告限制；诊断没有使图形本身变正确。

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

本次继续实施 G-01 的 2.3，并保留其他工作流的改动。以下是最新实施证据；随后保留的文档复核和较早实施记录均为历史证据，不能合并成一次全仓验收。

| 最新 2.3 实施复核 | 2026-09-09 记录 |
| --- | --- |
| 基线与范围 | 起点 `ade46bd5`，最新读取 HEAD `e092c458`；归属源码和测试仍在工作区。本轮未提交或推送 |
| authored 原始路径 | 在已有 expansion 中按实际 typed model/JSON clone 身份记录 owner，跟随真实 slot 替换；保留原始 JSON 路径与独立 scenePath，不改 canonical JSON 或旧 node map |
| 生成节点 | 组件展开节点标为 Generated；矢量图表子节点继承真实 chart owner；两层 repeat 的 instance/repeat 标识保持可区分 |
| source-bound 身份 | 原包与最终候选通过 part/cNvPr 身份连接；验证页/元素重排、嵌套组 readingOrder、删除及追加 overlay。未知或歧义身份继续明确 unmapped |
| 原生回归 | 45/45 通过：39 个预览用例（16 基础、12 authored、11 candidate）加 6 个既有组件、组排序及 canonical 展开回归；随后预览单独重跑 39/39，不重复计数。开关前后候选、canonical、node-map 字节不变有断言 |
| JS 回归 | presentation 4/4 通过，真实产物 `/tmp/officekit-ppj-preview-yxDA9B`；仍为旧 canonical-PPJ 绘制路线，不是新场景端到端验收 |
| 维护门禁 | gate-policy、两个能力生成器 `--check`、strict OpenSpec 与差异空白检查通过 |
| 测试输入调整 | “删除源对象并同时追加 overlay”被既有源前缀约束拒绝；拆成两条受支持路径后通过，没有放宽 compiler 规则或把前置拒绝记成渲染成功 |
| 下一项及未执行 | 3.1 JS 传输和完整性校验；尚未重建 NativeAOT、接通新场景 SVG、运行完整 smoke/全仓测试、外部视觉验收、人类校准或性能基准 |

#### 前次仅文档复核

以下为 33/33 时的记录；其中“当前”“本次”均指该次文档复核，不覆盖上表的新进展。

| 本次文档复核项目 | 2026-09-09 实际记录 |
| --- | --- |
| 代码基线 | HEAD `ade46bd5`；已有未提交的 candidate bindings、import/project/compile 接入、候选测试和任务记录。本文反映该工作区，不代表已发布包或 origin/main 状态 |
| G-01 清单 | 4/15 已勾选；2.3 未勾选。任务末尾 2.2 的“全部 unmapped”是当时实现记录，当前已有下述部分进展 |
| 当前 source-bound 归属 | 代码以源投影与最终候选的 `(part 路径, 非零原生对象 ID)` 对齐，恢复请求中的页/元素身份和路径；重复或无法确认的身份不授予归属。已增加页面/元素重排及歧义身份测试 |
| 当前 authored 归属 | observer 仍直接复制 expansion 的 `ProgramPath`，组件路径可能含 `#component(...)`；它不是可直接定位原始 JSON 的路径，2.3 仍需补齐 |
| 实际 SVG 输入 | `svg-preview.mjs` 仍解析 `compiled.programJson`；`native.mjs` 尚不转发 scene 选项或回执。不能用 C# 场景测试证明当前图片已经使用该场景 |
| 原生专项 | 本次使用 SDK 8.0.128 重跑 `FullyQualifiedName~PpjPreview`：33/33 通过，0 失败、0 跳过；16 基础、9 authored、8 candidate。包含新增重排/歧义身份断言，不是 NativeAOT 或完整 2.3 验收 |
| JS 回归 | `npm run test:slow -- --segment presentation` 4/4 通过；真实 authored/source-bound 产物位于 `/tmp/officekit-ppj-preview-eniX9e`，仅为本机临时证据 |
| 维护检查 | gate-policy、预览能力摘要 `--check`、能力矩阵 `--check` 均通过 |
| 文件身份 | 本次重新计算 renderer、publisher、canonical fixture 的 SHA-256，与下方 G-11 记录一致；实际 JS painter 没有随场景实现变化 |
| 未做的验收 | 未重建 NativeAOT、未运行新场景到 SVG 的端到端测试、全量 smoke、全仓测试、宿主视觉验收、人类校准或性能基准；未重跑 proto 生成检查 |

本次原生命令使用 `/tmp/officekit-preview-sdk-hfWOmf/dotnet`，指定 `DOTNET_CLI_HOME=/tmp`、`--no-restore`、`/p:SkipGetTargetFrameworkProperties=true` 和 `/clp:ErrorsOnly`。临时 SDK 路径仅记录本次环境；跨环境复核按第 9.1 节准备固定版本。33 个用例属于同一次筛选运行，不与下方历史 31 个相加。

#### 前次 G-01 2.1/2.2 实施记录

以下为 31/31 测试时的历史快照。任务完成只适用于其明确范围，不等于新场景已用于实际 SVG 绘制。

| 前次 G-01 实施复核项目 | 2026-09-09 当时记录 |
| --- | --- |
| 基线 | 从 HEAD `3b6272dd` 的未提交场景实现开始；收尾 HEAD `730071c6` 已包含场景源码/测试/协议基础，其与已测源码无差异；最新文档与任务记录尚有未提交更新。本轮未执行提交或推送 |
| 原生专项 | SDK 8.0.128；`dotnet test … --filter FullyQualifiedName~PpjPreview --no-restore /p:SkipGetTargetFrameworkProperties=true /clp:ErrorsOnly`，31/31 通过：16 基础、9 authored、6 candidate |
| authored 新证据 | 两种 native 协议入口的真实 minimum/canonical 编译开关前后字节一致；vector/native 图表、缺失索引、母版/布局背景、显式 grammar 文字优先级、普通/Morph writer 输入及页面释放均有断言 |
| candidate 新证据 | no-op、文字/位置叶编辑、语义 fill 与文件复用；对照单独的新导入，核对实际内容、源/候选 hash、非目标 ZIP、opaque 顺序和源 XML、资产字节；私有 PPJ 快照不代替 native 导入 |
| 新修复 | 普通编译接受的大写资产 SHA-256 不再导致场景失败；仅规范化场景证据，调用者资产不变 |
| JS 与维护门禁 | presentation 4/4、gate-policy、buf lint、能力摘要/矩阵 `--check`、strict OpenSpec 与差异检查通过；真实旧预览产物 `/tmp/officekit-ppj-preview-t0kuO7` |
| 收尾协议检查 | `npm run proto:check` 执行 lint 和生成成功，最终 git-diff 检查退出 1：另一工作流正在增加 chart text `language` field 12，生成绑定反映其未提交协议差异。随后 scene wire 回归与差异空白检查通过；没有暂存来制造全绿 |
| 未执行或未实现 | 未重建 NativeAOT；JS 尚未转发/消费新场景；未重跑全量 smoke、全仓测试、宿主验收、Skill 全套、人类校准和性能基准 |
| 当时下一项 | 2.3 完整节点归属：该快照的 candidate 节点仅有 native ID、scenePath 和明确 unmapped；当前部分进展见上表和第 2.3 节 |

本轮先确认既有 23/23，再补上下文、资产及候选测试。测试开发期间修正了 native 文字样式所在层、必须显式声明的 grammar 继承和导入空文本矩形的选择条件；最终 31/31 为上述限定范围的成功结果，没有修改 compiler 语义来迎合测试。JS 分段仍使用原有可加载 codec，不能算新场景端到端验收。

31/31 的代码快照对应随后提交的 `730071c6` 场景实现。收尾又出现 chart text language 的共享 proto/compiler 改动，本轮仅确认其差异归属并验证新生成绑定仍通过 scene wire 回归，未重新验收该在途功能。不要把本轮原生结果推广到收尾所有并行改动；生成绑定在该次检查后的 SHA-256 为 `6017156ba89cd81c5ae16a0c2fdfb65a05bd349a086220e7b98e935eb4b14c90`。

前次仅文档整理的记录如下；表内“未执行 C#”等仅描述该次整理，不再代表最新实施状态：

| 前次文档复核项目 | 2026-09-09 当时记录 |
| --- | --- |
| 代码基线 | 起点 HEAD `3eb58b3a`，当时场景协议、collector、authored observer 和趋势线富文本等存在未提交改动；收尾时另一工作流已提交趋势线富文本，HEAD 为 `3b6272dd`。本文反映各次读取的工作区，不等于发布包状态 |
| 当前执行链路 | `svg-preview.mjs` 仍从 `compiled.programJson` 作图；`native.mjs` 与 workspace 没有场景选项/回执转发 |
| authored 场景代码 | `PpjAuthoredPresentationCompiler` 已按选项包装原 build plan，在 writer 回调采集页面并设置 `receipt.PreviewScene`；新增测试覆盖实际编译、两种图表表示和页面生命周期，但本次未执行 C# 专项 |
| source-bound 场景代码 | 未发现最终候选导入后构造 PreviewScene 的接入；当前 compiler/projector 的在途差异属于趋势线富文本，不能当作 G-01 source-bound 完成 |
| 实际执行 | gate-policy 通过；slow/presentation 4/4 通过；预览能力摘要与能力矩阵的 `--check` 均通过 |
| 本次真实产物 | `/tmp/officekit-ppj-preview-2jRh8Q`，来自实际 codec+sharp 的 authored/source-bound 测试。临时目录仅供本机复核，不是已归档、跨环境可复现的证据包 |
| 文件身份复查 | renderer、publisher 和 canonical fixture 的 SHA-256 均与下方 G-11 表一致；实际 SVG painter 没有随在途 C# 场景接入而更新 |
| 未执行 | 全量 smoke、其余 fast/slow、C# 专项、NativeAOT 重建、proto 生成/检查、Skill 全套检查、外部 Office/宿主验收、人类校准与性能基准 |
| 修改边界 | 只整理本审计文档；保留既有改动，不修改 tasks 勾选，不提交、推送或变更 Skill 默认路由 |

以上通过的是限定回归，不能消除下文绘制反例。尤其是本次 JavaScript 测试使用现有可加载 codec，没有证明未提交 C# 改动已进入该二进制。

以下为前次文档复核、G-11 实施复核和初次快照；其中的“本轮”均指各表对应的历史轮次。

| 前次文档复核项目 | 记录 |
| --- | --- |
| HEAD | 起点 `93b89e67`；收尾复查为 `5118877b`，其前一提交为 G-11 实现 `f734890e` |
| 工作区范围 | 起点的 G-11 模块和 G-01 规划在核对期间被另一工作流提交；G-01 的 15 项任务仍全部未勾选。当前还有独立的趋势线标签布局规划，不纳入渲染实现验收 |
| 主链路核对 | `renderPpjToSvg` 仍解析 `compiled.programJson`；wire 与 JS compile 转发没有 preview scene 字段 |
| 文件身份 | 本轮重新计算渲染器、发布器、canonical fixture 的 SHA-256，与下方 G-11 记录完全相同 |
| 实际重跑 | gate-policy、slow/presentation 4/4、能力摘要及矩阵的 `--check` 均通过 |
| 本轮真实专项产物 | `/tmp/officekit-ppj-preview-5wyqgL`；测试完成真实 authored/source-bound codec+sharp 检查，仅为本机临时证据 |
| 未重新执行 | 全量 smoke、其他 fast/slow 分段、NativeAOT 构建、C# 专项、Skill 全套检查、宿主验收、人类校准、性能基准 |
| 本轮修改边界 | 仅更新本审计文档，不实现 G-01，不勾选其他变更任务，不代表已发布到 main 或安装包 |

G-11 实施复核记录（历史轮次）：

| G-11 实施复核项目 | 记录 |
| --- | --- |
| HEAD | 本轮从 `09f79e04` 开始，期间另一工作流提交 `0e1d8c8f`；G-11 为未提交工作区实现 |
| 已落地的输出修复提交 | `c216362a`，`feat(ppj): publish preview artifacts with verifiable output evidence` |
| Node.js | `v23.10.0` |
| `src/ppj/svg-preview.mjs` SHA-256 | `f98c59fb95478610748e4377b906c922156fef0e64898eb1659cad2d1d7986ff` |
| `src/ppj/preview-output.mjs` SHA-256 | `0d5aca1b4c89a5ed43e599250414dedc0312f7076cc928359112759d05f374df` |
| canonical fixture SHA-256 | `6b3e6c50a67da620494df75aac7a10bb343b0db88e6d97c9f803437b8e5ae048`，与初次相同 |
| 实际重跑 | gate-policy、slow/presentation、全量 preview smoke、能力摘要/矩阵 --check、JS syntax、Skill portability/reference-sync 和 strict OpenSpec；详细结果见第 6.1 节 |
| 真实专项产物 | `/tmp/officekit-ppj-preview-itVihO`，含 authored/source-bound 输出；只是本机临时定位，不是可移植证据包 |
| 工作区边界 | 本轮接通 G-11 运行时，增加测试并同步任务/文档；保留另一工作流的 C# 和 workbook 改动，不将其作为本轮独立验收 |
| 本轮未执行 | 完整 fast/slow、NativeAOT 重建、宿主 PowerPoint 验收、全量视觉人工校准和性能基准 |

初次审计快照（历史记录，不代表最新 HEAD）：

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

证据分为三类：**运行确认**表示相应审计轮次实际执行所得；**代码确认**表示当前实现可以直接确定的行为；**待验证**表示尚缺专门案例，不能推断成功或失败。G-11/G-12 各自说明已验证范围，其余编号的完成条件仍是后续要求。

本文术语：`authored` 指从结构化输入创建文稿；`source-bound` 指编辑仍绑定原始 PPTX、原生对象和所有权证据；`opaque` 指无法完整、安全建模的原生内容；`canonical JSON` 指规范化程序文本，不承诺已展开布局；`fixture` 指可重复运行的固定测试输入；`smoke` 仅检查基本路径能否跑通。

初次写作期间有其他工作流更新对数轴、registry 和能力矩阵，G-11 实施复核期间又有 errorBars 相关开发。本文保留测试发生时的观测，并在第 6、8 节注明复核变化；这些并行更新不算本审计完成的修复。

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
| 缺失数据不得当作 0 连线 | 普通 line 有有限拆段；登记的错误会降低可靠性 | 单点段仍被丢弃；若干图表仍通过 Number/num 将 null 转 0；显式 zero/connect 显示策略尚未统一到绘制与证据 | G-08～G-11 |
| 图表关系必须和数据拓扑一致 | 对已知比例、通道、层级、轴等错误有事实诊断 | 饼图比例、Sankey 边、树层级、OHLC 通道、主副轴等实际图形仍可能错误；诊断没有修复它们 | G-01、G-06、G-08、G-09 |
| source-bound / opaque 不得扁平化 | 复用编译流程与源身份；native 已从实际编辑候选采集，测试验证 opaque 原内容与非目标 ZIP 保留 | JS 还未消费候选场景；完整语义身份、第三方复杂输入、实际视觉与编辑保真仍缺全面验证 | G-01、G-07、G-12、G-14 |
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

### 2.3 G-01：预览没有拿到完整的已解析绘制结果

`renderPpjToSvg` 读取 `compiled.programJson`，而 authored 编译器返回的是 `validation.CanonicalJson`。编译器另外持有 expansion/build plan，返回的 JSON 并不等于已经展开、排好布局、解析完样式的绘制场景。

初次审计运行 canonical fixture 后，返回的 JSON 中仍有 **1 个 `component`**。当前预览端仍自行读取 label/value、安排位置；它没有复用该组件实际编译出来的元素树。

影响包括组件 repeat、dataset 编码、styleRef、主题和 grammar token 等高层表达：PPTX 编译可能处理正确，预览却绕过这些处理重新猜测显示。

来源：[workspace 编译转发](../src/ppj/workspace.mjs)、[authored 编译回执](../native/OfficeKit/src/OfficeKit.Codec/PpjAuthoredPresentationCompiler.cs)、[预览入口](../src/ppj/svg-preview.mjs)。

完成条件：明确一个共用的解析/布局输出边界，预览消费编译器同义的结果；组件和 dataset 的高层写法与等价展开写法应得到等价预览。

#### 已形成的方案与尚未实现的部分

现已建立 [G-01 提案](../openspec/changes/ppj-preview-compiler-scene/proposal.md)、[设计](../openspec/changes/ppj-preview-compiler-scene/design.md)、[规格](../openspec/changes/ppj-preview-compiler-scene/specs/ppj-preview-compiler-scene/spec.md) 和 [15 项任务](../openspec/changes/ppj-preview-compiler-scene/tasks.md)。1.1、1.2、2.1、2.2、2.3 已实施并勾选；其他 10 项未完成。wire 已有只读场景回执，authored 采集 writer 实际输入，source-bound 导入最终候选，节点保留真实 owner；JS 仍未转发/消费，旧组件 label/value 启发式仍在执行。

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

以上完成的是 C# 场景生产与身份层。原始 canonical JSON、原有 node map 和候选 PPTX 均有开关前后字节不变断言；旧 `#component(...)` expansion 路径没有被改写成新的可编辑契约。JS 实际消费、新场景对应的图片警示/清单绑定及重建 NativeAOT 待后续任务。G-11/G-12 已有的警示与发布清单继续有效，待做的是它们与新场景的对接。

方案复用 C# 已有 `PresentationArtifact`，不再定义一套作者语言，也不在 JS 中重新解析 OOXML。默认关闭的只读场景选项已加入 wire，`programJson` 保留既有含义。下表是整条方案及完成证据；协议与 collector 的基础已经完成，其余环节按上文区分在途与待实现：

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

## 3. 元素与页面视觉差距

下表对齐 [PPJ schema](../src/ppj/ppj-v1.schema.json) 中的 16 类元素。声明状态来自当前 [preview capabilities](../src/ppj/svg-preview-capabilities.json)，描述整个类型的保守边界；运行时还会检查实际字段与继承状态。当前没有任何整个元素或图表类型被声明为 supported；这不表示连一个简单原语也画不出来。

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

G-11 实施轮次记录本变更 9/9 项已完成，presentation 四步、gate-policy、生成摘要/矩阵、portability、reference-sync、链接和严格 OpenSpec 检查通过；同时真实 CLI 回归确认退出码 0 可与 failed 可靠性并存。最新文档复核仅重跑第 1.3 节列出的检查，不将这些历史记录重复算作新测。完整全仓测试、宿主验收和性能并未因此完成。

- 1.1、1.2、2.1、2.2：同源声明、诊断基础、实际字段遍历和字段限制。
- 2.3：每个已登记事实错误在绘制与发布后仍为失败，配色不能抵消；未知限制与已知矛盾分开。
- 3.1、3.2：真实绘制/警示/发布接通；嵌套重复局部 ID、失败保留、依赖惰性和源字节保留有断言。
- 4.1、4.2：常规门禁、相关文档/Skill reference 与严格 OpenSpec 检查；最终勾选以 [任务清单](../openspec/changes/ppj-preview-support-diagnostics/tasks.md) 为准。

真实 authored 样例成功发布，但为 partial/failed；简单 source-bound 样例成功发布，但为 opaque/requires-review。清单 `ok: true` 和 CLI 退出码 0 仍只说明发布成功，调用者必须读取可靠性。failed 不得作为正确性证据；requires-review/passed 也不取代人工视觉、结构和编辑保真检查。CLI 的整体交付策略仍属 G-13。

### 5.2 G-12：输出安全和失败处理

修复进度：下面列出的初始输出问题已有实现与专项回归。新增 `src/ppj/preview-output.mjs` 独占创建目录和文件，在失败时保留实际产物；清单记录 hash、字节数、原始页 ID、输入/源/候选/资产身份以及 Node/栅格后端信息。源码版本摘要也进入清单。SVG 和 PNG 仅在写入成功后列出，缺少 sharp 返回不完整失败而保留 SVG；pending/final 清单与错误码区分失败阶段。

前次修复验收中，`npm run test:ppj-preview-output`、`node test/ppj-svg-preview.mjs`、`node test/ppj-preview-capability-coverage.mjs` 均已通过；本轮通过 presentation 分段重新执行对应测试，并非另行逐条执行这些命令。真实 SVG/PNG 专项同时验证 authored 和原生重新投影后的 source-bound 输入，原始 PPJ/PPTX 字节保持不变；模拟故障测试覆盖目标冲突、并发、路径映射、资源变化窗口、页面栅格失败及清单写入失败。

集成修复进度：新轻量回归已加入 fast/slow gate，slow 的 presentation 分段同时执行诊断、输出发布、SVG 预览与覆盖声明检查。`node test/gate-policy.mjs` 校验当前 PPJ 入口、已退役入口不再出现及分段连续性。最新复核实际执行该 policy 和完整的 presentation 四步分段，均通过；未运行其他 fast/slow 分段。原始缺陷记录如下：

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

下表为 G-11 实施轮次的历史结果。最新文档复核重新执行的检查见第 1.3 节，G-01 基础实施证据见第 2.3 节；最新复核没有重新运行全量 smoke，因此下列 42/4/38/0 不属于本次重跑结果。

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
| 全量 preview smoke | 退出码 0；42 输入，4 rendered、38 compilerRejected、0 rendererFailed | G-11 实施轮次重现前置拒绝分布；真正进入绘制的四项均为 partial，诊断数见下文；不代表视觉通过 |

G-11 实施轮次的全量 smoke 中，minimum/canonical/aqua-impact-story/simple-dark-mode 的诊断数分别为 44/686/3674/1032，四项支持状态均为 partial。这表明检查已经进入实际绘制，也提示继承诊断量需要在 G-16 测量和优化；数量多不等于语义覆盖完整。以下初次运行表仍保留为历史记录，不能与该轮诊断数混用。

初次审计结果（保留原测试能力描述；真实专项测试后来增加了源绑定与输出证据断言）：

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

上表和第 3.1 节 SVG 是此前工具输出的审计摘录，不是本轮重新测量值。完整原始 stdout、临时 PNG/SVG/render.json 没有归档为仓库固定证据；本文不声称这些精确运行结果已具备跨环境复现包。后续正式验收应保存原始记录，而不仅保留汇总表。

G-11 实施轮次的通用 Skill 模板校验器另报告已有 `name: "Presentations"` 不符合其小写命名规则；该轮只改交付 reference，未改名或变更路由。该轮仓库专用 portability/reference-sync 已通过，此工具兼容性差异不冒充通用校验成功；本次文档复核未重跑它们。

### 6.2 G-14：测试覆盖不足

当前预览测试已经增加一个有限图片映射的“画对”断言，并在真实绘制和发布中识别已登记的事实错误、核对警示像素，但远未覆盖全部绘制。饼图只有圆、形状缺失、双轴错误仍存在；自动门槛会失败，但图片存在性或发布成功本身仍不能证明画对。

覆盖测试的变化与剩余问题：

- 初次图表名称靠硬编码和字符串匹配；现在按实际 schema 对齐 16 类元素和 16 类图表，并校验 registry 指针和 owner。全量视觉字段到具体行为断言的覆盖仍未完成。
- 初次只检查等级合并后的 Set；现在校验等级互斥、非法提升为 supported 和生成摘要漂移。规则中测试文件存在不等于该测试证明了整个类型正确。
- 初次只扫顶层；现在递归扫描 fixture 对象树，并有嵌套 group/未知类型反例。独立输入检查器也测试 component 定义、slot、重复局部 ID；真实布局展开与逐对象显示尚未验证。
- 当前 `test/fixtures` 下只有一个 `.ppj` 文件；这不代表仓库没有其他测试生成 PPJ，只是该扫描并未检查它们。
- 初次审计时专项脚本未进常规 gate。当前已修复：发布故障测试进入 fast/slow，真实 SVG/PNG 与声明覆盖进入 slow/presentation；全量仓库扫描 smoke 仍是独立命令，未设正向渲染数量门槛。

smoke 还通过错误字符串正则区分 compilerRejected/rendererFailed，而非实际失败阶段；且仅当 rendererFailed 非空才失败。即使所有输入都在编译阶段被拒绝，也可能返回成功退出码。

另一个待补回归来自输出保护的新边界：smoke 用成功数量 `report.rendered.length` 命名下一例目录。如果某例在取得目录后发布失败并保留该目录，下一例仍使用同一编号，会遇到 `preview.output.exists`，干扰后续故障分类。此为当前代码可以推导的条件路径，本轮未注入该故障；后续应按输入序号或唯一案例 ID 分配目录，并验证单例失败不污染后续案例。

完成条件：结构、数据语义、视觉、编辑保真分别有断言；正例必须确实进入渲染；预期拒绝独立列负例，不得用它们抵消正例失败。事实错误设硬门槛，不能被外观评分抵消。

### 6.3 G-15：能力台账尚未防止后续漂移

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

G-11 的共享检查、绘制与发布已接通；后续优先明确 G-01 的共用解析边界，再逐项修复实际绘制。P0 指会造成误判或破坏证据的问题，P1 指完成静态功能覆盖，P2 指性能与交付完善；P2 不等于最终目标可以不做。

表内 G-12 已完成的限定契约作为后续回归基线保留，不再列为待实现；大文稿、压力条件下的失败恢复由 G-16 继续验证。

| 阶段 | 关联缺口 | 交付结果 |
| --- | --- | --- |
| P0：先停止误报 | G-02、G-06、G-09～G-14 | 待修事实表达与 smoke 失败分类；G-11 准确报告、G-12 不覆盖与真实清单已完成，持续回归 |
| P0：统一语义输入 | G-01、G-04、G-08、G-10 | 明确复用 compiler expansion/样式/数据归一化的边界，避免继续增加第二套猜测逻辑 |
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
- [ ] 字段级映射和生成文档检查进入维护 gate，新功能不会静默落后于渲染器。
- [ ] 本地性能、依赖缺失和失败恢复完成实测，达到另行确定的使用目标。

### 7.3 后续开发如何与功能更新同步

下一项按已有 G-01 方案完成 3.1 的 JS 场景传输与校验，再接通实际绘制；目前清单勾选 5/15，writer/candidate 采集及原生归属已验证，G-01 整项仍未完成。G-11 是防误判措施，不能替代后续实现。对某个具体字段可以直接复用现有结果的，不必等待一套大型新框架。

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

## 8. 与接口补齐、Skill 实验的关系

### 8.1 接口已继续更新，不能混算渲染完成

初次审计时 `ppj-chart-axis-log-base` 尚在推进，清单曾从 0/4 变为 3/4。最新复核中，对数轴实现已提交为 `9efc8ce5`，当前任务清单为 5/5 勾选；趋势线 source-bound 列表编辑也已提交为 `fbddc7e7`。本节不再把它们笼统记录为“仅有在途文件”。

本轮没有重新执行这些 C# 功能的 NativeAOT 构建和专项验收，也未核验清单勾选对应的全部原始日志。“已提交/清单已勾选”是仓库事实，“已由本次独立验收”则不是。当前 preview 仍未读取 `logBase`、`trendlines` 或 `errorBars` 来绘制对应内容。

此前收尾时新增的 errorBars 预览诊断目前已在实际绘制代码中：普通图表兜底分支对存在 errorBars 的系列增加 `chart-error-bars-not-rendered` 的 partial 诊断，最新 presentation 分段也包含相应回归。这是诊断补充，不是误差线绘制实现；提前返回的其他图表分支仍不能据此宣称全部覆盖。

G-11 实施轮次从 `09f79e04`（literal custom error bar data）开始，另一工作流随后提交 `0e1d8c8f`（固定公式误差数据与内嵌 workbook 同步）；之前还有 `ccb93795`（source-bound scalar error bar lifecycle）。这些是可见提交事实，不等于本审计的 C# 独立验收。预览尚未绘制这些误差数据，不能用接口提交代替绘制完成。该轮随后又出现 `ppj-chart-trendline-label` 在途工作；生成矩阵的漂移检查捕获其新增 schema，已重新生成并通过检查，但没有将标签接口或绘制记为已验收。

上述对数轴变更的范围是 value axis 的 logBase 创建、读取、修改和删除，含 numeric X 与 combo 副轴、token、非法值及所有权保护。具体跟踪见 [对数轴任务](../openspec/changes/ppj-chart-axis-log-base/tasks.md)。

该接口即使通过 codec 回归，当前 preview 也未消费其缩放语义。它是“接口更新后渲染需要同步”的具体例子，不应以接口完成代替渲染完成。本文不修改或验收其他工作流正在编辑的文件。

前次文档复核补记：趋势线标签实现已提交为 `93b89e67`，不再仅是 G-11 实施轮次观察到的在途工作。预览与发布器源码 hash 没有变化；此提交不代表趋势线标签已在自有 SVG 中绘制。该次复核没有重跑该接口的 C# 专项或构建对应 NativeAOT。

前次实施快照补记：起点 HEAD 已包含 `f3a67617` 的趋势线标签手动布局及 `3eb58b3a` 的数字格式链接保留，当时工作区另有趋势线富文本改动；收尾时该功能已由另一工作流提交为 `3b6272dd`。它们继续扩大接口与预览的同步检查范围；本文只核对提交和差异归属，不将其算作已完成的趋势线视觉渲染或该次 C# 验收。

本次读取基线还包含 `307106dc` 的 chart text language 变更，以及 `ade46bd5` 对并行 preview bindings 归属的调整。当前 renderer 源码未变；图表文字语言、趋势线标签富文本、布局和格式是否进入最终画面仍要单独验证。不能因为这些字段已加入接口或生成矩阵，就认定渲染器已同步支持。

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

整段复核，从仓库根目录运行：

```sh
node test/gate-policy.mjs
npm run test:slow -- --segment presentation
node scripts/generate-ppj-preview-capabilities.mjs --check
node scripts/generate-presentation-capability-matrix.mjs --check
node test/ppj-preview-smoke.mjs
```

单项定位时，可分别运行 `node test/ppj-preview-diagnostics.mjs`、`node test/ppj-svg-preview.mjs`、`node test/ppj-preview-capability-coverage.mjs` 或 `node test/ppj-preview-output-evidence.mjs`；这四项已包含在 presentation 分段中。此处列的是可复跑命令，本次文档复核的执行范围见第 1.3 节，历史结果见第 6.1 节；不要把列出全量 smoke 命令当作已重跑。

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

测试数量可能随代码推进增加，以本次命令实际输出为准。NuGet 恢复或 SDK 失败应记为环境/构建失败，不能写成目标断言失败或通过。原生专项通过之后，仍需按 G-01 任务 5.2 重建并验证 JavaScript 真正加载的 NativeAOT，才有新场景端到端的运行时证据。

### 9.2 主要代码与文档

| 文件 | 核对用途 |
| --- | --- |
| [svg-preview.mjs](../src/ppj/svg-preview.mjs) | `drawElement` 绘制、`assessedDrawing` 检查合并及页面警示 |
| [preview-output.mjs](../src/ppj/preview-output.mjs) | G-12 非覆盖发布、失败状态、文件与输入身份清单 |
| [svg-preview-capabilities.json](../src/ppj/svg-preview-capabilities.json) | 当前 registry 派生的类型级保守声明 |
| [preview-capabilities.mjs](../src/ppj/preview-capabilities.mjs) | registry/schema 对齐、声明生成与漂移检查 |
| [preview-diagnostics.mjs](../src/ppj/preview-diagnostics.mjs) | 绘制和发布共用的诊断结构与可靠性聚合 |
| [preview-input-assessment.mjs](../src/ppj/preview-input-assessment.mjs) | 实际字段、继承、metadata 与源载荷边界检查 |
| [preview-factual-errors.mjs](../src/ppj/preview-factual-errors.mjs) | 当前绘制缺陷对应的错误级规则；不是绘制修复 |
| [G-11 任务清单](../openspec/changes/ppj-preview-support-diagnostics/tasks.md) | 9/9 任务、最终门禁和明确保留的其他 gap |
| [G-01 设计](../openspec/changes/ppj-preview-compiler-scene/design.md) | 共用 native 场景、实际候选导入、身份及预算的方案；不能用规划代替实现证据 |
| [G-01 任务清单](../openspec/changes/ppj-preview-compiler-scene/tasks.md) | 5/15 已勾选；JS 消费、真实 SVG 等价和所用运行时仍待验收 |
| [场景 collector](../native/OfficeKit/src/OfficeKit.Codec/PpjPreviewSceneBuilder.cs) | native 视觉字段复制、预算、摘要和载荷隔离 |
| [writer observer](../native/OfficeKit/src/OfficeKit.Codec/PpjPreviewSourceFreeBuildPlan.cs) | authored 同次 materialization 采集及节点绑定 |
| [authored 原始归属](../native/OfficeKit/src/OfficeKit.Codec/PpjPreviewOrigins.cs) | 现有 expansion/slot 替换中的只读 origin 跟踪，不改 JSON 和旧 node map |
| [authored 场景测试](../native/OfficeKit/tests/OfficeKit.Codec.Tests/PpjPreviewAuthoredSceneTests.cs) | 12 个实际编译、上下文、图表、writer 生命周期及嵌套/repeat/slot 归属案例 |
| [candidate 场景生产](../native/OfficeKit/src/OfficeKit.Codec/PpjPreviewCandidateScene.cs) | 实际候选 native 导入、资产复用、确定归属及未知身份的保守处理 |
| [candidate 身份匹配](../native/OfficeKit/src/OfficeKit.Codec/PpjPreviewCandidateBindings.cs) | 在途物理 part/对象身份连接；重复和缺失身份不授予归属 |
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

### 2026-09-09 提交快照复核

基于 `654cb3e5` 的独立提交快照已验证：原生预览专项 39/39 通过；JS 预览诊断及新增场景传输测试、`proto:check`、gate-policy、OpenSpec 严格校验和差异格式检查通过。此次包含原生归属实现和 JS 默认关闭的场景传输、身份与资产校验。传输实验使用合成回执和替代原生调用，只证明传输契约；尚未重建 NativeAOT，也未验证新场景的 SVG 绘制，任务 3.1 与 G-01 保持开放。以上是本次提交检查，前文“未提交”等描述保留为各轮历史快照。
