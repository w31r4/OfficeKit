# OfficeKit 本地 PPT 渲染器：当前能力与剩余差距

核对日期：2026-09-10。本次文档复核读取的 HEAD：`cd733a8e2e51b68e6efa236f67d83cde0ca1946c`。工作区另有未提交的原生、渲染器、测试、能力声明及文档修改，HEAD 不能单独标识这些修改。本文区分所读源码、留存运行结果与尚未验证的变更，不代表安装包或远端分支状态。

本文记录现有 PPT 静态功能的已知差距，不是完整字段覆盖率报告：目前缺少逐字段分母，不能据此声称已穷尽所有 OOXML 属性。本轮只核对并整理文档，没有重新构建原生包、执行完整集成、全仓测试、外部 Office 或人类验收。各历史运行仍按自己的版本和范围解释。

### 本次复核摘要

本次重新读取正式入口、内部 painter 的检查开关、OpenSpec 勾选及留存的 `tmp/officekit-native-scene-paint-PgPOma/integration.json`。报告时间为北京时间 2026-09-10 07:04:09；以下以这份报告汇总集成证据，不将其称为当前工作区全量验收：

| 核对项 | 结果 | 对完成判断的含义 |
| --- | --- | --- |
| 正式预览入口 | 仍读取 `compiled.programJson`，编译请求未启用 preview scene | 内部 painter 的进展尚未默认交付给 CLI 用户 |
| 内部输入可靠性检查 | `paintPpjSceneSvg` 的 `assessInput` 默认仍为 false，可显式启用 | 有合并检查的实现和限定证据，尚不是不可绕过的正式入口保证 |
| 成对 fixture | 嵌套组件/repeat/slot、命名样式/grammar、dataset/encoding、native line 与 vector heatmap 的固定对照通过 | 任务 5.1 指定类别已有证据；不代表所有字段、图表变体、workbook 或第三方源均完成 |
| 文字近期进展 | 小字号基线及文本框右边距的限定像素断言通过 | 完整换行、字体度量、垂直布局和 AutoFit 仍缺 |
| 图片 tile | 作者/源 no-op 的明确拒绝、占位和输入保留断言通过 | 通过的是防误画测试，不是平铺渲染成功 |
| 源变换字段删除 | rotation、flipH、flipV、组合四例失败 | 删除仍留下显式 0/false；画面恢复原位不能抵消编辑语义失败 |
| 报告整体状态 | `failed`；上述四例保留在 `transformProfileFailures` | 局部对照成功不能报告为整套通过 |
| G-01 任务 | 10 项勾选、5 项未勾选，共 15 项 | 只表示任务记录，不是渲染覆盖率或总体完成率 |
| 性能证据 | 重复组件/源候选有分阶段采样、独立进程驻留高水位及预算失败恢复记录 | 保留峰值与完整负载目标仍缺，5.3 未完成；托管释放证据另见 G-16 |

该报告使用的原生包是固定 `4b7cd9c6` 源码加两项 AOT 修复，包摘要见 G-14；并非本次 HEAD 及工作区所有修改的构建。本次核对性能文件实际 SHA-256 为 `b99d5ffbffc3c1722925701cc2ff95041d077140e9af0c2464afd4060503b320`，与集成报告记录一致。报告中的 SmartArt 源绘制是明确的 unavailable 反例，不因失败列表为空就变成视觉保真成功。历史成功报告保留为对应版本的证据，不能覆盖新增的失败断言。

## 1. 结论与阅读范围

**渲染器尚未整体完成。内部场景绘制已有真实成功案例，但正式 CLI 仍使用旧绘制路线。** 当前不能把“生成了 PNG”“测试通过”或“原生接口能保留某个字段”解释成现有功能已全面、正确地显示。

主要差距分为五组：

1. 内部场景绘制尚未接入正式 CLI 的检查、发布和证据链。
2. 文字排版、主题继承、效果及部分静态元素仍缺实际绘制。
3. 图表只有限定类型和变体具备内部语义回归，复杂坐标、关系及公共样式仍不完整。
4. 已有 source-bound 测试主要来自本项目生成的简单文稿，不能代表第三方复杂 PPTX。
5. 字段级防漂移、后续版本持续复验、性能基准和人类视觉校准尚未闭合；一个固定版本的构建验收已通过。

本文面向维护者和后续开发者，目标仍是**本地个人使用的 Skill 与 CLI 运行时**，不引入前后端系统，不要求 PowerPoint 像素级复刻。范围是 PPT 静态页面预览；Word、Excel、PDF、Live 宿主、媒体播放和动画执行不计入静态渲染完成度。

本文沿用[详细审计](ppj-svg-preview-gap-audit.zh-CN.md)的 G-01～G-16 编号。旧审计包含实施日志和历史反例，正文中部分“当前”属于较早快照；判断本次现状以本文明确标出的源码、运行时和证据范围为准。[历史附录](ppj-svg-preview-gap-audit-history.zh-CN.md)继续用于追溯。

### 1.1 几个术语

- **正式路线**：用户当前调用 `officekit ppj preview` 实际执行的代码。
- **内部路线**：已编写、可单独测试，但尚未接入正式预览发布的场景 painter。
- **native scene**：编译器提供的原生绘制场景，包含有效对象状态及来源归属；不是把 PPJ 再解释一遍。
- **source-bound**：编辑与原始 PPTX 中的对象和部件绑定，必须保留非目标内容。
- **opaque**：无法安全解释或编辑的原生内容。可展示可信源预览或明确占位，不得猜测其结构后重写。
- **重新投影**：将编辑后 PPTX 再导入为模型，检查实际写入结果，而非只检查内存中的编辑请求。

### 1.2 差距总表

状态按行为和交付范围判断。“部分实现”指存在可用子集；“待验收”指有代码或产物但缺指定回归；“未闭合”指整个交付条件尚未满足。下表不是新增的能力声明，也不按行数计算完成率。

| 编号 | 当前差距及状态 | 对用户的影响 | 关闭前必须补的证据 |
| --- | --- | --- | --- |
| G-01 | 场景路线部分实现，正式入口未接入 | 内部改进尚未成为日常 CLI 输出 | 真实 CLI 场景消费、检查与发布整链回归 |
| G-02 | 文字布局、几何求值部分实现 | 长文本可能错行或越界，复杂形状可能失真 | 换行/继承/删除、非矩形与复杂路径行为断言 |
| G-03 | AOT 投影已修复；显式归零通过，源字段删除仍失败 | 删除变换实际留下 0/false；复杂嵌套等也未完整验收 | 真正移除原生属性且保留显式零；嵌套/导入组合及正式接入 |
| G-04 | 主题、继承和效果未完整消费 | 色彩、可见范围和对比度可能误导检查 | 有效样式、覆盖/删除、透明叠加和效果边界 |
| G-05 | 图片已有裁切/有限遮罩；tile 与效果等仍缺 | 主体或透明边缘可能显示不完整 | 裁切/变换组合、解码失败及不确定范围诊断 |
| G-06 | 表格与连接线部分实现 | 合并边界、锚点或连线关系可能不准确 | 合并边框、目标移动/解绑、端点与源保留 |
| G-07 | 简单源生命周期有证据，SmartArt 源视觉不可用 | 语义编辑成功不能证明候选外观保真 | 实际目标缓存、关系、直接样式及非目标内容保留 |
| G-08 | 图表公共语义不完整 | 轴、单位、标签和尺度可能改变事实理解 | 双轴、数值格式、标题/标签与有效样式 |
| G-09 | 全部 chartType 仍为 partial | 能生成形状不等于正确表达该类数据 | 每类有效正例、复杂拓扑和实际图形断言 |
| G-10 | 缺失值只有限定类型回归 | 缺口、真实零、堆叠分母可能被混淆 | 各通道缺失、孤立观测及显式显示策略 |
| G-11 | 旧路线诊断已验收，新 profile 接入未闭合 | 内部通过不能使未检查节点或字段自动通过 | 逐规则证明、节点归属、未知字段与继承失败 |
| G-12 | 场景发布契约已验收，正式接入未完成 | 内部清单证据尚非默认 CLI 交付保证 | 接入后保留身份、失败状态及非覆盖发布回归 |
| G-13 | 页面选择与完整交付流程未闭合 | 用户难以选择页、辨认三类检查的完成情况 | 页过滤/编号、隐藏页策略、分项结果和失败重跑 |
| G-14 | 外部源与逐字段测试不足 | 自生成简单文稿的成功难以外推 | 第三方 PPTX、workbook、正反例分开统计 |
| G-15 | 防漂移有基础，行为覆盖未闭合 | 模型新增字段可能只传输而不实际显示 | 字段到绘制/诊断/测试的逐项对应与生成检查 |
| G-16 | 性能与稳定性没有完整基线 | 不能承诺大文稿耗时、内存及失败恢复 | 预先约定负载目标，测量分阶段耗时和资源峰值 |

优先关注会改变事实的差距：SmartArt 目标缓存损失、错误数据通道/轴/拓扑、缺失值以及误导性的通过状态。源组投影故障已有有界修复回归，仍需保留为后续版本的防退化案例。主题和文字问题也可能遮住数据或改变关系理解，不能一律当成装饰问题。各项实现位置、限定成功和剩余工作见第 4 节；执行顺序见第 6 节。

### 1.3 如何判断“成功”和“还差什么”

每项功能要沿输入、编译/导入、场景传输、实际绘制、检查发布五层核对。一个字段能被 C# 写入或通过 wire 传输，不表示 JS 已消费；出现 SVG 节点，也不表示位置、数据关系和效果正确。

| 证据等级 | 可以证明 | 不能代替 |
| --- | --- | --- |
| 源码与能力声明 | 存在入口、字段、分支或保守限制 | 真实运行与字段视觉效果 |
| 合成 scene 专项 | 给定原生状态的局部绘制和诊断行为 | 编译器确实生成同样状态、源文件确实保留语义 |
| 指定 NativeAOT 集成 | 指定二进制下输入到实际 SVG/PNG、候选及重新投影的限定行为 | 当前工作区新 C#、正式 CLI、第三方样本和完整字段覆盖 |
| 正式 CLI 端到端 | 用户入口实际采用该路线，返回值、文件和失败状态一致 | Office 宿主兼容性或人类视觉可用性 |
| 外部文件与人类检查 | 指定第三方样本、宿主或评审标准下的表现 | 任意文稿、交互/动画行为、未测字段 |

“预期拒绝测试通过”表示危险输入没有被伪装成正确结果；该输入的渲染能力仍未完成。“待验收”也不等于已知实现错误，可能只是当前版本没有足够证据。后续修复应分别减少已知错误和证据缺口，不通过删断言或降低门槛制造完成率。

## 2. 两条路线：为什么内部成功不等于用户已经可用

| 环节 | 正式 CLI | 内部场景路线 |
| --- | --- | --- |
| 输入与编译 | 加载 workspace、调用 PPJ 编译器 | 由编译接口请求并传输 `PresentationPreviewScene` |
| 绘制依据 | 解析 `compiled.programJson`，自行读取元素和数据 | 消费 scene view 中的原生状态、坐标、数据及资产 |
| 主要实现 | `src/ppj/svg-preview.mjs` | `src/ppj/preview-scene-svg.mjs` |
| 检查与发布 | 已有输入/绘制诊断、SVG/PNG、`render.json` 和非覆盖保护 | 场景身份、作者/源候选发布与发布故障矩阵已验收；正式接入仍缺，显式标记 integration pending |
| 当前意义 | 能运行，但保留旧语义和几何缺陷 | 证明有限功能可按原生状态正确绘制，尚非正式交付入口 |

正式 `renderPpjToSvg()` 目前调用编译器时只传 `includeNodeMap: false`，随后读取 `compiled.programJson`，没有调用内部 `paintPpjSceneSvg()`。因此，最近图片、表格、折线图、比例扇区、柱形堆叠和折点的内部修复，不能计为正式 CLI 已修复。

后续应继续复用编译器的组件展开、数据映射、布局和对象关系结果。JS painter 负责显示这些结果；不应另建一套 PPJ 数据集、样式或 OOXML 解析器。SVG 是绘制输出，PNG 由栅格后端生成；这一职责划分不需要另做一个 Office 宿主。

## 3. 已经成功的范围，以及证据到底证明什么

以下构建和实施记录均为此前留下的证据；段落中的“该轮/本轮”按紧邻的报告或提交解释，不指本次文档整理。当前汇总结论见开头，不将不同版本的局部成功拼成一次全量验收。

### 3.1 当前源码和测试能够支持的结论

| 功能 | 内部路线已实现或已有通过证据 | 不能由此推导的结论 |
| --- | --- | --- |
| 基础几何 | 有限预设、literal 自定义路径、二次/三次 Bézier、`arcTo`；形状与文字分别绘制 | 所有 preset、公式引用、连接点、文字区域均已求值和显示 |
| 圆角矩形 | 形状和图片遮罩共享 roundRect 逻辑；默认值、零值、非正方形和调整边界有合成断言 | 所有形状 adjustment 都有对应实现 |
| 文字 | 保留段落/run 边界、显式换行、有限对齐、字号、字体、粗斜体、直接颜色/透明度；单下划线、单删除线、有符号基线及零覆盖 | 完整字体度量、自动换行、列表、AutoFit、双线/波浪装饰或全部文字效果 |
| 组与变换 | 原生 frame/childFrame 映射、旋转、镜像、元素 hidden 分支 | 所有嵌套、导入变换及效果扩展边界均已验收 |
| 图片 | 正裁切、负边值透明留白、alpha；rect/ellipse/diamond/roundRect 遮罩及 literal 自定义遮罩；直接 RGB 边框 | tile、全部 fit/focus、主题色、阴影及复杂路径都已支持 |
| 表格 | 显式行列尺寸、合并范围、有限单元格样式；有合并像素和源移动回归 | 完整表格主题、banding、边框冲突和文字排版 |
| 连接线 | 原生端点、straight/elbow、箭头方向及相对尺寸、literal bendAdjustment；对象锚点两个集成案例 | 任意 connection-site、曲线、自动避障或导入路线精确复刻 |
| 普通 line | 缺失断段、真实零、孤立观测提示、有限 marker、显式范围/反向和有限对数轴 | 平滑/堆叠折线、全部轴、图例、标签和源主题 |
| pie/doughnut | 单系列比例、旋转、孔径、逐点颜色和径向分离；源修改/删除回归 | 多环、负值、完整标签随动或 Office 精确分离间距 |
| column/bar | 普通分组、正负分开累计、非负百分比堆叠、缺失与零总量提示；源编辑及像素回归 | 负值百分比、对数基线、混合/副轴及全部样式 |
| scatter | 每系列数值 X/Y、有限 marker、缺失与零区分；作者 X 变化及源 Y 编辑有像素回归 | 完整连接线、平滑、源 X 编辑、气泡 size 或完整轴/样式 |
| 源编辑 | 限定 fixture 的 no-op、目标修改、原源不变、非目标 ZIP 成员保留、重新投影 | 任意第三方 PPTX、内嵌 workbook 或跨图表 family 编辑安全 |

这里的“有实现”“合成断言通过”“真实原生集成通过”是不同证据等级。表格记录的是各项已有的限定证据，不表示每个字段都同时具备全部等级。

例如路径的多种指令有合成断言，留存集成报告另记录了 customArcPath 和 6 条 generatedBezierPaths；这不等于每一种指令都具备第三方源输入、编辑、删除和视觉对照的完整回归。

### 3.2 初次文档核验与旧包报告

初次文档核验重新核对了正式入口、内部 painter、任务勾选、留存报告和两份 JS 文件 SHA-256，并执行 `node test/ppj-preview-scene-svg.mjs`，退出 0。这是此前核验记录，不是本次文档修订重跑结果。该专项使用合成原生场景，不证明最新 C# 已构建或生产入口已切换。

初次核验使用的完整留存报告为 `tmp/officekit-native-scene-paint-2MsoJ4/integration.json`，记录时间 `2026-09-09T20:31:55.736Z`（北京时间 9 月 10 日 04:31）。报告为 passed，diagramFailures 和 relationFailures 均为空；测试包含图片、文字、段落、形状轮廓、表格、连接线和限定图表的内部绘制及源编辑回归。SmartArt 的通过范围单独见 G-07。

| 验证身份 | 报告记录 |
| --- | --- |
| 内部 painter SHA-256 | `ecc3ed194d5b3baa52b13f6d2f8b3b3677b487167d8e5097b70789d307d3fe12` |
| 原生集成脚本 SHA-256 | `875f6173199fd9bfe3761ef3927f9fe627928572afe9d53dc97397f70a30f4c4` |
| PPJ 二进制 SHA-256 | `d43dcf4ce8000aa50165277aabfa0ed906680184aa2797a16e9f6c6e13b37c17` |
| Office 二进制 SHA-256 | `22a1c33c67e3ed15fbf2f3224013c0f0005a053404ba1819b3e9d67b7c8f2703` |

文档初次核对时两份 JS 摘要与报告一致；后续散点实现已经改变 JS，当前不能再用这两项摘要代表最新源码。运行时仍来自 `tmp/officekit-preview-runtime-495daafc`，不是当前 HEAD 重建包；匹配 JS 身份不能补足 C# 版本差距。新几何接口、文字属性删除等后续原生变更仍需稳定源码重建和专项验收。报告未冻结全部工作区差异、字体及环境，不能宣称完整环境可复现。

图表数据为 literal fixture，没有内嵌 workbook；源文件主要由本项目生成后去掉私有快照，不属于独立第三方样本。报告明确排除 production scene routing 和 complete paint coverage。测试整体 passed 包含“预期不可用必须明确失败”的反例，并非每个输入都可正确显示。

初次文档核对没有重新运行 NativeAOT 集成；后续散点实施已复跑指定旧包，结果见 G-08～G-10。全仓测试、双构建可复现检查、外部 Office、人类校准和性能基准仍未在这些轮次验收。历史 `4 rendered / 38 compilerRejected / 0 rendererFailed` 只代表旧 smoke：38 个输入未进入绘制，不能算作 42 个渲染成功。临时文件只作为本机追溯入口，清理后须重新生成；它们不是可移植 fixture 或运行依赖。

### 3.3 冻结源码重建验收（2026-09-10）

此前构建轮次从提交 `4b7cd9c6c56cd9af97b2e385428dc491b8aafc86` 导出独立快照，逐一核对 3731 个受跟踪普通文件/符号链接的 Git blob，全部匹配。legacy 下的参考库 submodule 未展开，不参与该次原生构建。快照不包含之后的原生文字属性修改；该轮 JS painter/测试包含散点实现，单独以报告摘要标识，不能把两者混称为同一提交的完整工作树。

构建使用仓库 `scripts/build-office-kit.mjs`，SDK 8.0.128，linux-x64，包版本 2.0.0。独立输出为 `tmp/preview-runtime-4b7cd9c6-ksr777`，7 个生成文件共 106022896 字节，没有替换已安装包。

| 证据 | SHA-256 / 结果 |
| --- | --- |
| PPJ NativeAOT | `86dd41600b3aa403a4d3111670ab6458b5ebb09395430deac2c3d615ab24efe2` |
| Office NativeAOT | `dcf53e562b33ceb668cbf1b6d0c2f206835c509580db19752bc04f139239064c` |
| manifest.json | `a297f27df7ef87bfff912e9460397b5983ff112442b8eac3d59399e5389e4c57` |
| 新包真实 JS 集成 | `tmp/officekit-native-scene-paint-NeJHpQ/integration.json`，passed，含 manifest 不变断言的最终复跑 |
| 原生 scene 专项 | 冻结快照中 `FullyQualifiedName~PpjPreview`，39/39、0 跳过 |
| 原生专项文件 | `tmp/preview-build-4b7cd9c6-ksr777/tmp/test-results/preview-scene.trx` |
| 协议检查 | 该轮工作区 `npm run proto:check` 退出 0；当时生成绑定与冻结提交 SHA-256 相同，非本次重跑结果 |
| presentation / gate-policy / 语法 | 4/4 及专项检查通过，保留 G-11/G-12 检查 |
| 双构建一致性 | 冻结快照中 `npm run verify:office-kit-build` 退出 0，两次构建的 9 个包文件一致 |

集成报告新增 manifest/SDK/包版本和五份关键 JS/生成绑定摘要，并断言运行前后文件摘要一致、实际二进制仍与 manifest 一致。它不是全环境锁文件，不涵盖每个间接依赖、字体或 OS。报告明确包含预期失败反例，SmartArt 源缓存与 scatter 连接模式仍未完成；其整体 passed 不意味着所有页面可靠性 passed。

首次发布在 MSBuild 编译阶段异常终止；独立诊断重跑得到退出 135/SIGBUS，系统 `/tmp` 同时已满。设置仓库 tmp 为 TMPDIR 并关闭共享编译/节点复用后，同一快照构建成功。使用的环境设置为 `UseSharedCompilation=false`、`MSBUILDDISABLENODEREUSE=1`，PATH 优先指向指定 SDK。没有清空系统临时目录、改 SDK 锁定版本或消除已有 nullable/trim/AOT 警告。成功只证明该环境组合可构建，不等于确认了 SIGBUS 的唯一根因。

这一轮完成 G-01 任务 5.2，更新的是运行时验证基线，不能替代新几何字段各自的绘制正例、生产 scene 接入、第三方源/内嵌 workbook、人类校准或性能验收。C# 快照和当前 JS 是报告明确区分的两个来源，不代表之后的整个 main 或所有未提交工作已经验收。

### 3.4 内部诊断与发布的最近证据

本节按实施先后记录发布与诊断证据，后续源组失败另见 G-03。多页归属修复报告 `tmp/officekit-native-scene-paint-OFGUGQ/integration.json` 记录于 `2026-09-09T21:14:17.834Z`（北京时间 9 月 10 日 05:14），状态为 passed，运行时仍为第 3.3 节的固定原生包。该轮修复后重新执行了完整 JS 原生集成；报告记录七份 JS/生成绑定的 SHA-256，不把匹配范围扩大到整个工作树，也不表示重新构建了当前 C#。

这次报告增加两个作者案例的真实发布检查：诊断地址保留、返回值与落盘 JSON 相同、输出文件实际摘要正确、重复发布拒绝覆盖。内部 painter 现已生成页面和全局 assessment；诊断去重使用独立 `scenePath`，避免同一个语义组件生成的不同原生节点或不同字段互相覆盖。全局未知字段也进入页级诊断。

发布失败的多页归属另有故障注入回归。旧实现仅按语义 `path` 回填页面：两页均为 `$` 时，第一页失败会被第二页 assessment 替换，文档检查树丢失第一页身份（顶层失败仍保留，并非整个结果变绿）。现按 `pageId/id/path/scenePath` 匹配，局部失败保留页地址，全局失败保留文档地址。首/末页单独栅格失败、两页都失败、全局依赖失败和 PNG 写入 ENOSPC 均通过，断言每页与文档子项、返回与落盘一致，其他页和输入 assessment 不被改写。注入 PNG 不作为真实像素证据；上面的 NativeAOT 集成用于确认既有真实绘制/发布未退化。

场景身份绑定随后已实现，最终报告为 `tmp/officekit-native-scene-paint-dBwhGX/integration.json`（`2026-09-09T21:19:38.058Z`，passed）。清单记录 version/origin/digest、program/candidate 摘要，绘制和发布共用 receipt 验证器。两个作者输入、源 no-op 和源文字编辑的真实发布均验证返回/落盘一致、实际文件 hash 和 PNG 顶部警示色；源 no-op 的候选 hash 等于原源，文字编辑后不同。旧场景配新候选、损坏场景摘要、缺失资产明确拒绝，身份失败先于创建目录和栅格加载。首次真实调用因 JS mimeType 与 wire contentType 差异失败，已抽取 scene view 原有转换作为共享验证入口后复跑，未放宽资产检查。

发布故障矩阵随后通过，最终报告为 `tmp/officekit-native-scene-paint-yuV8XP/integration.json`（`2026-09-09T21:23:50.212Z`，passed）。四个真实输入分别注入依赖加载、首张栅格、首个 SVG/PNG 写入、pending/final 清单写入失败，共 24 例；逐一验证场景和源/候选摘要、页面及文档状态、未受影响页、实际文件摘要、剩余 PNG 可解码、清单生命周期和输入不变。另对真实作者散点连接失败与源 SmartArt 缓存失败发布红色警示，核对 PNG 像素和失败清单。场景缺失/版本错误及缺少资产在真实 receipt 上被拒绝。首次矩阵测试错误地将空源字节当作源文件，已按无源契约修正后完整复跑。

以上完成 OpenSpec 任务 4.2 的发布契约；正式 CLI 未切换，renderer profile 事实规则调整仍未完成。报告的 passed 包含预期 unavailable 反例，不能理解成页面事实检查全部通过；原生包仍是第 3.3 节固定快照。

任务 4.1 的节点检查树已有后续进展：页面下保留真实 group/diagram 子节点，每个节点携带原始语义 page/id/path 和独立 scenePath；共享语义 owner 的生成节点不会合并。字段诊断按最近的真实原生节点归属；父节点隐藏、失败或未展开时，其未访问子节点明确为 unassessed。节点局部通过不抵消页面或全局限制，也不授予源编辑能力。最终 `tmp/officekit-native-scene-paint-WbYaHv/integration.json` passed，在每次真实 savePaint 中逐个比对检查树与 scene bindings 的数量和身份，并保留前述发布故障回归；没有解除旧事实规则或完成 profile 接入。

后续已实现可选输入/场景合并：内部 `paintPpjSceneSvg(receipt, { assessInput: true })` 在完成所有页的真实绘制后调用 native profile，再把输入诊断按最近语义 owner 映射到原生节点。输入 path 不变，同 owner 的多个生成节点各有 scenePath；无对应节点的限制进入页或全局，最终警示与 publisher 使用合并状态。公共资产 ID 与原生 ID 不同的情形通过已验证 MIME/hash 对应同一份字节；首次真实合并曾误报 asset-missing，修复后复跑，没有绕过资产校验。

实际报告 `tmp/officekit-native-scene-paint-ewzfZH/integration.json` 中 `combinedAssessment` 四例均完成返回/落盘一致性、输入及原绘制诊断保留、实际 PNG 警示、候选与程序字节不变检查：minimum 为 requires-review，canonical、源 no-op 和源文字编辑为 failed。这里通过的是证据合并契约，不是这些页面的事实检查。完整集成仍因 G-03 四个删除断言失败而退出 1；不得记为整体通过。运行时仍是 G-03 的有界修复包，本轮没有改 C# 或重建；合成专项另验证共享 owner、多节点输入失败、未知全局字段、公共/原生不同资产 ID 及大写摘要。

合并选项默认关闭，尚未接入正式 CLI；其余事实规则逐条映射、更多归属边界及任务 4.1 整体验收仍开放。已有合并代码不再记为“完全未实现”，也不据此勾选整个任务。

## 4. 按功能核对剩余差距

### G-01：共用场景与正式入口

**现状：10/15 个任务勾选，实际绘制与正式入口接入未完成。** 编译采集、传输、归属与适配已有基础，OpenSpec 任务 5.2 的固定版本构建、4.2 的场景发布和 5.1 的成对 fixture 验收通过；任务 3.3/3.4 的绘制要求、4.1 的可靠性接入以及 5.3、5.4 仍开放。这些是任务编号，不是本文章节编号。

剩余工作：把正式入口改为消费真实 scene；按 renderer profile 区分仍适用和已修复的事实规则，并接入已验证的场景发布契约；保留任务 5.1 已通过的固定成对案例，新增字段和组合继续扩充。不能用“两边都编译成功”或“两边都显示占位”代替等价性。

完成证据：真实 CLI 输出包含正确场景几何/数据/样式；返回值与落盘清单一致；缺场景、资产、栅格及写入失败均有明确结果；旧发布保护不退化。

### G-02：文字和形状

显式垂直对齐已有内部显示：top/center/bottom 沿用现有逐行排版，将整段文字块放入上下 inset 界定的可用高度；bottom inset 在该分支实际消费，显式零保留。九项合成正例覆盖三种 anchor、底边距 0/10/30 和两行文字；三项反例保留未知 anchor、空间不足和边距越界的字段级 unavailable。未显式指定 anchor 时不猜继承值，旧排版和未消费字段诊断保留。

真实报告 `tmp/officekit-native-scene-paint-6r8OMm/integration.json` 的 textAnchors 共 12 例通过：作者与去快照源 no-op、三种 anchor、底边距 10/30。40pt 字形在底边距 10 时，居中/底对齐相对 top 下移 16/32px；底边距 30 时为 6/12px，墨迹数均为 911，源 no-op 与作者一致。原始程序和源字节不变，候选 no-op 等于源。使用既有 runtime-slot-fixed，未重建当前 C#；Presentation 4/4 通过，完整集成仍因四项源变换删除失败退出 1。

该映射使用现有简化行高和逻辑下伸空间，并非按字体真实墨迹求出文字块高度；`text-layout` 限制继续存在。字体度量、自动换行、AutoFit、分布式锚定、center/bottom 溢出位置、anchorCenter、源 anchor 修改/删除及完整垂直文字仍未完成。这里只补齐有界的直接锚定显示，不提升完整文字支持等级或切换正式 CLI。

文本框 inset 的证据分类已细化：原先整个 bodyProperties 被标为未消费，现在通过原生描述符逐字段检查，仅豁免实际使用的 left/right/top inset；bottom inset、reset 和其他未使用属性继续定位到具体字段。六项合成对齐/边距案例验证输入不变及底边距限制保留。真实 PgPOma 报告四个作者案例验证右 inset 0→30pt：右对齐墨迹左右边界均左移 30px，居中均左移 15px，墨迹数量相同。右 inset 的绘制计算此前已有，本轮修正的是诊断粒度并补足像素证据，不声称新增独立段落右边距或完整垂直排版；整套仍有四项删除失败。

内部基线推进已修正字号覆盖错误：非空行只按实际 run 的有效字号计算，不再把已被全部 run 覆盖的默认字号算入最大值；空行仍使用默认字号。合成 8pt run / 32pt 默认的反例先失败后通过，另保留默认字号及空行断言。真实 `tmp/officekit-native-scene-paint-oYWAyO/integration.json` 验证作者与源 no-op 的 8pt 文字基线 111.6、实际墨迹及旧位移区域为空，源/no-op 字节不变。未重建原生包，完整集成仍因四项删除失败退出 1；本修复不代表字体度量、换行和 AutoFit 已完成。

内部文字目前还覆盖了以下有界行为，均有合成测试及第 3.2 节运行时的真实回归：

- `fontSpacingPoints`：正负字间距、继承和显式零，映射到 SVG letter-spacing。
- `fontBaselinePercent`：有符号基线偏移，逐 run 恢复、换行重算，不擅自缩小字号。
- `sng/sngStrike`：单下划线和单删除线，保留 none/noStrike/0 覆盖。
- point 段落间距：指定行距、段前、段后以及零值；默认行高仍采用简化排版。
- `marginLeftEmu` 与左对齐 `indentEmu`：整段边距、首行有符号偏移，后续行恢复段落边界。PPJ hanging 以负缩进进入原生状态。

真实断言不仅检查属性：包括基线上下 12 像素、字间距对应墨迹宽度、行距编辑对应 -10/+5/-10 像素、边距清零对应 -20/-30/0 像素。相关源编辑从原始源执行，重新投影并检查非目标 ZIP 保留；限定案例仅改变 slide1.xml。

正式路线仍存在带文字形状只画文字、普通几何退化为矩形、同段 run 错拆成多行等问题。内部路线已修复有限几何与 run 边界，但文字仍使用简化行高和基线。

剩余工作包括字体度量和替代、中文/混排换行、倍数段落间距及继承删除语义、列表、垂直布局、AutoFit、溢出、双线/波浪等装饰格式，以及剩余 preset 与 adjustment。custom geometry 中的 guides、公式引用、text rectangle、独立路径 viewport 等也应分别核对保留、求值和绘制状态；当前 literal 路径函数会拒绝未解析引用。

风险：文字可能溢出、遮挡、错位，节点外形可能不能表达其含义。最小收口案例应包含带文字非矩形、同段混合格式、长中文、显式零值和复杂路径；同时断言实际边界与未支持字段的诊断。

### G-03：变换与可见性

四类旋转/镜像的作者与源 no-op 回归现已通过，完整 G-03 仍开放。新增源编辑回归中显式归零通过，但字段删除失败；首次保留这些失败的完整集成 `tmp/officekit-native-scene-paint-REM5Lk/integration.json` 为 failed、退出 1，本次复核的 PZD4Ht 报告仍保留同样四项失败，具体原因见本节末尾。历史固定包曾在四个去除私有快照的源组投影上返回 `codec_failure`，未进入源 no-op 或绘制；`tmp/officekit-native-scene-paint-eIGtvx/integration.json` 保留该旧失败。投影故障的修复证据与新增删除缺口分别判断。

为定位失败，临时托管测试直接调用固定源码快照的投影器，在 IncludeNodeMap 关闭/开启时均通过同一个旋转组。随后独立诊断包保留的 `tmp/preview-transform-diagnostic-jqYdBe/probe-result.json` 明确记录 NativeAOT 异常：`ProjectGroup` 经 `JsonNode.ConvertFromValue<T>` 请求 `System.String` 的运行时 JSON 元数据，但 `EmptyJsonTypeInfoResolver` 无法提供。失败发生在组的 readingOrder 构造，不是 SVG 旋转函数，也不能解释为所有旋转输入都失败。

当前源码已将 `readingOrder.Add(childId.GetValue<string>())` 改为 `readingOrder.Add(StringNode(childId.GetValue<string>()))`，复用页面 readingOrder 使用的 AOT 安全字符串节点构造；测试增加完整子节点顺序断言。独立副本的 `tmp/group-fix-results/preview-group-aot-fix.trx` 记录 39/39 托管 scene 专项通过、零跳过。诊断异常详情补丁已恢复为原逻辑，再构建 `tmp/preview-transform-diagnostic-jqYdBe/runtime-fixed/`；随后真实 NativeAOT 探针成功，完整集成最终报告 `tmp/officekit-native-scene-paint-3gbXxj/integration.json` 为 passed，退出 0，transformProfileFailures 为空。

最终回归包含 90 度、水平镜像、垂直镜像、组合变换各自的作者和源候选，共 8 个实际 SVG/PNG 像素案例；四个源输入去掉私有快照，验证完整 readingOrder、no-op 字节相同和输入不变。报告记录每例 source/candidate/scene 摘要，原始源 PPTX 另存为 `profile-*-source.pptx`。只有全部 owner 绑定的原生字段匹配，且实际成功绘制了变换，才按字段解除旧 transform 事实错误；缺绘制记录和无关规则继续保留。

修复包 PPJ SHA-256 为 `6b2c3f0c65624ef9f9fc81dc8e9fba721edbbd0ea9a674f05b172469ee3202b4`，manifest 为 `073333a05ec5a4a414d9f7a2e14631f41bab1af3c6241a5d1474fd12e6220653`。这是固定 4b7cd9c6 快照加组字符串节点修复，不等于当前工作区全部 C# 已验证，也没有对新修复包重做双构建一致性检查。该轮 presentation 4/4、两项生成摘要检查、gate-policy 和严格 OpenSpec 验证通过。限定 AOT 故障已关闭；源变换修改/删除、复杂嵌套、完整 profile 合并和正式入口仍未完成。

此前实现的 scene profile 可见性映射也保留：只有编译场景及输入身份匹配，且某语义 owner 的所有原生绑定均为 hidden 并实际经过 painter 隐藏分支时，才免除旧 renderer 的 visibility 事实错误。默认 canonical profile 保持旧检查；上述变换映射单独验收，不解除其他限制。真实作者和源 no-op 案例验证同一形状从橙色像素变为隐藏后的背景像素；缺记录、错输入/场景或缺 registry 映射保持错误或拒绝。早期报告 `tmp/officekit-native-scene-paint-SYsUNl/integration.json` 使用固定 4b7cd9c6 原生包通过，这些案例也保留在修复包最终回归中；正式入口及输入/节点检查树合并未完成。

正式路线未完整处理 rotation/flip、group childFrame 和 hidden。内部已有这些基础映射，但还需完整的嵌套组、组件展开、旋转后锚点、导入变换和效果外扩案例。

组 childFrame 的事实规则已有有界映射：全部语义 owner 绑定必须是成功绘制的原生组，外框与 childFrame 的八个值逐项匹配，才解除旧 `group-coordinates-ignored` 错误。合成反例覆盖输入/场景尺寸不匹配、隐藏组和零子尺寸绘制失败；这些状态继续失败。真实 `tmp/officekit-native-scene-paint-KEpYTd/integration.json` 的 groupCoordinates 四例通过：两层嵌套、非零子原点、两种缩放，各含作者与去快照源 no-op。改变外层 childWidth 100→200 后，橙色子形状中心从 (150,136) 移至 (125,136)，互相对应的旧位置为背景；源 no-op 字节相同，缺绘制记录或 registry 映射不通过。两个组 owner 分别检查，不把父子当一个节点。

该轮只修正内部 profile 的有证据误报，保留字段级 partial 和无关规则，未新增源 childFrame 编辑/删除或复杂组件验收。旧变换测试相应增加“组错误确实解除、缺记录时仍存在”的精确断言，而非豁免所有其他诊断。presentation 4/4、两项生成检查与 gate-policy 通过；完整集成仍因四个源变换删除断言失败而退出 1，不能记为整体绿色。原生包不变，正式入口仍未切换。

隐藏页是否包含、如何编号和如何选择也需要明确契约。元素未显示可能是合法隐藏，也可能是渲染失败，不能在证据中混为一类。完成条件是同一有效场景中的坐标、层叠顺序和可见性进入实际 SVG，并有嵌套/反向/边界测试。

#### G-03 源编辑补充：删除被写成显式零/false

同一修复包上的新增回归分别从原始源投影发起请求，不串联候选编辑。四类源组（rotation、flipH、flipV、三者组合）各执行显式归零和字段删除，结果如下：

| 请求 | 候选画面 | 重新投影与实际 XML | 验收 |
| --- | --- | --- | --- |
| 显式设置 rotation=0、flipH/flipV=false | 橙色子形状恢复原位，旧位置为背景 | 对应字段保留明确的 0/false | 4/4 通过 |
| 从 frame 删除对应字段 | 像素也恢复原位 | 字段仍存在；组 xfrm 留下 rot="0"、flipH="0" 或 flipV="0" | 4/4 失败 |

两组候选均通过实际像素、完整 readingOrder、子节点 frame 与非目标 ZIP 检查：文件清单不变，仅 `ppt/slides/slide1.xml` 有差异。删除例随后在字段缺失断言失败，不能因画面相同就计为删除成功。最终报告记录请求/实际 frame、源/候选/scene 摘要；`profile-*-delete.pptx` 与 `profile-*-delete.reprojected.ppj` 保留实际结果。原作者/no-op 八例及独立回归继续执行，四个删除失败最终抛 AggregateError，未跳过或改成预期成功。

当前 `PpjPresentationCompiler.TryCollectFrameLeafMutations` 根据解析后的 Rotation/FlipH/FlipV 值生成快速编辑，把缺失字段的默认值也写成 0/false；它没有按原始 JSON 属性存在性区分删除和显式覆盖。组 `ApplyFrame` 的内存更新不能纠正随后选中的 token-splice 候选。scene 再导入该候选，正确反映了仍存在的零值；这不是 painter 自行补出的属性。

修复需明确原生 frame 快速编辑的属性移除契约，并检查 presence-only 删除、组合变换中只删一项、源与非目标内容保留，以及显式零/false 继续存在。它涉及源编辑输出语义，超出当前只读 scene change 的边界，待确认扩展或另建原生 change 后实施；本轮没有修改编译器。首次测试将 program 误传字符串，按 workspace 的 Uint8Array 契约修正后才得到上述原生失败。presentation 4/4 仍通过，但不含这些新的真实原生删除断言，不能抵消该集成失败。

### G-04：样式、主题和效果

内部普通形状和 literal 自定义路径已复用共享 linePaint，覆盖虚线、点线、点划线、cap/join、宽度和 alpha；保留 none、零值及路径 stroke=false。合成和真实源编辑测试检查线段/空隙像素及重新投影。虚线节距仍是明确标注的 review 近似，不等于 Office 精确轮廓；主题、复合轮廓和路径端点箭头仍未完成。

内部直接 RGB、部分 alpha 和线条属性已有映射，但未完成 theme/master/layout、命名样式及继承状态的统一解析。渐变、图片填充、阴影、内阴影、glow、reflection、soft edge，以及效果缩放/斜切等不能由已有字段保留推导成已绘制。页面背景也只覆盖有限直接状态。

风险不止外观：低对比度、透明边缘和阴影范围错误会让 reviewer 错判内容是否存在或是否越界。应消费确定的有效样式；未解析主题值不得静默替换后宣称正确。完成证据应包括直接值覆盖继承值、删除后恢复继承、显式 0/false、半透明叠加及效果边界。

### G-05：图片

**不要再把裁切、有限遮罩和边框整体记为“内部未实现”。** 它们已有第 3 节的限定实现及真实证据，仍未进入正式 CLI。

当前具体残留：

- tile 尚未实现实际重复绘制；内部 painter 已修正未裁切 tile 被画成单张 stretch 的错误，现在有/无裁切均明确失败并保留对象占位。真实 jHmDS0 报告验证作者/源 no-op 四例，源与 no-op 候选不变；三种裁切存在性另有合成回归。这是防误画修复，不是 tile 渲染通过，也未接入正式 CLI。
- 全部 fit/focus 输入到编译后状态的等价性尚未验证；不能只看最终 frame 就认定 cover/contain 等全部正确。
- 遮罩仅覆盖少量 preset 与可解析 literal 路径；其他 adjustments、复杂填充模式和公式引用仍缺。
- 边框需要已解析 RGB；未解析主题色会产生 unavailable 诊断，该边框不绘制，图片内容仍可保留。虚线长度和部分线条轮廓仍是有说明的近似。
- 阴影、反射、软边及复杂透明合成未完成。裁切几何正确不等于识别出了图片主体。

后续使用同一资产快照，不能重新读路径造成编译/显示不一致。最小案例要验证非居中裁切、负边留白、透明边缘、遮罩与旋转组合、阴影扩边、空字节和解码失败；主体范围不确定时不擅自裁切。

tile 的后续实现需要先补足尺寸语义：当前 PresentationImage 仅有参数为空的 tiled 布尔状态，scene 不提供已解析的平铺单元物理尺寸；源图片解析接受 useLocalDpi 的 0/1，但没有对应传输字段。不能把所有源图片一律按固定 DPI 或按图片框大小平铺。需要核对原生导入、有效 DPI/固有尺寸及缺省规则，再验证实际重复数、偏移、裁切、透明度和遮罩组合；不把当前失败占位视为此项的最终交付。

### G-06：表格和连接线

内部表格已用真实尺寸和合并范围，正式路线仍简化为均分网格。剩余包括合并外围边框、冲突边优先级、继承/banding、单元格边距和排版。含可见文字的覆盖格也需要明确处理，不能通过合并静默抹掉内容。

内部连接线已使用原生端点，并消费 elbow 的 literal bendAdjustment，显式 0、负值和超过 100000 的值有合成覆盖。剩余是 curved、完整导入变换来源、任意 connection-site 求解/编辑、嵌套及组件锚点广泛回归。自动避障若不在当前支持范围，应明确限制；不能以 frame 中线或“最近对象”猜测关系。

对象锚点的两例通过只证明那两个场景。新 connection-site 接口或在途变更需要独立核验，不能提前记为 SVG 已覆盖。完成证据应包含目标移动、重定向/解绑、反向端点、折点变化、箭头方向，以及源文件和非目标部件保留。

### G-07：source-bound、opaque 和静态检查边界

SmartArt 必须分开判断作者输入、源导入和源编辑保留：

| 范围 | 已验证结果 | 剩余差距 |
| --- | --- | --- |
| 作者输入的 verified drawing | 内部按实际缓存子形状、文字、连接线及组坐标绘制；process 节点颜色和连接线有断言 | 仅限定样本；不等于所有 SmartArt 布局已覆盖 |
| 源导入的 drawing | 当前主动返回 unavailable，占位并记录 `diagram-import-incomplete` | 导入器只按语义节点重建有限几何/文字，缺完整填充、轮廓及连接线 |
| 源 no-op 与文字修改 | no-op 字节相同；候选语义文字、重新投影和限定部件所有权检查通过 | 不代表编辑后源图视觉保真；旧缓存清理和复杂共享部件未验收 |

`PptxSmartArtCodec.ReadCachedDrawing` 的 `drawingCacheVerified` 只证明有限节点缓存几何，不能当作完整视觉证明。补齐时应复用已有原生解析能力，覆盖真实缓存形状及连接线后，才解除源导入绘制限制；不得把 SmartArt 扁平化为可编辑 slide 子形状。

源编辑的部件路径可以合法变化：`PptxCodec` 对独占目标图先创建新部件图，再删除旧目标部件。因此旧测试“ZIP 清单不变且只能修改 /diagrams/”曾失败，但该假设不符合实现契约。当前测试从 graphicFrame 的 dm/lo/qs/cs 和 dataModelExt 的 relId 解析 data/layout/style/colors/drawing 五类实际目标，不按目录整片豁免；图外文件逐字节检查，共享 slide、关系和 Content_Types 检查非目标内容，layout/style/colors 检查内容相同。属性顺序及已知冗余同 URI 根命名空间声明被有限归一化，兄弟文字修改反例仍必须失败。

第 3.2 节通过报告取代前次 `XyIRJd` 的所有权断言失败结论，不取代导入缓存不完整的结论。源图视觉仍不可用；旧 drawing 的保留/清理、外部共享图和最新原生运行时仍要分别核对。历史失败详情保留在原始审计，不作为当前仍失败的重复条目。

#### G-07 补充核对：缓存写出与源编辑也存在损失

2026-09-10 继续核对时 HEAD 为 `362dffe6`。该次源码和第 3.2 节留存 PPTX 显示，缺口不能只归因于导入器。以下“当前 Drawing/源码/目标缓存”均指该次核对对象；本次文档修订没有逐行重审最新 C#，没有新的修复证据前继续保留待关闭状态：

- `PptxSmartArtCodec.BuildCachedDrawing` 只索引 Shape 并遍历语义 nodes 调用 CachedShape，没有写入 drawing 中的 Connector。测试中 authored scene 有节点间连接线，但落盘缓存只有两个节点、零个连接对象。这证明捕获 writer 输入不等于验证 writer 完整消费了输入；不据此推断 Office 宿主重算后的画面。
- `ReadCachedDrawing` 没有读回节点的直接填充；源文字编辑走语义重建后，当前被引用的新 drawing2.xml 中，原节点的直接 `CC5500` 已变成 `accent1` 主题填充。即使主题偶然显示同色，直接样式语义也已经丢失。
- `CachedShape` 还自行构造预设几何和轮廓，使用固定 `lt1` 线色及 solid 虚线类型；修复时应核对当前 Drawing 中的样式、变换及自定义路径，而不是仅补连接对象计数。

实际检查的文件位于 `tmp/officekit-native-scene-paint-2MsoJ4/`：

| 文件/当前缓存 | 节点数 | 连接对象数 | 节点直接橙色填充 |
| --- | --- | --- | --- |
| diagram-source.pptx / ppt/diagrams/drawing.xml | 2 | 0 | `CC5500` |
| diagram-candidate.pptx / ppt/diagrams/drawing2.xml | 2 | 0 | 已替换为 `accent1` |

源文件 SHA-256 为 `4e22e42428e261abdb78d9ca63bbf266cbd896a211cc0006093b590e9f24ab18`；候选为 `fac1592318e28232ad96b2cc6697064186f8545590ce9aeb75a6b8f2307f5b2b`。候选仍保留旧 drawing.xml，但 data2.xml 的 dataModelExt 关系指向 drawing2.xml；检查旧缓存不能证明当前目标图保真。上述为旧运行时实际产物和当前源码核对，不是最新运行时重跑结果。

此外，同一个 Diagram.Drawing 被 `PpjPresentationCompiler` 的 detach 分支 Clone 为普通 Group，`PpjPresentationProjector` 根据 DrawingCacheVerified 提供 detachSmartArt，`PptxCodec` 使用 Drawing 相等性验证 detach。直接修改共享导入结构会影响可编辑投影与导出，不能当作仅改变预览的内部细节。

因此后续需单独明确 SmartArt 缓存写出、导入、语义编辑和显式 detach 的完整契约，再实现并重建原生包。验收增加：有向边和实际缓存连接一致；文字修改保留直接颜色与轮廓；创建、源 no-op、文字/关系修改、detach 均核对实际目标缓存；同时保留非目标部件与共享关系检查。现有 passed 报告没有这些断言，不能证明上述缺口已关闭。

简单源编辑的生命周期已有证据，复杂源覆盖仍不足。需要补第三方 PPTX、主题/master/layout 继承、原生缓存、OLE/媒体海报及无预览 opaque 的可辨识展示。

修改后继续显示旧快照会误导 reviewer，必须说明预览对应原源还是编辑候选、是否仍有效。保留源预览是显示策略，不是授权把可编辑内容转成图片。原包和非目标拓扑必须保留，无法安全改写时明确拒绝。

完成证据至少覆盖原始源 no-op、目标修改、重新投影、非目标 ZIP 保留，以及过期/缺失预览的诊断。动画、媒体和交互只记录静态检查范围，不能凭 PNG 宣称运行行为通过。

### G-08～G-10：图表公共语义、类型与缺失值

散点实施（2026-09-10）：内部 SCATTER 已消费每系列独立 X/Y、Y 缺失索引、真实零及有界标记；支持线性/对数/反向轴和显式边界，超界点保留数值但不钳到边缘。缺失 pair 不进入数值域，也不补零或跨缺口连线。合成验证非等距 X、多系列不同 X 数组、缺失极值不改变域、反向/对数/非法值、输入不变；有直接 line 状态的连接路径只有合成证据。

真实集成 `tmp/officekit-native-scene-paint-laY7Yi/integration.json` 为 passed：作者 marker 图 X 10→25 对应像素横移 60，源 no-op 字节相同，源 Y 2→5 对应上移 63 像素；缺失不产生零值 marker，末端孤立观测仍可见。新投影保留 X `[0,10,50,100]`，Y 为 `[0,5,null,10]`；源字节及所有非目标 ZIP 成员不变，仅 `ppt/slides/charts/chart1.xml` 改变。使用既有 495daafc 包，不是新 C# 验收；PNG 已查看，非人类校准。

本轮不能算散点全面完成：源码 `OpenXmlChartSpaceCodec` 对所有 scatter 调用 markerOnly 写线属性，`XlsxChartSeriesLineStyleCodec` 写出 `ln/noFill`；即使 style 是 lineWithMarkers 也不能据此画默认连接线。真实反例同时检查 ChartML lineMarker 与 noFill，内部返回 `scatter-line-unresolved`，不生成假线。源 X 修改当前被拒绝为 source-owned，测试保留准确拒绝断言，并非已支持该编辑。首轮误用了原生 IR 不接受的系列 stroke 和 raw ChartML style token，已按源码契约修正；缺失像素断言也从轴线上移两像素，避免把合法坐标轴当成缺失 marker。平滑曲线、气泡 size、完整主题/标签、X 缺失表达与源 X 编辑仍未完成，不能删除这些差距。

正式路线仍自行读取有限 `categories/values`，不能完整覆盖 dataset/encoding、类型继承和复杂通道。内部路线应继续消费编译后的数据，不重新推断数据拓扑。

| 类型或变体 | 内部剩余项/待核对项 | 不能省略的风险案例 |
| --- | --- | --- |
| line | 平滑、堆叠、完整轴/标签/图例、显式缺失显示策略 | `[1,null,3]` 的孤立点不丢失；对数轴不接受非正观测 |
| column、bar | 负值百分比、对数基线、复杂源及完整样式 | 正负分开累计；缺失分母与全零分母不能伪造百分比 |
| pie、doughnut | 多系列/多环适用范围、负值策略、标签随动与继承 | 比例交换必须改变扇区；真实零与缺失分开；多环归属明确 |
| area | 面积、基线、堆叠及缺失区域语义 | 多系列正负值和中间缺失，不用折线冒充面积 |
| combo | 各系列类型、主副轴及单位 | 不同量纲的柱线不能共用错误尺度 |
| scatter | 数值 X/Y marker 已有真实回归；连接模式 noFill、平滑、源 X 编辑及完整标记/轴样式仍缺 | 非等距/各系列独立 X；Y 缺失不落为零点；连接语义与实际线属性一致 |
| bubble | size 映射、面积/宽度模式及各通道缺失仍缺 | 相同 Y、不同 size；单通道缺失 |
| radar | spoke/value axis、闭合策略、系列与标签 | 中间缺失不擅自闭合成完整观测 |
| heatmap | 编译生成结果、色域/色板及缺失的完整回归 | 相同值跨图颜色一致；缺失不等于最小值 |
| treemap、sunburst | 层级归属及编译生成路径的完整语义回归 | 两层以上父子，不能画成单层比例块/饼图 |
| waterfall | delta/total 角色、累计起终点及尺度 | 中间 total 和负增量，不能把 total 再次累加 |
| candlestick | open/high/low/close 通道及上下界 | 四通道不同且部分缺失，不能只有背景或错误蜡烛 |
| sankey | 源/目标/流量关系与生成路径回归 | 汇流后分流，路径和粗细与关系一致 |
| symbol、stream 变体 | 固定图标单位、实际带状累计和缺失策略 | 单位变化影响数量；不把粗折线当带状面积 |

这张表覆盖 16 个 PPJ chartType，symbol/stream 另属变体。内部 native chart 分支目前直接处理 LINE、PIE、DOUGHNUT、BAR（其中含 column/bar 方向）及有界 SCATTER；其他图表有的由编译器生成形状/路径。因此，“没有对应 native chart 分支”不等于“一条路径也画不出”，反过来，少量生成路径成功也不等于整个高层图表通过。

公共残留包括 legend、标题/轴标题、numberFormat、dataLabels、marker/pointStyles 的完整范围、trendlines、errorBars、chart/plotArea 样式及源主题。轴和标签可能影响事实理解，不能全部归入低优先级美化。

缺失值应保留原观测身份：默认不得补 0 或跨缺失连线；真实 0 必须可区分。显式 `zero/span/connect` 应被识别为显示策略，记录其与原始数据的区别；尚未实现时明确拒绝，不能忽略或擅自选择。堆叠中未知段会影响后续起点，百分比中未知值会影响分母，不能只用“跳过 null”解决。

### G-11、G-12：已有检查和发布契约，不是绘制完成

支持诊断清单为 9/9，输出安全与证据清单为 8/8。已具备的限定能力包括支持状态/事实失败区分、图内警示、资产快照复用、非覆盖发布、失败产物和输入/输出摘要记录。

这些清单完成结论归属正式旧路线。内部 scene 发布另已完成任务 4.2：场景身份绑定、两个作者输入和源 no-op/文字候选的真实发布、24 例操作失败及两例事实失败警示都有第 3.4 节证据；页级/全局 assessment 和独立场景地址也已实现。尚缺的是 profile 与正式入口的完整接入，不能把已验收的内部发布写成“未实现”，也不能将其扩大为正式 CLI 已使用新路线。

当前声明为 16 类元素中 9 类 partial、7 类 opaque；16 种 chartType 均 partial，没有整个类型提升为 supported。这是保守声明分布，不是完成率。

剩余是 G-01 对新路线的完整对接：节点检查树已有第 3.4 节的可选原始输入合并，仍需完成其他事实映射、更多归属边界及正式入口接入。场景身份、源 no-op/文字候选发布、操作失败与事实失败警示已有证据；新类型或新错误分支仍须扩充回归。旧事实规则只能在有对应修复断言后按 profile 调整，不能为了让新路线变绿全部删除。

### G-13：CLI 与交付流程

正式渲染函数未提供页面选择选项，仍遍历全部页面；页选择参数的解析、实际过滤、隐藏页策略及清单编号尚需端到端闭合。最终摘要应让用户知道哪些页可参考、哪些事实失败、哪些只能保留源预览。

结构检查、渲染检查、编辑保真检查需要分别记录。文件成功落盘不是三项通过；外观评分也不能抵消事实硬门槛失败。失败之后应保留错误与产物，按具体原因调整输入、依赖或实现后复跑，并保留最终结果，不把中途失败静默抹掉。

### G-14：测试覆盖

任务 5.1 完成证据：`tmp/officekit-native-scene-paint-bt1cbj/integration.json` 的 nestedPairFailures、stylePairFailures、datasetPairFailures 均为空。dataset 固定对照现同时走 native line 和 vector heatmap；后者验证六格几何、实际 RGB、缺失绿色与真实零黑色、完整原生组载荷及整页栅格，生成节点分别保留 scenePath 和原图表 owner。加上嵌套/repeat/slot、命名样式/grammar 及其组合，任务指定的类别已齐。原生包仍为下述 runtime-slot-fixed，未重新构建当前 C#。

该任务完成不等于 G-14 完成：更多图表变体、字段级分母、第三方源、workbook、跨版本和人类证据仍缺。完整集成继续因四项源变换删除失败退出 1。下文保留各历史步骤的具体证据，任务 5.1 当前统一按已完成记录。

嵌套样式组合另在 `tmp/officekit-native-scene-paint-x2Cj1e/integration.json` 验证：保留 plain 对照，新增 styled 对照，将下述固定嵌套组件与命名样式 fixture 组合。两个重复 slot 的字体/字号/false 和蓝色墨迹、定义与 slot 原始路径、不同实例身份、缺失/真实零、完整原生载荷及整页栅格均通过。报告明确记录两个 variants；独立 style/dataset 对照仍通过。该轮仍使用相同 runtime-slot-fixed，未重建 C#；整套仍有四项删除失败。小 slot 的文字超出 frame，本例检查实际墨迹范围，不据此宣布溢出排版已完成。

命名样式/grammar 新增[固定对照](../test/fixtures/presentation/preview-style-grammar-equivalence.json)及[案例说明](../test/fixtures/presentation/preview-style-grammar-equivalence.md)。本次实施实际运行 `tmp/officekit-native-scene-paint-BhbhbY/integration.json`：styleGrammarPair 通过，stylePairFailures 为空；形状/文字命名样式中的颜色、字体、字号 token 与独立显式值的完整原生载荷和整页栅格一致，显式 false 覆盖命名 true、实际橙色/蓝色墨迹、owner/frame、缺失值与真实零均有断言。两侧各自 scene 开关不改变候选字节，输入与 fixture 摘要不变。

首轮 lgFqET 的 grammar token 与基础主题 ink 重名，按既有主题优先规则得到不同颜色，像素断言正确失败；改用独立 label-ink 后通过，没有改变编译器语义。使用 G-14 下文相同的两项 AOT 修复包，没有重建当前 C#。嵌套/dataset 两组仍通过，完整集成仍因四项源变换删除失败退出 1；presentation 4/4 不抵消这些失败。该对照关闭的是有限作者样式证据缺口；完整主题继承、源样式编辑/删除及全部图表变体的覆盖仍开放。

任务 5.1 已验证一对[固定嵌套重复/slot 对照案例](../test/fixtures/presentation/preview-nested-repeat-equivalence.json)，[案例说明](../test/fixtures/presentation/preview-nested-repeat-equivalence.md)列出比较边界。高层侧为两层组件、重复实例及 slot 文本 0；显式侧独立写出坐标，两侧共同包含 `[1,null,0]` 折线风险。实际比较的有序原生视觉载荷逐字节相同、整页 raw raster 相同，来源路径/实例身份单独验证，scene 开关不改变各自候选字节。比较没有把组件输出反生成为显式输入；外层 ID/归属不作为视觉载荷字节比较，而有独立断言。

旧修复包实际报告 `tmp/officekit-native-scene-paint-UO1dDt/integration.json` 为 failed：显式侧通过，高层侧在 scene 编译时报 codec_failure。诊断包将异常定位到 `PpjPreviewOrigins.Index` 中的 `JsonSerializer.Serialize(slot.Key)`：NativeAOT 禁用反射序列化，无法构造 slot 来源路径。当前代码改用 `JsonEncodedText` 加 JSON 引号，未改组件展开语义；固定快照托管 scene 39/39 通过。随后按仓库命令构建独立 `tmp/preview-transform-diagnostic-jqYdBe/runtime-slot-fixed`，实际 `tmp/officekit-native-scene-paint-SxRCLN/integration.json` 的 nestedRepeatSlotPair 为 passed、nestedPairFailures 为空；失败的编译侧已有直接修复验证。完整套件仍因四个原生变换删除反例退出 1，未把局部 pair 通过算成整体通过。

新包 PPJ SHA-256 为 `ca442bdfcef428e9d0a6444f661fd77f63ad1a593beda02d406fbffdebb0b9a6`，manifest 为 `a37862d1b45875a150bcecab933a972595c4f626ab727a8e8bade4ac7f070223`。原生来源是 4b7cd9c6 加组 readingOrder 和 slot 来源转义两项 AOT 修复，不是当前整个工作区；未替换安装包，也没有对新包重做双构建一致性检查。fixture 摘要纳入集成前后身份校验。实施时已查看输出图片，但不是人类校准。更多向量/原生图表及其他组合仍需单独对照；命名样式/grammar 的有限成功见本节开头，dataset/encoding 见下文。

现有专项和像素断言证明了真实局部进展，但缺少完整的字段分母、第三方复杂源及所有类型的有效渲染正例。测试中的自行生成文稿去掉私有快照，仍不等于独立第三方 fixture。

dataset/encoding 另已补[固定数据对照](../test/fixtures/presentation/preview-dataset-equivalence.json)及[说明](../test/fixtures/presentation/preview-dataset-equivalence.md)。两系列行交错，混用数组/对象行、列名/列索引通道，独立显式答案为 Alpha `[1,null,0]`、Beta `[0,4,5]`。实际 `tmp/officekit-native-scene-paint-PZD4Ht/integration.json` 的 datasetPair 为 passed：系列名称/顺序、原生值与缺失索引、两个孤立观测和唯一连续线段、整个原生 chart 载荷及整页 raw raster 均与显式侧一致；两侧各自 scene 开关候选字节相同，原始输入与 owner 保留。fixture 纳入运行前后摘要检查。

该轮沿用上述 runtime-slot-fixed，没有修改 C#、增加 JS 数据解释器或重建包。嵌套组件 pair 继续通过，完整集成仍因四个原生变换删除断言失败而退出 1；dataset 的局部通过不覆盖这些失败，也不代表内嵌 workbook 的源编辑已验收。指定类别的后续补齐见本节开头，全部数据通道和跨功能组合仍不是已验收范围。

需要固定预期可编译正例与预期拒绝反例，分别统计 `compilerRejected`、`rendererFailed`、实际渲染及可靠性状态。任一必需正例发生非预期编译拒绝、绘制失败或可靠性硬门槛失败，都应使对应验收失败；不能只在全部正例失败时才报错。预期拒绝反例必须匹配指定原因，不将任意异常算作成功。像素断言要验证比例、端点、空隙和边缘等含义，不能只断言“存在 SVG/PNG”。

外部 PPTX、workbook、跨版本重新投影、人类校准仍需单独证据。Agent 看图不冒充人类评分，LibreOffice/PowerPoint 打开或截图也不代替源编辑保真检查。

### G-15：持续维护和字段防漂移

已有 schema/registry 派生摘要检查，内部 painter 也利用生成描述符诊断部分未消费字段。它们不能保证新增嵌套字段已被正确绘制：登记类型、把字段加入 handled 列表、保留 wire 数据，都可能没有实际视觉消费。

每个新增视觉字段都应记录：模型定义、编译/导入语义、scene 是否保留、绘制入口、限制诊断、fixture 与行为断言。删除字段还要检查继承恢复；source-bound 字段还要检查原源和非目标部件。生成文档和任务状态必须跟实际证据同步。

自定义几何等接口持续新增，是直接的同步风险。第 3.3 节已验证一个更新的原生快照及 JS 集成，但没有为每个新字段补齐视觉行为断言；之后的提交还要更新同样的版本证据，不能只更新能力表。

### G-16：性能和稳定性

**当前结论：限定的作者与 source-bound 托管对象释放已通过；复杂源、NativeAOT 宿主生命周期和保留数据峰值仍待验收。** 分阶段耗时、响应字节、Linux 驻留高水位及一次预算失败后的恢复已有采样。下面按各次运行保留证据；某个早期报告未测对象释放，不表示后续四条源路径也没有证据，更不表示整个性能任务已完成。

真实字节上限恢复已验证：`tmp/officekit-native-scene-paint-AperkA/performance.json` 的 sceneBudgetRecovery 为 passed，性能文件摘要与 integration.json 一致。20 重复组件请求在 `maxUncompressedBytes=4096` 下明确返回 `preview_scene_budget_exceeded`，候选字节为空、scene 不存在；同一 NativeAOT 进程恢复默认上限后，返回完整 41 个绑定和正常候选摘要，输入不变。测试直接断言该错误码，不把其他编译拒绝算作 scene 上限通过。

本例覆盖作者场景字节预算及失败后重试；节点/深度上限目前仍是托管单元证据，不等同所有资源耗尽情形。使用既有 runtime-slot-fixed，没有改生产实现或重建包。整套 AperkA 仍因四项源变换删除失败退出 1；峰值保留对象的计量和 G-16 全部性能目标仍缺，任务 5.3 不勾选。

source-bound 的托管释放回归也已补齐四条限定路径：no-op、文字、frame、语义填充。保留原 PPTX、请求和投影结果时，丢弃返回结果后的 scene、presentation、slide、顶层元素载荷、bindings 和资产引用可被回收。场景与候选独立重新导入一致；no-op 候选等于源，编辑候选不同于源；随后关闭 scene 的候选与开启时相同且不返回旧场景，源和恢复后的请求字节不变。

验证使用包含图片资产及 opaque 兄弟对象的既有源 fixture，新增四例与作者释放等专项合计 47/47 通过、零跳过。精确结果在 `tmp/preview-transform-diagnostic-jqYdBe/tmp/candidate-release-results/preview-candidate-release.trx`，仍是固定源码副本加测试补丁、SDK 8.0.128，不是最新整个工作区或新 NativeAOT 包。下文旧记录中的 source-bound 释放缺证据，现仅由这四种托管案例补足；复杂源拓扑、原生宿主生命周期、分配与保留峰值仍待验收。

对象释放已有作者路径的限定证据：新增 `ReleasedPreviewResponseDoesNotRemainRootedByLiveRequest` 覆盖 minimum/canonical 两种输入和两个托管协议入口，共四例。非内联 helper 丢弃强引用后，保留原请求及资产并强制 GC；响应、结果、scene、presentation、slide、element、形状/组/图表及 bindings 的弱引用全部失效。同一请求改为 scene 关闭后不返回旧场景，候选字节一致，请求恢复后字节不变。

固定 `4b7cd9c6` 加两项 AOT 修复的独立源码副本加入该测试后，SDK 8.0.128 托管 preview 专项 43/43 通过、零跳过，TRX 为 `tmp/preview-transform-diagnostic-jqYdBe/tmp/response-release-results/preview-response-release.trx`。未修改原生生产实现、未重新发布 NativeAOT 包。这证明被检查的作者响应对象可回收，不证明 source-bound 导入对象、AOT 实际进程或 scene 关闭时零分配，也不代表所有对象的保留峰值已量测；5.3 仍开放。

原生内存补充：`tmp/officekit-native-scene-paint-FfHlHq/performance.json` 为两类输入分别启动 scene 开/关独立 PPJ 进程，记录握手后、一次编译后的 Linux VmRSS/VmHWM、PID、启动/调用耗时及实际响应字节。每次使用新进程，完成后在 finally 退役，不借用复用进程的历史峰值。`integration.json` 记录该性能文件 SHA-256，已核对一致。重复组件的关闭/开启驻留高水位为 67.74/86.57 MiB，源候选为 75.80/81.45 MiB；各模式各一个独立样本，不是统计分布或目标上限。

VmHWM 是操作系统报告的整个进程驻留高水位，包含启动、缓存及编解码工作；不是 scene 对象大小、托管堆峰值或泄漏证据。独立调用固定 includeNodeMap=true。非 Linux 明确记录 unavailable；Linux 读取失败或字段缺失令测试失败，不补零。四个测量和候选/输入不变断言通过，整套仍因四项删除失败退出 1。请求结束后对象保留、scene 关闭时分配与资源上限验收仍未完成，任务 5.3 不勾选。下文 DsJ91a 的“未采集原生进程峰值”仅描述该轮历史证据。

已有一轮可复跑的局部基线：`tmp/officekit-native-scene-paint-DsJ91a/performance.json`，由真实集成脚本生成，使用上述 runtime-slot-fixed。同一进程环境先预热 scene 开/关，再各采三次；下表为中位数。重复案例包含 20 个嵌套组件实例、41 个 scene 绑定，源案例为真实源文字编辑候选。响应大小是传输层收到的完整响应字节，包含候选文件，不是单独 scene 大小。

| 案例 | 响应字节：关 → 开 | 编译 ms：关 → 开 | 开启后的 SVG / PNG ms |
| --- | --- | --- | --- |
| 20 次重复组件 | 28234 → 48044 | 10.34 → 14.71 | 3.54 / 18.18 |
| 源文字编辑候选 | 43021 → 46054 | 25.79 → 31.15 | 1.05 / 8.93 |

每次编译验证只有一个原生请求，关闭时无 scene，开关不改变候选字节，原始 program/source 摘要不变。编译耗时包含公共 JS 调用和传输；invoke 耗时另存，不能等同纯 C# 编译 CPU 时间。SVG 与 PNG 单独计时，不含文件发布；本轮与其他门禁同时运行，不能作为独占机器下的吞吐承诺。

报告也记录 JS 进程的阶段前后内存快照，但没有采集原生进程峰值，也没有证明对象释放后不再保留。因此任务 5.3 仍开放。初稿页宽 4400pt 被 Open XML 校验拒绝，改为合法 4000pt 并保留 20 实例后才采样；未放宽格式限制。整套 DsJ91a 仍因四项源变换删除失败退出 1，基准有数据不表示功能验收通过。

当前没有足够证据承诺“高性能”。需要先约定典型页数、对象数、图片大小、文字量和可接受耗时/内存，再量测冷启动与重复运行、编译、scene 传输、SVG 绘制、PNG 栅格和写盘各阶段。

还需检查关闭 scene 时不额外保留预览数据、大图/重复组件/复杂源候选的内存峰值、字体缺失、解码失败、磁盘不足、超时与清理行为。既有运行曾遇到 `/tmp` 的 ENOSPC；应单独分类为资源失败，并验证失败发布契约。不能靠删除旧证据后只留下绿色结果。

## 5. 五条可靠性规则的验收口径

| 规则 | 必须观察的结果 | 当前边界 |
| --- | --- | --- |
| 缺失数据不默认当 0 连线 | 原值、缺失索引、显示策略和 mark 可对应；单点不丢失 | 内部 line/柱条/circular 有限定证据，尚未统一全部类型与正式路线 |
| 图表关系与数据拓扑一致 | 轴、层级、累计、源目标及 OHLC 等按真实通道显示 | 部分基本图表已修；复杂图表及正式旧路线仍缺 |
| source-bound/opaque 不扁平化 | 原源不变、修改范围可解释、重新投影、未知内容明确保留 | 简单源回归通过，复杂第三方和快照有效性未完整验证 |
| 图片信息不确定时保守处理 | 不猜主体；透明、裁切、遮罩、边框及效果边界有证据或明确限制 | 已知裁切/有限 mask/边框有进展，完整效果和输入组合仍缺 |
| 输出前检查独立且不可被外观抵消 | 结构、渲染、编辑保真各自状态；事实失败保持失败 | 诊断/发布有基础，新场景接入与完整交付验收尚未完成 |

## 6. 建议的收口顺序与单项完成标准

按小功能持续补齐，不用等全部功能做完才验证，也不为每个字段新增一套庞大框架。

1. **维护验证快照**：以本次摘要及 G-14 指定的两项 AOT 修复包作为当前回归基线；第 3.3 节保留原始冻结构建与双构建证据，不退回未修复包。后续修改记录源码、dirty diff、二进制 manifest/hash 和测试脚本身份；新字段继续补视觉断言，原生语义改变时再更新精确包验证。当前基线中的四项删除失败仍须保留，不能称作全绿基线。
2. **补基础显示**：文字格式/布局、剩余几何、主题及图片效果；每次选一个字段或一组直接相关行为。
3. **补事实风险高的图表与关系**：源 X 编辑、X 缺失表达、散点连接语义、气泡 size、双轴、层级/流向/OHLC、缺失显示策略；保留已有数值 X/Y marker 回归，优先复用编译生成结果。
4. **完成正式接入**：内部绘制与 profile 检查、来源身份、G-12 发布保护同步切换；未解决项继续明确失败或标注限制。接入本身不等于全面覆盖。
5. **补外部样本和运行验收**：第三方源、字段级 gate、性能基线和人类校准；最后再判断总体目标和 Skill 路由。

这不是要求先补完所有绘制才能接 CLI：可以在冻结的有限支持范围内推进第 4 项，同时继续第 2、3 项；前提是新路线的可靠性诊断、来源身份和发布保护一起验收，未支持范围清楚保留。

每个适用的字段至少应有：作者输入的创建正例；缺失/零值/复杂关系等风险例；实际 SVG 或 PNG 行为断言；源绑定时的原始源 no-op、目标修改及重新投影。涉及删除语义时额外验证删除后恢复默认或继承。某层不适用应写明理由，不能默认为通过。

整体完成不能按任务数换算百分比。至少要同时满足：现有静态功能逐字段有归属和实际覆盖；可靠性硬门槛不绕过；正式 CLI 使用场景；指定当前运行时回归通过；第三方源保真有证据；性能达到事先约定目标。仍不支持的外部原生内容必须有明确保留/拒绝边界。

渲染器面向视觉 review 的最终可用性还需人类检查代表性输出，确认警示可辨认、布局可读、关系不误导；验收样本和标准应预先确定。它与第 7 节决定默认 Skill 路由的盲评及四组人类校准是不同验收，不相互代替。

### 6.1 开放任务与交付物对应表

G-01 的 10/15 是任务勾选数，既不是实现比例，也不是 G-01～G-16 的总进度。剩余五项如下；文字、图表等跨任务问题仍以第 4 节的功能边界为准。4.2 的已完成证据见第 3.4 节，5.1 见 G-14。

| 任务 | 尚需交付 | 最小验收依据 |
| --- | --- | --- |
| 3.3 基础绘制 | 完整消费计划内的形状、文字、图片、组及生成路径，移除正式路线组件猜测 | 成对高层/显式案例的实际几何、样式、ID 与像素；剩余字段明确限制 |
| 3.4 关系与原生内容 | 图表、表格、连接线、源和 opaque 内容沿 scene 绘制 | 数据通道、端点、顺序及候选保留断言；不重解 PPJ、不伪造拓扑 |
| 4.1 可靠性接入 | scene/source 归属、registry 与 renderer profile 事实规则统一 | 未知字段和未解析继承必须保守失败；每条解除的旧限制有对应修复回归 |
| 5.3 性能回归基线 | scene 开关、重复组件和源候选的开销记录 | 响应大小、分阶段耗时、保留数据/内存及关闭 scene 的行为；不等同 G-16 全部验收 |
| 5.4 同步与收口 | registry、派生资料、输出说明、Skill 指导和审计一致 | 生成检查、链接、适用的 Skill 门禁与严格 OpenSpec 验证；其余任务完成后才关闭 G-01 |

### 6.2 修复归属和交付边界

| 工作 | 应修改和核对的层 | 收口方式 |
| --- | --- | --- |
| 正式 scene 接入、节点诊断与发布 | JS 编译调用、scene view、painter、assessment、publisher、CLI | 沿现有 G-01 任务推进，正式入口回归必须实际走新路线 |
| 字体排版、几何、图表、图片和表格显示 | 优先消费已有原生有效状态；缺语义时核对原生模型和导入器 | 每次选一个有界字段/行为，补正例、风险例、实际像素或几何断言及限制 |
| 变换字段删除、SmartArt 缓存/源编辑损失 | 原生编辑计划、PPTX 读写、投影与候选保留 | 属于文件编辑语义，不由 painter 猜测修补；需明确原生 change 范围，再验证实际候选与非目标内容 |
| tile 的尺寸/DPI | 原生图片导入、有效尺寸信息、scene 字段和绘制 | 先确定物理尺寸契约，再画重复单元；缺信息时继续明确不可用 |
| 外部 fixture、字段覆盖、性能、人类校准 | 测试输入、报告、环境与验收标准 | 独立记录证据，不把自生成文件、操作系统高水位或 Agent 看图替代相应验收 |

建议先处理会静默改变事实或源语义的问题，并同步推进有限范围的正式接入。只读 scene change 的完成不能顺带宣告原生源编辑问题解决；反过来，原生文件写出修复也不能替代显示层的像素回归。本文记录修复归属，不在这轮文档整理中扩展实现或修改其他在途工作。

### 6.3 新增功能的维护记录格式

后续每补一个视觉字段，更新对应 G 编号，并留下以下信息即可；不必另建一套验收系统：

- **输入和语义**：字段定义、合法值、缺失与显式零/false、继承与删除行为。
- **实现链路**：作者编译、源导入/编辑、scene 字段、具体绘制函数；哪一层不适用要写明。
- **可观察结果**：位置、比例、文字墨迹、透明区域、边缘或连接关系应如何变化。
- **拒绝边界**：未知/未支持输入的诊断地址、状态和可见提示，不以近似输出冒充正确。
- **验证身份**：fixture、断言、命令、运行时版本/摘要、通过和失败结果；区分合成、原生集成、正式 CLI 和人类评审。
- **同步范围**：对应测试、能力声明、生成资料和文档；禁止只加 handled 字段就宣布支持。

## 7. 与 Skill 评测的关系

本次没有重新核验独立评测工作树，也没有执行六个外部 PPTX fixture、四路线作者任务、两轮盲评或四组人类校准，不能宣布这些实验完成或据此把 Kimi 设为默认路由。

恢复评测时需先核对工作树和冻结版本，再落实各路线 1→10 编辑、每场景的缺失数据/复杂关系/source-bound 输入、事实与视觉分开评分、硬门槛失败不可抵消，以及报告和 `tasks.md` 对齐。评测、渲染器和原生接口是三个相关但独立的验收范围。

## 8. 复核入口与维护方式

| 入口 | 用途 |
| --- | --- |
| [正式 painter](../src/ppj/svg-preview.mjs) | 确认用户实际走哪条路线 |
| [内部 painter](../src/ppj/preview-scene-svg.mjs)、[scene view](../src/ppj/preview-scene-view.mjs) | 核对原生状态到实际图形的消费 |
| [能力声明](../src/ppj/svg-preview-capabilities.json)、[registry](../src/ppj/capability-registry.json) | 核对保守等级与派生元数据 |
| [SVG 专项](../test/ppj-preview-scene-svg.mjs)、[原生集成](../test/ppj-preview-scene-native.mjs) | 区分合成场景与真实编译/源编辑证据 |
| [G-01 tasks](../openspec/changes/ppj-preview-compiler-scene/tasks.md) | 10/15；五项剩余任务见第 6.1 节 |
| [G-11 tasks](../openspec/changes/ppj-preview-support-diagnostics/tasks.md)、[G-12 tasks](../openspec/changes/ppj-preview-output-evidence/tasks.md) | 限定检查/发布契约，分别 9/9、8/8 |
| [输出说明](ppj-preview-output.md) | 发布状态、清单与文件行为 |
| [原始审计](ppj-svg-preview-gap-audit.zh-CN.md) | 历史反例、实施日志和详细恢复入口 |

可按影响范围复跑：

```sh
node test/ppj-preview-scene-svg.mjs
node test/ppj-preview-scene-native.mjs /absolute/path/to/exact-runtime-package
npm run test:slow -- --segment presentation
node test/gate-policy.mjs
git diff --check
```

原生集成参数必须替换为实际构建且核验 manifest/hash 的包目录，不直接复制本机临时路径当作通用依赖。协议变更另跑 `npm run proto:check`；原生包按 checked-in workflow 构建及验证；Skill 变更另跑 portability/reference-sync。命令列出不代表本次全部执行。

以后更新本文时，直接修订相应功能段和证据范围，将旧运行日志留在历史审计；不要在正文不断叠加相互矛盾的“最新补记”。只有行为断言和实际发布范围改变，才调整对应完成结论。
