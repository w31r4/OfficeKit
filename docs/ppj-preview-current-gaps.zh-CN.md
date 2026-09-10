# OfficeKit 本地 PPT 渲染器：当前能力与剩余差距

文字外阴影复验（2026-09-10）：qCO4bS 使用冻结源码 `text-shadow-snapshot-OenB3z` 与同一 035472e9 原生包，新增 34 项形状文字阴影回归全部通过，正式入口共 473 项。原十二项源删除失败仍使整套 failed。外部九项直接阴影对照为七项通过、两项竖排位置失败；模糊与换行仍有限制。详见 G-04 的“形状文字外阴影”小节；没有重建 C#、提升完整覆盖等级或宣称宿主验收。

文字区域复验（2026-09-10）：已补上当前绘制预设的内部文字矩形和有限自定义边值；p4Au9s 在同一 035472e9 默认原生包下通过新增 78 项作者/源 no-op 回归，正式入口共 409 项。两组外部文字对照分别为 8/8、39/39，通过预设边界阈值；不是完整字体或人类验收。整套仍为 failed，原十二项删除失败保持不变。无 preload 的 presentation 门禁 4/4；具体范围见 G-02 的形状内部文字区域小节。 TzaE9j 与旧轮廓对照保留为修复前的历史记录。

核对日期：2026-09-10。当前默认 linux-x64 原生包来自冻结提交 `035472e9ca8f7fd58987148285da5a3f406286d2`，构建、备份及实际结果见第 3.6 节。工作区另有持续更新的原生、测试和能力声明，HEAD 不能单独标识未提交修改。本文区分所读源码、已验证的本机安装包与尚未验证的变更，不代表其他平台或远端发布状态。

本文记录现有 PPT 静态功能的已知差距，不是完整字段覆盖率报告：目前缺少逐字段分母，不能据此声称已穷尽所有 OOXML 属性。最近实施已重建指定快照并运行内部集成及窄门禁；已有九个早期文字方向/边距、两组各五个轮廓案例，以及本轮八个变换和三十九个文字区域/锚定案例的 LibreOffice 对照，轮廓缺省连接组含两项失败，尚无全仓测试、PowerPoint 或人类验收。各历史运行仍按自己的版本和范围解释。

### 本次复核摘要

文字区域阶段整体复验为 `tmp/officekit-native-scene-paint-p4Au9s/integration.json`（北京时间 2026-09-10 17:21:21）：同一 035472e9 默认原生包，正式入口 409 项，新增 78 项文字区域作者/源 no-op 与原 66 项多边形全部通过限定断言、均为 requires-review。36 项调整值源修改/清零/删除继续保留非目标内容，75 项方向及其他独立回归保留。报告仍为 failed，原十二项删除失败未消失。另有 8+39 个本轮外部文字对照通过；此前缺省线条连接的两个轮廓差异仍未关闭。

提交快照复验采用 `tmp/officekit-native-scene-paint-1fIaRG/integration.json`（2026-09-10）。它使用第 3.6 节默认安装包与固定的暂存 JS 快照；随后仅补充能力说明和失败记录。下方历史报告保留各自的版本边界。正式入口已接 scene，默认包已更新，无测试 preload 的 CLI 回归通过：

此前文字方向阶段复验为 `tmp/officekit-native-scene-paint-xTtqoh/integration.json`，北京时间 2026-09-10 16:24:19，使用同一 035472e9 默认原生包与该轮 JS。正式入口有 265 项：此前 190 项加 75 项文字方向，不把重跑累加为新案例。横排、顺/逆时针竖排按各自的阅读坐标布局，消费物理边距并保留独立文字旋转、外层变换和防镜像处理。三类文字 owner 的方向/锚定、组合变换和九项独立源切换/显式横排/删除均为 requires-review。原 47 项旋转、36 项翻转及渐变、图片等回归保留；2 项平铺拒绝、2 项表格 run 内换行 opaque 和 12 项源删除失败仍单独记录。另有九个自生成方向/边距案例经过 LibreOffice 对照，不计作第三方 fixture 或人类验收。并行原生修改没有由本轮重建，不属于这个默认原生包的验收范围。

| 核对项 | 结果 | 对完成判断的含义 |
| --- | --- | --- |
| 正式预览入口 | 请求 `includePreviewScene: true` 并实际绘制 scene，旧组件/图表猜测分支已删除 | 本机默认 linux-x64 包的无 preload 回归通过；其他平台和后续版本仍需复验 |
| 输入可靠性检查 | 正式 `renderPpjSceneSvg` 强制合并原始输入与场景检查，没有关闭选项 | 内部 paint-only 仍可独立测试；公共入口不能借此跳过检查 |
| 正式入口案例 | 473 项文字阴影/形状图片阴影/文字区域/多边形/文字方向/翻转/旋转/背景/形状填充/组件/样式/dataset/源候选通过限定内容像素、身份、发布与输入保留检查；另有 2 项平铺背景拒绝发布及 CLI 子进程回归 | 成功发布不代表可靠性通过；34 项文字阴影、30 项基础阴影、78 项文字区域、原 66 项多边形、75 项方向及原 36 项翻转、47 项旋转等案例仍为 requires-review |
| dataset 折线检查 | 两个真实数值变体解除旧通道忽略误报；1→2 对应孤立点实际上移 42px | 缺失、真实零、显式数据对照通过；heatmap 和其他未验证通道仍保留错误 |
| 显式连接线端点检查 | 六个实际场景通过逐端点匹配，已绘制的直线/折线不再误报端点被忽略 | 对象锚点、坐标不符、缺绘制记录及其他限制仍保留 |
| 成对 fixture | 嵌套组件/repeat/slot、命名样式/grammar、dataset/encoding、native line 与 vector heatmap 的固定对照通过 | 任务 5.1 指定类别已有证据；不代表所有字段、图表变体、workbook 或第三方源均完成 |
| 文字近期进展 | 小字号基线、物理边距、直接锚定、旋转/防镜像及三种方向；新增 78 项预设/自定义文字区域，保留 75 项方向回归，九项源切换/删除通过 | 完整换行、字体度量、其他预设/公式文字区域、多列/WordArt/东亚竖排、upright、原生图表独立标签及 AutoFit 仍缺；单 run 换行表格仍会导入为 opaque |
| 直接图片背景 | 12 项作者/源 no-op/独立源编辑通过原生状态、RGBA、前景与共享媒体保留检查 | 支持透明像素、正负裁切和 opacity 存在性；共享图片的整个背景删除在原生编译阶段失败 |
| 形状图片填充 | 五种几何各有作者/源 no-op，加六项源修改与删除，共 16 项实际裁切、RGBA、轮廓/文字及正式入口回归通过 | 仅已映射几何和完整 imageFill；旧资产身份字段、平铺、其余 preset/效果未完成；仅解除已有完整几何记录的旧事实错误 |
| 图片 tile | 作者/源 no-op 的明确拒绝、占位和输入保留断言通过 | 通过的是防误画测试，不是平铺渲染成功 |
| 源变换字段删除 | rotation、flipH、flipV、组合四例失败 | 删除仍留下显式 0/false；画面恢复原位不能抵消编辑语义失败 |
| 源段落间距删除 | 行距/段前/段后倍数的七种非空删除组合全部被 PPJ 编译器拒绝 | 尚未生成候选，不属于渲染失败；原始源和七份请求保留 |
| 线性与居中径向渐变 | 共享直接 RGB/alpha 绘制，原线性/背景与 22 项径向案例保留；新增 36 项特殊亮度插值及普通插值切换案例通过 | 特殊曲线已有有限精度绘制，不是仅提示；非居中焦点、其他路径类型、继承背景及其他消费者仍缺 |
| 报告整体状态 | `failed`；4 项变换删除、7 项间距删除、1 项共享图片背景删除失败，其余回归失败数组为空 | 十二项失败均保留；背景删除报 `presentation_element_binding_mismatch`，尚未生成候选 |
| G-01 任务 | 10 项勾选、5 项未勾选，共 15 项 | 只表示任务记录，不是渲染覆盖率或总体完成率 |
| 性能证据 | 重复组件/源候选有分阶段采样、独立进程驻留高水位及预算失败恢复记录 | 保留峰值与完整负载目标仍缺，5.3 未完成；托管释放证据另见 G-16 |

该报告的原生来源为冻结提交 `035472e9`，包摘要见第 3.6 节；之后的工作区原生变更不在该包内。性能证据按该报告的 `performance` 字段关联，不能沿用先前报告的文件摘要。报告中的 SmartArt 源绘制是明确的 unavailable 反例，不因失败列表为空就变成视觉保真成功。历史成功报告保留为对应版本的证据，不能覆盖新增的失败断言。

## 1. 结论与阅读范围

**渲染器尚未整体完成。正式 CLI 已消费编译器场景，本机默认安装包已有真实成功案例。** 当前不能把“生成了 PNG”“测试通过”或“原生接口能保留某个字段”解释成现有功能已全面、正确地显示。

主要差距分为五组：

1. 场景绘制、检查和发布已接入正式 CLI；默认包交付已完成一次验证，但剩余字段映射及后续版本的配套复验仍需落实。
2. 文字排版、主题继承、效果及部分静态元素仍缺实际绘制。
3. 图表只有限定类型和变体具备内部语义回归，复杂坐标、关系及公共样式仍不完整。
4. 已有 source-bound 测试主要来自本项目生成的简单文稿，不能代表第三方复杂 PPTX。
5. 字段级防漂移、后续版本持续复验、性能基准和人类视觉校准尚未闭合；一个固定版本的构建验收已通过。

本文面向维护者和后续开发者，目标仍是**本地个人使用的 Skill 与 CLI 运行时**，不引入前后端系统，不要求 PowerPoint 像素级复刻。范围是 PPT 静态页面预览；Word、Excel、PDF、Live 宿主、媒体播放和动画执行不计入静态渲染完成度。

本文沿用[详细审计](ppj-svg-preview-gap-audit.zh-CN.md)的 G-01～G-16 编号。旧审计包含实施日志和历史反例，正文中部分“当前”属于较早快照；判断本次现状以本文明确标出的源码、运行时和证据范围为准。[历史附录](ppj-svg-preview-gap-audit-history.zh-CN.md)继续用于追溯。

### 1.1 几个术语

- **正式路线**：用户当前调用 `officekit ppj preview` 实际执行的代码。
- **内部入口**：独立测试的 paint-only 调用，与正式预览共用 painter，但保留内部提示；它的默认选项不是正式入口的检查开关。
- **native scene**：编译器提供的原生绘制场景，包含有效对象状态及来源归属；不是把 PPJ 再解释一遍。
- **source-bound**：编辑与原始 PPTX 中的对象和部件绑定，必须保留非目标内容。
- **opaque**：无法安全解释或编辑的原生内容。可展示可信源预览或明确占位，不得猜测其结构后重写。
- **重新投影**：将编辑后 PPTX 再导入为模型，检查实际写入结果，而非只检查内存中的编辑请求。

### 1.2 差距总表

状态按行为和交付范围判断。“部分实现”指存在可用子集；“待验收”指有代码或产物但缺指定回归；“未闭合”指整个交付条件尚未满足。下表不是新增的能力声明，也不按行数计算完成率。

| 编号 | 当前差距及状态 | 对用户的影响 | 关闭前必须补的证据 |
| --- | --- | --- | --- |
| G-01 | 正式入口与本机默认包接入已验证，剩余任务未完成 | 支持的输入可预览，未覆盖的语义仍需检查 | 完整场景字段、剩余事实映射及后续版本配套复验 |
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
| G-12 | 场景发布与本机默认包的限定回归通过 | 发布完成不等于事实/保真通过 | 新错误分支与后续实际环境验收 |
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

## 2. 正式入口已接场景，为什么仍需配套原生包

| 环节 | 正式 CLI | 内部 paint-only 入口 |
| --- | --- | --- |
| 输入与编译 | 加载 workspace，请求 `includePreviewScene: true` | 测试提供已编译、可验证的 receipt |
| 绘制依据 | 原生有效状态、坐标、数据及资产 | 同一 painter 的原生 scene view |
| 主要实现 | `svg-preview.mjs` → `renderPpjSceneSvg()` | `preview-scene-svg.mjs` 的 `paintPpjSceneSvg()` |
| 检查与发布 | 强制输入/场景合并，验证候选与场景身份，复用 SVG/PNG、清单及非覆盖保护 | 可单独查绘制，带内部提示，不能代替正式检查 |
| 当前意义 | 当前默认包下已通过无 preload 的 CLI 回归 | 隔离验证绘制行为，不证明完整交付 |

正式 `renderPpjToSvg()` 现请求 `includeNodeMap: false, includePreviewScene: true`，通过只读场景验证后绘制。旧 label/value 组件排版、重新读取通道的图表绘制和均分表格分支已删除，没有隐藏回退。`programJson` 仍保留 canonical 身份和原始输入检查含义，不作为第二套绘制模型。

当前 p4Au9s 的 `productionEntryCases` 有四百零九项：文字区域七十八项、多边形六十六项、文字方向七十五项、文字翻转三十六项、文字旋转四十七项、亮度插值三十六项、径向渐变二十二项、形状图片填充十六项、图片背景十二项、直接线性背景渐变九项，以及组件/样式/dataset/源文字候选十二项。两项平铺负例另列 `backgroundImageRejections`，两项表格 run 内换行反例另列 `textRotationOpaqueCases`，均不计入这 409 项。测试只注入内存加载，真实编译、强制合并检查和发布走正式函数；每页除顶部 24px 审查条外的完整像素与已验证内部结果一致，输入/源字节、scene/候选摘要、JSON 序列化及返回/落盘一致均通过。文字方向/翻转/旋转、dataset line、图片背景、形状图片与渐变正例均为 requires-review；encoded heatmap 仍为 failed。形状仅解除已有完整几何及自身/祖先变换记录的 `preview.fact.shape-geometry-omitted`，其他限制保留；具体映射见 G-11/G-12。

此前 7pSsUc 复核将十六张形状图片填充的正式 PNG 与 1A1hYE 对应图片比较，除顶部 24px 警示条外，全部内容区像素逐字节相同。警示由红色失败变为棕色待复核，并未通过改变页面内容隐藏问题。实际源翻转及本轮径向/自定义曲线 PNG 已由 Agent 查看，不计为人类校准。

真实 CLI 子进程回归验证作者、去快照源 no-op 和重复输出目录拒绝，本轮文字区域及 unavailable 诊断回归修正后的无 preload 产物为 `tmp/officekit-ppj-preview-054XQA`，presentation 4/4。旧 isytO3 曾用 preload 选择独立包，旧默认包曾返回 `preview.scene.missing`；两者属于更新前记录。现在默认包已更新，后续仍需随运行时变更重跑，而不是用一次成功覆盖所有版本。

先前 bIwSP1 的十项入口及七项渐变回归属于 c8b0d324 包；L4nZbV 增加数据变体，aVCT4i 增加渐变背景，oUjI1t 增加图片背景并发现共享图片背景删除失败。当前整套因四项变换、七项间距及一项背景删除失败退出 1。重跑不重复累加案例。独立新进程 root/内存 SVG 懒加载检查继续保留；图片的 Agent 检查不算人类校准。

后续应继续复用编译器的组件展开、数据映射、布局和对象关系结果。JS painter 负责显示这些结果；不应另建一套 PPJ 数据集、样式或 OOXML 解析器。SVG 是绘制输出，PNG 由栅格后端生成；这一职责划分不需要另做一个 Office 宿主。

## 3. 已经成功的范围，以及证据到底证明什么

以下构建和实施记录均为此前留下的证据；段落中的“该轮/本轮”按紧邻的报告或提交解释，不指本次文档整理。当前汇总结论见开头，不将不同版本的局部成功拼成一次全量验收。

### 3.1 当前源码和测试能够支持的结论

| 功能 | 内部路线已实现或已有通过证据 | 不能由此推导的结论 |
| --- | --- | --- |
| 基础几何 | 有限预设（新增五种多边形）、literal 自定义路径、二次/三次 Bézier、`arcTo`；形状与文字分别绘制 | 所有 preset、公式引用、连接点、文字区域均已求值和显示 |
| 圆角矩形 | 形状和图片遮罩共享 roundRect 逻辑；默认值、零值、非正方形和调整边界有合成断言 | 所有形状 adjustment 都有对应实现 |
| 文字 | 保留段落/run 边界、显式换行、有限对齐、字号、字体、粗斜体、直接颜色/透明度；单下划线、单删除线、有符号基线及零覆盖 | 完整字体度量、自动换行、列表、AutoFit、双线/波浪装饰或全部文字效果 |
| 组与变换 | 原生 frame/childFrame 映射、旋转、镜像、元素 hidden 分支 | 所有嵌套、导入变换及效果扩展边界均已验收 |
| 图片 | 正裁切、负边值透明留白、alpha；rect/ellipse/diamond/roundRect、五种多边形遮罩及 literal 自定义遮罩；直接 RGB 边框 | tile、全部 fit/focus、主题色、阴影及复杂路径都已支持 |
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

### 3.5 更新原生回归基线与文字源编辑（2026-09-10）

先用旧 runtime-slot-fixed 跑新增断言，报告 `tmp/officekit-native-scene-paint-IcwdcL/integration.json`：锚定 middle/top 和底边距 0 三项修改通过，删除 anchor、bottom inset、两者一起删除三项失败，候选保留原属性。旧快照没有当前源码中的删除分支，不能把这三项判为当前源码缺陷，也不能删除断言回避版本差距。

随后从提交 `c8b0d324561bc9d390f9f71aa108fad05bd9344b` 导出独立快照，3847 个受跟踪普通文件/符号链接的 Git blob 全部一致；一项 legacy submodule 未展开，不参与构建。使用仓库 `scripts/build-office-kit.mjs`、SDK 8.0.128、linux-x64、仓库 TMPDIR，以及关闭共享编译/节点复用的环境设置，生成独立 `tmp/preview-current-native-iv6UMg/runtime`：7 个文件、106076144 字节，未替换安装包。

| 验证身份/范围 | 结果 |
| --- | --- |
| PPJ NativeAOT SHA-256 | `d1a1dd53bbda25e0146058799e86ca718c030cb0a4e4ae45496a1e96536622d6` |
| Office NativeAOT SHA-256 | `1a4b88211b813ba388c06031e7e616d1e1fed4de055b265ffc37472a95e2cb3f` |
| manifest SHA-256 | `7ae2fe9f6d277dae531e497e90d22d4d2eacd3cf69ff0a6ea2612032b6131cfb` |
| 托管 scene + margin 专项 | 52/52、零跳过；`tmp/preview-current-native-iv6UMg/tmp/test-results/preview-current.trx` |
| 托管枚举 body 属性生命周期 | 30/30、零跳过，含五类 owner 的 verticalAlignment；同目录 `body-enum-current.trx` |
| 新包真实 JS 集成 | vAmZ1g：文字源编辑 6/6，四项变换删除仍失败，整套退出 1 |
| Presentation 门禁 | 4/4；正式旧路线产物 `tmp/officekit-ppj-preview-480pKE` |

六个源请求均从原始源重新投影：middle、显式 top、bottom=0、删除 anchor、删除 bottom、同时删除。新包的候选 XML、重新投影及 scene 存在性均符合请求；显式 top/0 保留属性，删除才移除。每例仅 `ppt/slides/slide1.xml` 改变，文件清单及其他 ZIP 成员逐字节不变，原源/请求不变。实际文字区域 raw raster 与独立作者输入一致；删除不再仅凭画面判断。

这轮只增加 JS 回归并重建已存在的原生实现，没有修改生产 C#。nullable/trim/AOT 警告仍保留。没有为该包重做双构建一致性、全仓/完整字段、人类或宿主验收；其证据仍按该轮范围解释。当前默认包已由第 3.6 节后续构建替代。

### 3.6 默认运行时交付与复验（2026-09-10）

从提交 `035472e9ca8f7fd58987148285da5a3f406286d2` 导出 `tmp/preview-default-package-T8N9eZ/snapshot`，核对 3882 个受跟踪普通文件/符号链接的 Git blob；一项未展开的 legacy submodule 不参与构建。先将原安装包完整备份到同级 `previous-package`，再在快照内使用仓库 `npm run build:office-kit -- --output <默认包目录>` 构建 Office 与 PPJ 两个 profile，输出到 `packages/office-kit-codec-linux-x64`。本机旧二进制和 manifest 已被替换，可从备份恢复；package.json、许可证、notices 和 SBOM 内容未变化，没有移除用户输入。

| 项目 | 身份或结果 |
| --- | --- |
| SDK / 平台 / 包版本 / wire | 8.0.128 / linux-x64 / 2.0.0 / 2 |
| manifest 生成文件 | 7 项，106092528 字节；连同 manifest/package.json 打包为 9 项 |
| PPJ 可执行文件 SHA-256 | `eeca85542b5758c567dff712cbf1d37874f9518b7471e42216c597e3fec4bd71` |
| Office 可执行文件 SHA-256 | `a2217c0cc39fd70c5c0e739c75990d7e5e6602bbbed81fabc550ac485b1e4613` |
| manifest SHA-256 | `e95b6748bc86d89ae6a863a4661012d9e387398e6c778fa671dbf99258518662` |
| 快照托管 scene 专项 | `FullyQualifiedName~PpjPreview` 47/47、零跳过；`tmp/preview-default-package-T8N9eZ/test-results/preview-default-package.trx` |
| 双构建一致性 | 快照内 `npm run verify:office-kit-build` 退出 0，两次构建的 9 个文件一致 |
| 默认包门禁 | 无 `NODE_OPTIONS` / `OFFICEKIT_PREVIEW_TEST_RUNTIME` 的 presentation 4/4，背景分支后产物 `VYB5wP` |
| 通用核心与传输 | `test/office-kit.mjs`、`test/office-kit-native-transport.mjs` 通过；前者覆盖 Word/Excel，不是 PPT 全功能验收 |
| 包检查 | `test/package-contents.mjs` 通过；包目录 `npm pack --dry-run --json` 成功，9 项、解包 106094991 字节 |
| 本包完整 scene 集成 | idYOSh、L4nZbV、aVCT4i 均因原十一项删除失败；图片背景 oUjI1t 另发现共享图片背景删除失败，共十二项，整套 failed |

构建使用仓库 TMPDIR，关闭共享编译和 MSBuild 节点复用，没有更改 SDK 锁或屏蔽已有告警。托管 47 项与第 3.5 节的 scene+margin 52 项过滤范围不同，不是少执行五项后宣称同等验证。冻结提交之后的并发原生修改不在此包中；JS、registry 和绑定身份单独记录在集成报告里。

同轮 Skill portability、纯 reference sync 和 Claude 插件门禁通过；完整 `test/reference-skills.mjs` 在 PDF smoke 调用 `pdftoppm` 时因依赖缺失失败。通用 Skill 校验器还拒绝仓库既定的 `name: Presentations` 大写名称；仓库 reference 测试明确要求该名称，因此没有为通过通用校验而改名。这两项不能记为通过，也不是自有 SVG painter 依赖 LibreOffice/Poppler 的证据。没有全仓 `npm test`、其他平台安装、人类校准或 Office 宿主验收。

## 4. 按功能核对剩余差距

### G-01：共用场景与正式入口

**现状：10/15 个任务勾选，正式函数与本机默认包已接场景，整体绘制未完成。** 编译采集、传输、归属与适配已有基础，任务 5.2 固定版本构建、4.2 场景发布和 5.1 成对 fixture 验收通过；3.3/3.4 剩余绘制、4.1 未完成事实映射以及 5.3、5.4 仍开放。这些是任务编号，不是本文章节编号。

剩余工作：按 renderer profile 逐条核对事实规则，特别是 dataset/vector 非折线状态仍有旧事实失败的情况；补足当前字段与复杂组合的绘制/诊断覆盖；后续原生更新继续做匹配包复验。保留固定成对案例，不用“两边都编译成功”或“两边都显示占位”代替等价性。

完成证据：真实 CLI 输出包含正确场景几何/数据/样式；返回值与落盘清单一致；缺场景、资产、栅格及写入失败均有明确结果；旧发布保护不退化。

### G-02：文字和形状

#### 形状内部文字区域：预设定义与自定义坐标

`shapeTextFrame()` 现先确定形状的内部文字矩形，再交给共享文字布局消费边距、段落和锚定。它覆盖当前已绘制的十二个原生预设/别名：rect、textbox、flowChartProcess、roundRect、ellipse、diamond、flowChartDecision、triangle、rtTriangle、trapezoid、parallelogram、chevron。textbox 是文本框的原生标记；不能把十二个标记解释为十二个公开 PPJ preset。

矩形来自[现有预设表所固定版本的定义](https://github.com/plutext/docx4j/blob/0eec2587ab38db5265ce66c12423849c1bea2c60/docx4j-core/src/main/resources/org/docx4j/model/shapes/presetShapeDefinitions.xml)，不是轮廓外接框或手工估计留白。轮廓与文字区域共用调整值缺省和钳制计算；例如，三角形区域随顶点水平移动，平行四边形不能简化为左侧错位量，而 chevron 在凹口达到特定范围时使用定义中的分支。roundRect 使用圆角半径导出的内缩，ellipse 使用内接矩形，diamond/flowChartDecision 使用四分之一边距。

自定义形状的 `textRectangle` 消费形状局部 EMU 边值，与 path viewBox 无关。目前支持 literal 边值和 `l/t/r/b/w/h/hc/vc` 引用；缺省仍是完整形状区域。其他引用、未知预设、未映射调整、冲突或非正区域给出 `preview.scene.paint.text-rectangle`、状态 unavailable；文字可保留在原外框供阅读，但这是明确失败的布局证据，不是正确位置。未知区域的 unavailable 不能被同页 opaque 或 partial 状态覆盖。

方向旋转、独立文字旋转和防镜像使用该文字区域的中心；外层形状/组变换仍作用于整个对象。文字区域的位置随外层变换移动，填充、轮廓和图片裁切不改用文字区域。表格继续使用自身单元格区域。没有修改候选 PPTX、添加通用公式解释器、运行时网络下载或渲染依赖。

本轮验证（默认原生包仍为 035472e9）：

| 证据 | 实际范围与结果 |
| --- | --- |
| 合成 scene | 首个三角形回归先复现 x=17.2/y=59.6 的错误，修复后为 x=92.2/y=119.6；另有 378 个区域/方向/锚定/水平对齐组合、42 个文字旋转与翻转中心断言、9 个失败反例；场景字节不变 |
| 真实作者与源 no-op | p4Au9s 的 `textRegionCases=78`、`textRegionFailures=[]`：11 个公开预设加 literal/built-in 两种自定义矩形，各含三组方向、锚定和组合变换，再分别验证作者与去快照源输入 |
| 独立像素对照 | 参照输入是在相同外层 group 中直接放置一个矩形文本区，不从 painter 输出取坐标；两行 F0/IL 的边界、重心、双向 2px 墨迹邻域及面积比例通过，橙色兄弟形状保持原位 |
| 正式入口与源身份 | 78 项均 requires-review；强制输入检查、整页内容像素、scene/candidate 摘要、返回/落盘清单、非覆盖保护与输入保留通过；39 份去快照源 no-op 与原始源逐字节一致 |
| 调整值源编辑 | 原 66 项多边形继续执行。自定义参照现在具有独立计算的相应文字矩形，整页内容仍逐像素相同；36 项原源调整值设置/清零/删除的非目标保留断言未削弱 |
| 外部对照 | `tmp/text-region-probe-CzGoIx/comparison.json` 为 8/8，`tmp/text-region-presets-kfb0wR/comparison.json` 为 39/39；实际 PPTX 经 LibreOffice 26.8.0.3 → PDF → 公开 MuPDF 72dpi，与正式本地 preview 的文字边界相比，预设上限 3px，实测最大 2px |

p4Au9s 的完整报告时间为北京时间 2026-09-10 17:21:21，正式入口 409 项。状态仍是 failed/退出 1：原四项变换、七项间距及一项共享图片背景删除失败保持不变；新增区域和其他回归失败数组为空。无 preload 的 presentation 门禁为 4/4，产物 `tmp/officekit-ppj-preview-054XQA`。此轮未重建原生包，也不声称并行 C# 修改已验收。

中间失败记录保留：8dkVVf 的参照 group 误用了 `children` 而非公开 PPJ 的 `elements`；Il5HE3/u2wTMh 将旧 24pt 旋转测试的固定 100 像素下限套用于 14pt 文本。最终仍使用原两行 F0/IL，以字号平方归一化墨迹密度下限，旧 24pt 案例仍要求超过 100 像素；边界、重心、逐点邻域和面积比例检查未放宽。正式诊断回归另更新了未知预设文字区域应为 unavailable 的预期，仍检查两个几何事实错误、红色警示、重复 ID 和独立 opaque 兄弟节点。

最终核对包含 17 份实现/测试/绑定/fixture/预设表、两份原生二进制、manifest、全部 66 项多边形及 78 项区域候选/源/参照/请求、78 份正式输出对应图像与两组外部文件，共 577 个唯一文件摘要一致。78 份正式清单的 scene/candidate 身份和 requires-review 状态另逐项核对。181 个本地文档链接/锚点、生成检查、gate-policy 和严格 OpenSpec 校验通过。此前 pINJDG 已有相同 409 项结果，但结束后共享 registry 新增了并行 bullet startAt 声明；p4Au9s 在保留该条目和更新本轮文字区域说明后重跑，不把并行原生实现算成本包验证，也不把重跑累加为新案例。

这些自生成对照不是第三方 fixture、PowerPoint 或人类校准。已查看八格本地及 LibreOffice 图片，但外部自动断言只证明上述文字边界，不是整页像素保真。完整换行、字体度量、AutoFit、其他预设与 guide/formula 求值、主题继承和效果仍开放，G-02 和任务 3.3 不勾选完成。

#### 五种多边形：共享轮廓与尚未解决的文字区域

本小节保留文字区域修复前的轮廓阶段证据；文字区域的当前限定验证见上一小节。

共享 painter 已实际绘制 `triangle`、`rtTriangle`、`trapezoid`、`parallelogram`、`chevron`，覆盖普通形状、形状图片填充和独立图片遮罩。实现位于 [presetPolygon](../src/ppj/preview-scene-svg.mjs)，默认调整值复用[现有预设表](../src/ppj/preset-geometry-profiles.json)，数学关系来自该表固定版本的[公开预设定义](https://github.com/plutext/docx4j/blob/0eec2587ab38db5265ce66c12423849c1bea2c60/docx4j-core/src/main/resources/org/docx4j/model/shapes/presetShapeDefinitions.xml)。运行时没有下载定义、解析 OOXML 或新增通用公式解释器。

| 预设 | 轮廓调整的含义 | 缺省值与钳制 |
| --- | --- | --- |
| triangle | 顶点沿整个宽度水平移动 | 50000；限制到 0～100000 |
| rtTriangle | 左上、左下、右下组成直角三角形 | 无调整槽；额外调整值明确失败 |
| trapezoid | 两个上顶点相对左右边的内缩 | 25000；按短边缩放，内缩不超过宽度一半 |
| parallelogram | 上左、下右顶点的水平错位 | 25000；按短边缩放，错位不超过宽度 |
| chevron | 尾部凹口和头部三角形的水平深度 | 50000；按短边缩放，深度不超过宽度 |

宽高不相等时不能把短边比例替换成宽度比例。显式零、缺省和允许的有符号输入分别保留在原生状态里，只有显示时做定义要求的钳制；非整数先被 wire 拒绝，越出协议范围或多余调整槽在对应字段报不可用。图片与轮廓共用路径，负裁切留白、素材透明带、半透明填充及独立边框不被铺成实色矩形。已完成路径进入现有几何事实映射，缺失绘制记录仍不能解除旧错误；文字、继承和其他限制保持独立。

测试证据：

- 合成专项含 30 个固定坐标/三消费者断言、48 个宽高比与钳制对照、26 个非法调整绘制反例、8 个小数 wire 拒绝，以及四个缺省/显式连接方式诊断断言。首个三角形断言修复前失败；输入场景保持不变。小数反例按实际发生的 wire 层统计，不称为 painter 拒绝。
- TzaE9j 的 `polygonPresetCases=66`、`polygonPresetFailures=[]`：15 个作者、15 个去快照源 no-op、36 个从原始源分别设置 25000、设置零、删除调整值的编辑。三类消费者均与独立 literal 自定义路径的整页内容像素逐字节相同；chevron 同时包含 15° 旋转和水平翻转。完整正式入口另核对强制检查、身份、清单与非覆盖保护。
- 所有 36 项编辑重新投影正确，零与缺省不混淆，目标原生对象除调整值外的全部字段保持一致；只有 `ppt/slides/slide1.xml` 改变，ZIP 清单、其他成员和原始输入不变。候选、源、请求、重新投影文件及摘要保留在报告目录。17 份 JS/测试/绑定/fixture/预设表与两份已安装原生二进制组成该轮身份边界，不代表最新工作区 C# 已重建。

外部对照保留在 `tmp/polygon-preset-probe-9VUZDo/`：五份作者 PPTX 由 LibreOffice 26.8.0.3 转 PDF，经公开 MuPDF 路线以 72dpi 生成检查图。比较只提取洋红色边框，预先要求双向 2px 邻域、四边差不超过 2px、墨迹面积差小于 15%。它不检查文字位置，也不证明图片透明度或完整视觉保真。

缺省连接的 `comparison.json` 为 **failed**：三角形左边缘差 4px，chevron 右边缘差 3px；其他三个案例通过。这些 PPTX 没有直接 join，当前 SVG 的 miter 审查近似与外部继承结果不同。因此可见笔画缺少直接连接方式时，新增 `preview.scene.paint.line-join-inherited` 字段诊断，而不是默认宣称角点范围正确。原反例和输入保留，未放宽阈值。

源编辑保真检查进一步比较同一 slide 的完整 XML，仅遮去第一个目标预设的调整列表。中间 RHLSSi、xqJQDD 报告暴露了重复命名空间位置差异：图片编辑可在根部补上同 URI 的 r 声明，并把 srcRect 的 a 声明提升到根部。最终比较先证明所有 r 属性已有相同局部绑定、图片裁切的 a 绑定来自相同局部或根部 URI，再仅规范化这两处及既有属性顺序；任意其他内容保持比较。修改兄弟形状颜色或关系命名空间的反例仍会失败。TzaE9j 的全部 36 项通过此加强检查；172 个相关唯一文件（实现与二进制、候选、源、请求和重新投影）摘要核对一致。中间失败不算新的原生功能失败，也未通过忽略整个 slide 消除。

`round/` 是五份独立新输入，只显式指定 `stroke.join=round`，通过正式 CLI build/preview 后重新外部对照：5/5 通过，四边差均为 0px、每个边框像素都有 2px 内对应点，墨迹面积差最大约 3.18%。它支持直接 round 与本轮轮廓计算的结论，不能覆盖缺省连接失败；两个报告分别记录结果。比较脚本是收集器，命令退出成功表示写完报告，实际验收状态以 JSON 的 status/failures 为准。

**轮廓阶段曾复现文字区域缺陷。** 当时 LibreOffice 把 F0 放在预设内部，本地却按整个外框排版，文字落到轮廓外。上节已补充本轮修复和独立验证；此处的旧轮廓正例本身仍不能证明文字正确，也不证明任意预设、主题连接继承、连接点、阴影边缘、复杂自定义公式、PowerPoint 或人类验收。

#### 文字方向：阅读坐标与物理边距

新增反例：显式 Liberation Sans、24pt、缺省边距的 vertical/vertical270 在文字阴影外部对照中出现 3～4px 的字形位置差，超过该组预设 3px 阈值。字形和阴影同步偏移，两侧阴影相对字形均准确右移 40px；这证明局部投影位移，不能证明竖排文本定位正确。具体 PNG、边界和失败报告见下方 G-04 的“形状文字外阴影”。后续应核对竖排缺省边距与文本锚点，不能用此前其他文字 profile 的成功关闭此反例。

[原生文本体 codec](../native/OfficeKit/src/OfficeKit.Codec/PptxBodyPropertiesCodec.cs) 支持的 `horizontal`、`vertical`、`vertical270` 现由共享文字 painter 实际消费。后两者分别对应整段文字顺时针/逆时针 90° 阅读，不能等同于 WordArt 逐字堆叠或东亚竖排；原生接口也未承诺这些其他模式。方向先决定逻辑排版区域，再完成段落/显式换行/锚定，之后依次施加方向、独立文字角度、镜像补偿和外层形状/组变换，填充和轮廓仍使用原 frame。

以物理边距 L/T/R/B 表示左/上/右/下，三种模式使用以下阅读坐标：

| 模式 | 排版宽、高 | 阅读坐标的左、上、右、下边距 | 方向角度 |
| --- | --- | --- | --- |
| horizontal 或缺省 | W、H | L、T、R、B | 0° |
| vertical | H、W | T、R、B、L | +90° |
| vertical270 | H、W | B、L、T、R | −90° |

交换宽高时保持 frame 中心不变。边距仍属于物理边，不会把原左边距错误地当作所有方向的行首；明确的零边距也保留。锚定在阅读方向的行块坐标中计算，顺时针竖排的 top 因而靠物理右侧，逆时针靠左侧。属性边界与默认值见 [Microsoft 的方向说明](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.vertical?view=openxml-3.0.1)、[左边距说明](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.leftinset?view=openxml-3.0.1)和[方向相关锚定说明](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oodf13/d62e2585-7cba-49d9-83d7-afe7d201b339)。独立角度不覆盖方向角度，这也对应 [LibreOffice 区分两种文字旋转的修复](https://cgit.freedesktop.org/libreoffice/core/commit/?id=7e23cbdbb6ec0247a29ed8a8f744c01e10963ea0)。

合成测试先复现方向被忽略，再覆盖 72 项方向/锚定/段落对齐/缺省或零或不对称边距组合、16 项独立角度及翻转组合；六项未知方向、upright、扭曲、多列或非正可用区域明确 unavailable。原生场景序列化保持不变。显式 horizontal 被消费，不再误报该字段未绘制；删除仍以字段缺省表示，不能用 horizontal 代替。所有字体度量、换行、AutoFit 等剩余限制继续保留。

文字方向阶段最终 xTtqoh 的 `textDirectionCases=75`、`textDirectionFailures=[]`：

| 输入与操作 | 数量 | 验证范围 |
| --- | --- | --- |
| 文本框、带文字形状、单元格 × 三种方向 × top/middle/bottom，作者和去快照源 no-op | 54 | 原生方向、两行 F0 / IL 的实际墨迹、物理边距/阅读区域、完全相同的 no-op 源字节 |
| 三类 owner × 两种竖排与正负/小数独立文字角度、外层 90° 和水平翻转组合，作者和源 no-op | 12 | 方向、角度与防镜像的实际组合；不把简单角度替换当成排版 |
| 三类 owner 分别从原始 vertical/middle 源独立改为 vertical270、明确 horizontal、删除方向 | 9 | 真实候选像素、重新投影、vert/horz/vert270/属性缺省、全部段落/run/样式/换行及非目标内容保留 |

实际像素与独立作者输入比较：参考输入使用已明确交换的 frame、重排后的物理边距和相应横排文字角度，不读取 painter 的 SVG 来构造预期。九项源编辑只改变 `ppt/slides/slide1.xml` 中目标 `bodyPr vert`；沿用已证明范围的属性顺序/重复命名空间规范化，改动文字仍会被检出。所有正式结果为 requires-review；整套仍因原十二项删除失败退出 1。

本轮还做了一个独立对照：`tmp/vertical-direction-probe-UR41ql/` 中的九格自生成 PPJ，经正式 `officekit ppj build` 生成 PPTX，再由本机 LibreOffice 26.8.0.3 转 PDF；用 `officekit run` 和公开的 `office-kit/pdf/mupdf` 生成检查图。相同 PPJ 的正式 `ppj preview` 与外部图均以 72dpi 检查。`comparison.json` 保留输入/PPTX/PDF/两张 PNG/清单/比较脚本摘要，九个两行蓝色文字区域的四边差最大均为 1px（预设检查上限 3px），行列方向、锚定和不对称边距一致。

这个对照仅包含普通矩形、20pt 字号、两行文字和三种锚定，没有验证外部旋转/翻转、表格、自动换行、复杂字体或任意文稿。LibreOffice 与本地字形仍不同；1px 结果不能解释为完整字体保真，也不等于 PowerPoint 或人类验收。它是 OfficeKit 自生成输入的外部解释器对照，不增加第三方 fixture 数量。LibreOffice/MuPDF 没有加入本地 SVG 路线的必需依赖；普通内存绘制的懒加载门禁仍需通过。

#### 文字翻转：抵消镜像但保留外层变换

以下为 LBG93d 阶段的记录；当时竖排未实现，现有三种原生方向的进展以上节为准。

此前 painter 将形状或组的翻转直接作用于整棵 SVG 子树，文字也随之镜像。旧 LU0TNH 的 47 项旋转测试中，四项带水平翻转的组合把镜像墨迹作为预期，因此旧通过结果不能证明文字方向正确。本轮先增加会在旧实现失败的断言，再纠正绘制及这些期望，没有删除组合案例。

现在沿祖先组累计水平/垂直翻转；累计反射次数为奇数时，在当前文字 frame 中心增加水平补偿，包在文字自身旋转之外。原有外层形状/组变换仍作用于位置、填充和轮廓。这使字形不再镜像，但不会强制文字永远朝上：例如单独垂直翻转仍可能得到旋转 180° 的字形。该处理与 [Apache POI 的 DrawTextShape](https://raw.githubusercontent.com/apache/poi/trunk/poi/src/main/java/org/apache/poi/sl/draw/DrawTextShape.java) 对祖先翻转及文字旋转的处理一致；这里使用现有原生场景自行计算，没有引入 POI 或第二个文件解析器。

[合成回归](../test/ppj-preview-scene-svg.mjs) 新增 192 种组合：形状、单元格、已验证的作者 drawing cache，各自穷举 owner 和两层组的水平/垂直翻转。组还包含旋转、非零 childFrame 原点和非等比缩放；另一个顶层兄弟保持不翻转，防止遍历下标误作继承状态。32 项文字角度/锚定测试检查补偿位于文字旋转外层。所有输入的原生序列化字节保持不变；未解析字段和 `text-layout` 限制仍保留。

真实 LBG93d 报告新增 `textReflectionCases=36`、`textReflectionFailures=[]`：

| 输入与操作 | 数量 | 实际检查 |
| --- | --- | --- |
| 形状/单元格各八种自身、单组及嵌套翻转，作者和去快照源 no-op | 32 | 原生变换与组坐标、补偿次数、不对称两行 F0 / IL 的真实像素、原源与 no-op 完全相同 |
| 两类 owner 的嵌套旋转及非等比缩放输入，分别从原始源将文字角度 90° 改为 −90° | 4 | 候选墨迹、重新投影、全部段落/run/样式/换行不变，整页只允许目标 bodyPr rot 变化，其余 ZIP 部件逐字节不变 |

像素参考由零角度实际字形和测试独立的坐标运算生成，不读取 SVG 变换来推导预期。检查边界 ≤2px、重心 <1.5px、双向 2px 墨迹邻域及按缩放面积修正后的墨迹数差 <20%；这些容差用于栅格 hinting，不是 Office 保真误差承诺。正式入口的完整内容区像素、强制检查及候选身份另行验证。36 项均为 requires-review，原 47 项旋转回归同时通过。

首次 wGyVFw 在保存源 no-op 候选时与源文件重名，非覆盖保护返回 EEXIST；原源没有被覆盖。现在源与 candidate 明确分开命名，该次失败报告保留。四项源编辑复用已有、限定范围的 XML 属性排序/重复命名空间规范化，修改文字的反例仍被检出。

仍缺完整 upright/竖排/扭曲排版、真实字体度量和自定义文字区域；原生图表的独立轴标签等未全部使用共享文字 painter，不能据此宣称所有图表文字翻转已修复。drawing cache 在此只有合成绘制证据，不改变导入 SmartArt unavailable 的结论。原十二项源删除失败继续使整套退出 1；未重建原生包、未做人类或 Office 验收。

#### 文字自身旋转：独立于形状的变换

以下为 LU0TNH 阶段的记录。其四项镜像期望已由上节纠正，当前 47 项回归以补偿后的实现为准。

[共享文字 painter](../src/ppj/preview-scene-svg.mjs) 已消费 `bodyProperties.rotationAngle60000`，按有符号角度绕当前文字 frame 中心旋转。它先完成已有段落/锚定布局，再旋转文字块，最后叠加外层形状/组变换；填充和轮廓不随文字自身旋转。角度范围沿用原生 ±360°，显式 0 保留独立旋转节点，字段缺省不产生该节点。旋转后没有擅自缩放或裁掉超出形状的文字。`bodyPr/@rot` 与形状变换独立的语义见 [Microsoft 的属性说明](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.rotation?view=openxml-3.0.1)；当前 frame 中心与逻辑文字块的组合是本地 review 绘制，尚非复杂自定义文字区域或宿主排版验收。

真实报告 `tmp/officekit-native-scene-paint-LU0TNH/integration.json` 的 `textRotationCases` 为 47 项，`textRotationFailures=[]`：

| 输入与操作 | 数量 | 已验证内容 |
| --- | --- | --- |
| 文本框、带文字形状、单元格各 5 个角度：−90、0、90、180、12.25°，作者与去快照源 no-op | 30 | 原生角度、真实文字墨迹、方向/中心、源 no-op 字节不变 |
| 文本框/形状各 2 个组合：外形 90°＋文字 −90°，外形 90°＋水平翻转＋文字 90°；作者与源 no-op | 8 | 文字变换的先后次序、方向、独立前景保留 |
| 三类 owner 各从原始 90° 源文件独立改为 −90°、设为 0、删除字段 | 9 | 实际候选绘制、原生与 XML 存在性、重新投影及非目标部件保留 |

像素断言使用不对称的两行 `F0` / `IL`，以零角度墨迹为参考，独立计算旋转/翻转后的坐标。核对边界（≤2px）、重心（<1.5px）、双向 2px 墨迹邻域与像素数差（<20%，容纳栅格字体 hinting）；这不是对 PowerPoint 的误差承诺。形状/单元格填充及独立橙色前景保持原位，正式入口内容区与内部结果逐像素相同。九个候选只改变 `ppt/slides/slide1.xml`，请求、候选、重新投影和原源摘要均有记录；删除确实移除 `bodyPr rot`，不能用显式零替代。

合成回归先复现旋转被忽略，再覆盖 ±360°、正负小数、显式零、四种锚定状态和外层旋转/翻转。非法角度以及非零旋转与 upright、竖排、文字扭曲的未解析组合明确 unavailable；相关字段及全局 `text-layout` 限制保留。文字换行、字体度量、复杂形状文字范围、溢出、完整竖排、upright 与宿主的形状/文字翻转差异仍需后续实现和验收。

本轮还复现了**表格 run 内换行的导入差距**：最初 GRB5f4 的单个 run 为 `F0\nIL`，作者可绘制，去快照导入却成为 opaque；测试最初错误地继续读取 table.rows，因而产生五个访问异常。这不是元素顺序改变。[PptxTableCodec 的 TryReadCell](../native/OfficeKit/src/OfficeKit.Codec/PptxTableCodec.cs) 接受独立 `A.Break`，但拒绝 `A.Text` 内含换行。旋转生命周期输入改用两个文字 run 加独立 break，仍保留两行和原文字；原写法的 0°/90° 两例单列为 `textRotationOpaqueCases`，断言原源/no-op 字节不变、明确 opaque 占位且不伪造旋转绘制。它们证明保守保留，不证明表格导入或视觉缺口已修复。若要统一两种换行写法，需要扩充原生导入/编辑拓扑契约，不能由 painter 扁平化绕过。

最终 LU0TNH 进一步逐个比较九个源编辑的全部原生段落/run/换行/样式，并核对目标 slide 除 `bodyPr rot` 外的内容。比较仅处理 SDK 的属性排序和已证明与祖先一致的重复命名空间声明；人为修改文字或命名空间的反例仍能被发现。此前 2R6pAt 的三项表格严格比较失败来自 bodyPr/lstStyle/p 增加相同命名空间声明，不是文字拓扑变化；这些中间失败报告保留。

47 项正式结果均为 requires-review；既有十二项源删除失败仍使整套退出 1。原生包未重建，未执行全仓、Office 或人类验收。实际旋转 PNG 已由 Agent 查看，不计为人类校准。

#### 其他文字与几何进展

字符项目符号已有有界绘制（2026-09-10）：直接 bulletCharacter、字体、RGB、point 字号和左对齐负缩进齐全时，符号绘制在 paragraphLeft+indent，正文各行从 paragraphLeft 开始，显式续行不重复符号。保留直接 alpha；noBullet=true 不绘制符号。字体度量及符号/正文间隙仍需 review；缺字体、主题颜色、百分比字号、无悬挂空间和非左对齐目前明确 unavailable，不猜继承或 tab 布局。自动编号、图片符号、完整列表继承和删除仍未完成。

真实 `tmp/officekit-native-scene-paint-BPCuPP/integration.json` 的 characterBullets 记录作者、去快照源 no-op、源 leaf 修改及独立作者对照。●/CC5500 的实色符号像素为 155，改为 ■/0066CC 后为 210；正文 x=137.2、符号 x≈117.2，两行正文不随字符修改移位，续行区域无符号像素。作者/no-op 和源修改/对照的文字区域逐像素相等，重新投影保留新字符，仅 slide1.xml 变化，原源和作者输入不变。已查看修改后的 PNG（Agent 检查，非人类校准）。合成正例先失败后通过，并保留五种不确定布局/样式反例；Presentation 4/4 和生成检查通过。沿用第 3.5 节包，未重建或切换正式入口，完整集成仍有十一项删除失败、退出 1。

大小写显示已有有界映射（2026-09-10）：内部 `fontCaps=all` 只转换绘制文字，`none` 保留原字符；直接 run 覆盖段落默认，再使用表格 fallback。显式语言参与 Unicode 大写转换，语言/shaping 的原有限制不因此解除。`small` 和未知 token 明确 unavailable，不伪造小型大写字形。合成覆盖默认/直接覆盖、原 scene 不变、土耳其语 i/ı、ß 展开、换行和 XML 转义；新增正例先失败后通过。

沿用第 3.5 节包的 `tmp/officekit-native-scene-paint-je2i3l/integration.json` 记录 `capitalizationCases` 五例：all/none × 作者/去快照源 no-op，以及 all→none 源 leaf 编辑。实际文字区域与独立显式文字作者结果逐像素相等；原生文字和重新投影仍为 `Hello abc`，仅显示改变。候选编辑仅 slide1.xml 变化，原源/输入不变。未重建原生包，未验证完整小型大写、大小写属性删除或正式入口；完整集成仍因四项 frame 删除与七项间距删除失败退出 1。

源段落间距删除已形成失败回归（2026-09-10）：最终报告 `tmp/officekit-native-scene-paint-rsK9nS/integration.json` 中 `spacingDeletionFailures` 为七项、`spacingDeletions.cases` 为零。每个请求都重新投影同一去快照原源，删除 lineSpacingMultiplier、spaceBeforeMultiplier、spaceAfterMultiplier 的七种非空组合；全部在编译阶段报 `ppj.source.unsupportedMutation`，未产生候选，不能声称 XML 删除、重新投影或像素通过。目录保留原源及七份 request.ppj；原源摘要不变，独立回归继续执行，最后退出 1。先前 7p9fqz 同样失败，但未保存逐请求文件，由本次补足。

源码边界：`PptxParagraphSpacingCodec.Apply` 已支持三个 No*Spacing 原生删除 oneof，`PptxCodecTests.ParagraphSpacingAuthorsImportsEditsAndDeletesWithoutChangingUnits` 是底层契约测试；它不能证明 PPJ 路径可用。当前 `PpjPresentationCompiler.ApplyRichText` 的 `MaskTextValues` 允许部分已实现样式生命周期，却没有段落间距删除映射，先以 rich-text topology or styling change 拒绝。因此需要补 PPJ 能力声明、差异识别和原生删除映射，并继续验证单项/组合删除与其他字段保留；不应放宽所有富文本变更、把删除改成零或直接用底层 codec 绕开 PPJ。该原生写出修复超出现有只读 scene 变更，待规划范围确认。整套现为原有四项 frame 删除失败加七项段落间距拒绝；Presentation 4/4 通过不覆盖这些真实原生失败。无新原生包或生产修复，10/15 不变。

倍数段前/段后间距已有内部显示（2026-09-10）：`spaceBeforeMultiplier` / `spaceAfterMultiplier` 分别使用首行/末行的有效逻辑高度，显式零保留，不使用被 run 覆盖的默认字号或整段总高度。该高度仍为简化 review 度量，字体精确排版限制保留。合成测试覆盖 0/0.5/1/2、负数与无穷拒绝、原始 scene 不变，以及 20pt/10pt 混排覆盖 80pt 默认字号。新增测试先失败后通过；追加混排测试一度缺少 schema import，补齐后通过。

真实 `tmp/officekit-native-scene-paint-HIDS0u/integration.json` 的 `internalPainting.paragraphSpacing.paragraphMultiplierCases` 有十例：10/20pt × 0/0.5 倍 × 作者/去快照源 no-op，共八例；另有两种字号分别将源段前/段后从 0.5 改为 0 的编辑。三段墨迹起点与独立 point 作者对照一致，源候选重新投影保留显式零，仅 slide1.xml 改变，原源/输入不变。沿用第 3.5 节原生包，未重建或切换正式入口；整套仍有四项变换删除失败，退出 1。继承、删除恢复、字体精确度量及完整 G-02 仍开放。

倍数行距已有内部显示（2026-09-10）：`lineSpacingMultiplier` 乘以现有逻辑基线步长，首行位置不变；固定 point 行距和未指定行距的行为保留。它不是字体真实行高度量，`text-layout` 限制继续存在。合成测试覆盖 0.5/1/1.5/2 倍、输入不变，以及零、负数、无穷值明确失败，新增反例在修复前确实失败。

真实报告 `tmp/officekit-native-scene-paint-0PGFXu/integration.json` 沿用第 3.5 节 c8b0d324 包，`internalPainting.paragraphSpacing.multiplierCases` 记录七例：1/1.5/2 倍 × 作者/去快照源 no-op，以及源 1→2 倍编辑。20pt 三行文字的实际墨迹起点间隔分别为 24/36/48px，与独立 point 行距作者对照一致；源编辑重新投影为 2 倍，仅 slide1.xml 变化，原源和作者输入不变。Presentation 4/4 与两个生成检查通过；完整集成仍有四个原生变换删除失败，退出 1。未重建原生包、切换正式入口或提升完整文字等级；倍数段前/段后间距、删除继承及字体精确度量仍待补齐。

显式垂直对齐已有内部显示：top/center/bottom 沿用现有逐行排版，将整段文字块放入上下 inset 界定的可用高度；bottom inset 在该分支实际消费，显式零保留。九项合成正例覆盖三种 anchor、底边距 0/10/30 和两行文字；三项反例保留未知 anchor、空间不足和边距越界的字段级 unavailable。未显式指定 anchor 时不猜继承值，旧排版和未消费字段诊断保留。

真实报告 `tmp/officekit-native-scene-paint-6r8OMm/integration.json` 的 textAnchors 共 12 例通过：作者与去快照源 no-op、三种 anchor、底边距 10/30。40pt 字形在底边距 10 时，居中/底对齐相对 top 下移 16/32px；底边距 30 时为 6/12px，墨迹数均为 911，源 no-op 与作者一致。原始程序和源字节不变，候选 no-op 等于源。使用既有 runtime-slot-fixed，未重建当前 C#；Presentation 4/4 通过，完整集成仍因四项源变换删除失败退出 1。

该映射使用现有简化行高和逻辑下伸空间，并非按字体真实墨迹求出文字块高度；`text-layout` 限制继续存在。限定源 anchor/底边距修改与删除随后已有第 3.5 节的新包证据。字体度量、自动换行、AutoFit、分布式锚定、center/bottom 溢出位置、anchorCenter、复杂继承及完整垂直文字仍未完成。不提升完整文字支持等级或切换正式 CLI。

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

剩余工作包括字体度量和替代、中文/混排换行、段落间距的继承删除语义、列表、垂直布局、AutoFit、溢出、双线/波浪等装饰格式，以及剩余 preset 与 adjustment。custom geometry 中的 guides、公式引用、text rectangle、独立路径 viewport 等也应分别核对保留、求值和绘制状态；当前 literal 路径函数会拒绝未解析引用。

风险：文字可能溢出、遮挡、错位，节点外形可能不能表达其含义。最小收口案例应包含带文字非矩形、同段混合格式、长中文、显式零值和复杂路径；同时断言实际边界与未支持字段的诊断。

### G-03：变换与可见性

第 3.5 节的新包 vAmZ1g 再次复现下述四项 rotation/flip 删除失败，因此它们不再仅是旧 runtime-slot-fixed 的失败记录。当前快照已包含文字 body 属性删除实现，但不能据此推导 frame 变换删除也已修复。

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

#### 形状文字外阴影：字形投影与仍失败的外部定位

最终 `tmp/officekit-native-scene-paint-qCO4bS/integration.json`（北京时间 2026-09-10 18:15:52）使用冻结 JS 快照 `tmp/text-shadow-snapshot-OenB3z/` 和原有 035472e9 默认原生包。新增 `textShadowCases=34`、`textShadowFailures=[]`；原 30 项基础阴影及其他独立回归继续通过。正式入口合计 473 个不同案例，不累加重复运行；原四项变换、七项间距、一项共享图片背景删除仍使整套 `failed/exit1`。

这次补的是**形状自身外阴影包含其可见文字**，不是 run/段落各自的阴影效果。绘制内容只定义一次：滤镜引用同一份内容产生投影，另一份不经滤镜的引用绘制原对象。投影位移放在滤镜外，滤镜范围使用包含字形的对象范围并为线宽和模糊留边；不另建字体测量引擎，也不让阴影滤镜裁掉原文字。横向/纵向溢出、上标越框、大外框内短文字均有实际像素验证，但本地文字如何换行、使用哪些字体仍受 G-02 限制。

| 新增验证 | 内容与判定 |
| --- | --- |
| 15 个 profile × 作者/去私有快照源 no-op | 普通文字、横向和纵向溢出、上标越框、文字自身旋转、两个竖排方向、形状旋转开关、翻转、形状填充加文字、8pt 模糊、400×200 外框中的 16pt 模糊短文字、模糊溢出、显式字体，共 30 项 |
| 原源独立编辑 4 项 | 改阴影色、透明度归零、删除形状阴影、F0 改 IL；每次从同一原始投影发起。新候选重新投影正确，仅 slide1.xml 改变，其他 ZIP 成员与原输入保留 |
| 独立像素参照 | 另行编译不带阴影的 PPJ，得到文字/形状 alpha，再计算阴影合成；不从待测滤镜提取期望。非模糊最大通道差 1，模糊三组为 6/3/5，分别低于原定 2/8 阈值；参照 PPJ/SVG/PNG 摘要纳入结果 |
| 正式入口 | 34 项均检查整个内容栅格、无关橙色对象、候选/scene 身份、强制输入检查、发布与输入保留；可靠性继续是 requires-review |

首次 v5TT1w 有 11 个源 no-op 阴影被 `language: en-US` 的未映射提示拦住，另一个模糊案例因测试把 sharp 的 RGB 输出当单通道读取而失败。检查现按原生字段路径区分语言元数据，保留语言和文本排版提示；实际文字使用的字体与东亚字体完全同名时，也保留提示但允许本地字形投影。不同字体、未知后代、未解析几何、独立文字效果仍拦截。模糊参照显式输出灰度并断言通道数与长度；没有放宽容差。中间 rW0bLp 的 28 项通过后，才增加后三组边界案例到最终 34 项。合成测试另覆盖 run/default 的三种语言、同名/不同名字体、未知字段和独立文字阴影，以及 run/default/bullet 的部分透明度拒绝。

外部对照由 public `office-kit` / `officekit run` 调用 CLI build/preview，再经 LibreOffice → PDF → MuPDF 72dpi；最终在 `tmp/text-shadow-probe-Gsc1V4/run-cAiEmX/comparison.json`。九项直接阴影使用事先写入的 3px 边界阈值：普通、多行、填充矩形、填充三角形、文字旋转、外框旋转和翻转七项通过，最大差 1px；两个竖排失败，最大差 4px。顺时针竖排的 host/local 字形左上角分别为 (881,319)/(878,315)，阴影分别为 (921,319)/(918,315)；逆时针分别为 (61,615)/(64,619)，阴影为 (101,615)/(104,619)。两侧均右移 40px，字形本身的位置差仍待修复。整份外部报告保持 failed，不把七项成功覆盖两个失败。

其余三项仅作观察：两种模糊的边界差至多 2px，未做完整宿主核/边缘验收；窄文本框在 LibreOffice 自动换行，本地仍横向溢出，差异明显。uCAC59 首次显式字体导致阴影全被保守省略，6pITxk 修正后得到七过两失败；最终 cAiEmX 重复验证同样结果，旧输入和输出均保留。Agent 已查看图片，不是 PowerPoint 或人类校准。

仍缺：独立 run/段落阴影、部分绘制透明度与效果 alpha 的组合、主题阴影、复合效果、阴影缩放/斜切和其他 owner；字体排版、两个新增竖排反例、完整模糊保真也未关闭。`shadow-text-layout`、`shadow-blur-approximation` 等诊断继续保留。此进展不关闭 G-02/G-04 或整个渲染器目标。

收尾检查：1,516 个唯一源码/二进制/输入/候选/参照/发布/外部文件摘要核对通过，包含 473 项正式发布；17 项证据源码在冻结快照与当前工作区一致。无 preload 的 presentation 4/4（`/tmp/officekit-ppj-preview-PCkq4s`）、SVG foundations、两个生成检查、gate-policy、184 项本地文档链接/锚点、严格 OpenSpec 与 `git diff --check` 通过。完整 npm test、其他平台、PowerPoint 和人类验收未执行；未提交或推送本轮修改，任务仍为 10/15。

#### 外阴影：实际轮廓、透明像素和保守边界

本小节是前一轮基础形状/图片阴影的 30 项基线。下述“可见文字组合未覆盖”和文字省略结果描述该轮版本；当前文字进展及仍存边界以上一小节为准。

最终报告 `tmp/officekit-native-scene-paint-gKp0xm/integration.json`（北京时间 2026-09-10 17:49:40）使用源码快照 `tmp/shadow-snapshot-jbQYNe/` 和原有 035472e9 默认原生包。12 种方向/透明度/模糊/旋转/图片裁切 profile 各有作者与去快照源 no-op，再加形状、图片各自从原源独立改色、归零、删除，共 30 项通过；正式入口累计是 439 个不同案例，不累计重复运行。图片含 alpha=0/128/255 的固定区域，实际阴影 RGBA、原对象、无关控制对象、候选/scene 身份、强制输入检查和正式发布内容像素均检查。六项编辑重新投影正确，只修改 slide1.xml，其他 ZIP 成员逐字节保留。原四项变换、七项间距、一项共享背景删除仍失败，其余失败数组为空。

复验修复了把源形状空 txBody 误判为可见文字的问题；图片源编辑 fixture 也改为投影实际使用的顶层 shadow，而非作者样式的 style.shadow。uYY70W 因运行期间共享 registry 改变而未通过身份检查，没有计为成功报告；随后在独立快照完成 gKp0xm。收尾时 17 个证据源码摘要与当前工作区一致，两个原生可执行文件及 manifest 未变。检查了 1,297 个唯一源码/二进制/输入/候选/参考/发布产物文件摘要，包含全部 439 项正式发布和新增阴影源文件；无 preload 的 presentation 4/4（WzzjxR）、SVG foundations、两个生成检查、gate-policy 和严格 OpenSpec 通过。完整 npm test、其他平台及人类验收未执行；任务仍 10/15。

共享 [`outerShadow()`](../src/ppj/preview-scene-svg.mjs) 已消费普通形状与图片的直接外阴影，作用于已绘制结果的 `SourceAlpha`，不另画矩形代替主体。图片透明像素、裁切留白及几何遮罩参与轮廓；阴影在原对象下方合成。不改变原生场景、PPTX 或编辑权限，也没有新增图形依赖。

直接 RGB、透明度、距离、方向及显式零已有绘制分支。`rotateWithShape=false` 将偏移向量逆变换到对象局部坐标，使最终偏移留在父坐标轴；true 或缺省时跟随对象变换。旋转/翻转仍附 `shadow-transform-review`，没有声称完整宿主保真。九种 alignment 在缩放 100%、斜切 0 时不改变平移；非恒等缩放/斜切尚未绘制。字段含义见[微软 OuterShadow 文档](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.outershadow?view=openxml-3.0.1)。

非零 blur 使用 `sigma=blur/2` 的 SVG 高斯近似，过滤范围包含原轮廓、偏移后轮廓、线宽和 3σ 扩展；这是本地 review 选择，不是 Office 模糊核等价证明。每次均保留 `shadow-blur-approximation`。LibreOffice 的[阴影实现](https://raw.githubusercontent.com/LibreOffice/core/master/drawinglayer/source/primitive2d/shadowprimitive2d.cxx)和[模糊实现](https://raw.githubusercontent.com/LibreOffice/core/master/drawinglayer/source/primitive2d/GlowSoftEgdeShadowTools.cxx)采用独立栅格处理，不能仅凭半径数字宣称相同。

剩余组合：可见文字与形状的阴影组合、半透明绘制属性与阴影 alpha、主题色、复合效果、效果自身缩放/斜切，以及其他 owner。已知未覆盖状态不画猜测阴影并保留 partial 警告；非法几何/透明度、未知字段仍为 unavailable。空 `txBody` 不等于可见文字；其他几何/图片分支失败时，不把其占位框当投影主体。

自生成外部对照在 `tmp/shadow-probe-NF7q5C/run-T4bRSJ/`：同一 16 对象 PPTX 经 LibreOffice → PDF → MuPDF 72 dpi，与修改前、修改后本地 PNG 比较。`comparison-after-2.json` 的九项直接/零阴影范围检查通过，阈值 2px，实际边界差均为 0；三项文字/半透明不确定组合只验证明确省略阴影，不计为保真成功。其余四项是未验收观察：两项模糊边界差至多 2px，两项随形状变换与 LibreOffice 明显不同。该样例中 LibreOffice 对 rotateWithShape true/false 给出相同画面，故保留差异，不反向修改开关语义追求全绿。Agent 查看双方 PNG，不是人类或 PowerPoint 验收。

首次接入把已知未绘制的文字阴影升级为 unavailable，使 canonical fixture 发布失败；已恢复原有 partial 分类，继续明确省略不可靠阴影，没有修改发布器或放宽非法字段失败条件。外部旧 `after/` 失败产物保留，修正后 `after-2/` 与最终 `final/` 均发布 complete、可靠性 requires-review。最终 `comparison-final.json` 复验同样的九项直接检查、三项保守省略与四项未验收观察，并记录原始 PPTX/PDF、前后 PNG、SVG 和 receipt 摘要。原十二项源删除失败仍独立存在。

内部普通形状和 literal 自定义路径已复用共享 linePaint，覆盖虚线、点线、点划线、cap/join、宽度和 alpha；保留 none、零值及路径 stroke=false。合成和真实源编辑测试检查线段/空隙像素及重新投影。虚线节距仍是明确标注的 review 近似，不等于 Office 精确轮廓；主题、复合轮廓和路径端点箭头仍未完成。

形状、表格单元格及直接页面背景共享渐变函数，消费 2～16 个有序直接 RGB 色标及逐色标透明度。重复位置保留硬边，透明度 0 不回落为不透明。线性方向使用对象实际宽高和原生 `scaled=false` 的物理角度，非正方形的 45° 不按缩放后的坐标轴猜测；居中径向渐变见下节。独立 SVG 定义 ID 避免背景和对象相互覆盖。未知渐变类型、非法角度/色标/透明度及已检查的冲突填充仍明确失败。

真实 NativeAOT 报告 PfrKIJ 验证 0°、45°、90° 各自的作者和去快照源 no-op，以及源形状 90°→0°、表格仍 90° 的单独修改，共七例。实际红/透明/蓝平台像素、重复色标、原生角度、重新投影、原源/请求不变均通过；修改只改变 slide1.xml。45° PNG 已由 Agent 查看，不是人类或 Office 对照。合成检查另覆盖 180°/270°、端点、色标边界和非法输入；首次越界 alpha 反例暴露通用透明度工具会钳制数值，渐变入口已增加严格范围检查。该轮 presentation 4/4 和生成检查通过，整套仍有十一项源删除失败。

直接背景渐变新增九项真实回归，最终见 aVCT4i 的 `backgroundGradientCases`：0°/45°/90° 各有作者及去快照源 no-op；另从原始 90° 源独立请求改角度为 0、将透明平台 alpha 从 0 改为 0.5、删除整个背景。检查原生方向/色标、实际 RGBA 红蓝平台和透明像素、前景橙色像素、正式入口内容区像素及发布证据。透明背景没有垫成白色；三项编辑重新投影符合请求，只改变 `ppt/slides/slide1.xml`，其他 ZIP 成员逐字节保留，原源/请求不变。删除还检查 native/PPJ 背景消失及 XML 无 `p:bg`，其白色输出只代表本 fixture 的默认背景，不证明 master/layout 继承已实现。

合成背景正例先复现白底误画，再验证方向、alpha、背景先于前景及共享定义 ID；九种非法/不支持/冲突状态要求页面 failed、背景 unavailable、前景和另一有效页面仍保留。失败发生在页级背景绘制，不中断整份文稿。沿用 035472e9 默认包，没有修改 C# 或重建运行时；原七项渐变与其他独立回归仍执行，完整报告仍有十一项删除失败。

#### 径向渐变：实际路径范围和源编辑

原生 [`PptxGradientFillCodec`](../native/OfficeKit/src/OfficeKit.Codec/PptxGradientFillCodec.cs) 已支持 `path=circle` 且 `fillToRect` 四边均为 50% 的直接径向填充。本轮 JS 消费同一场景字段，没有新增协议、依赖或原生解析器。中心取实际几何范围的中心，半径为 `hypot(width, height) / 2`；形状只占部分外框时，不能继续使用整个外框。此映射依据[微软对径向渐变及实际路径范围的说明与后续更正](https://learn.microsoft.com/en-us/answers/questions/2248059/non-preset-a-tilerect-behaves-strange-in-case-of-g)。SVG 使用物理坐标下的 `userSpaceOnUse`，避免非正方形对象把圆拉成椭圆；坐标语义见 [SVG 渐变规范](https://www.w3.org/TR/SVG2/pservers.html#RadialGradients)。

[`nativePathBounds()`](../src/ppj/preview-scene-svg.mjs) 复用已有 literal 路径求值，计算线段端点、二次/三次曲线内部极值及圆弧经过的轴向极值，并合并多条路径。控制点外框、孤立 moveTo、线宽和阴影不当作填充范围。未解析引用、空或退化范围明确报 `preview.scene.paint.gradient-geometry` unavailable；可读文字保留。合成回归覆盖横/竖/正方形、部分外框路径、曲线极值、正反/整周/非等比圆弧、无 viewport 的 EMU 坐标、多路径及失败时原场景不变。

两色及首尾同色三色的特殊亮度曲线已接入共享函数，见下节。其 SVG 采样精度仍明确标记，不计为 Office 宿主验收；非居中焦点、其他 path 类型、参数化 tileRect、主题色标及文字/图表等其余消费者仍未完成。

TZi2W5 的 `radialGradientCases` 有 14 项：同一文稿中的形状、单元格和背景各有径向填充，作者及去快照源 no-op 两例，加上三类对象各自独立改色、改 alpha、改为线性、删除填充，共十二项源编辑。检查实际原生状态、红/透明/蓝平台、左右及垂直等距采样、独立轮廓/文字、前景和正式发布。十二项编辑均重新投影确认三个消费者的填充与请求一致；删除保留缺失，线性保留显式 angle=0；完整 ZIP 清单不变且仅 `ppt/slides/slide1.xml` 修改。

另有 8 项 `radialGeometryCases`：只占外框左半边的自定义矩形、二次曲线、三次曲线分别有作者与源 no-op；前者另有从原始源独立修改颜色和移动/加宽外框。实际范围分别为 100×100、200×50、200×75，移动/加宽后为 200×100；原生路径、SVG 中心/半径和 RGBA 均验证，不能用外框或控制点范围冒充曲线范围。两项源修改比较完整语义页、权限操作/字段范围及有序叶子类型和值；候选的新 source hash/revision 单独核验，重新签发的 handle/叶子 ID 不要求与旧源相同。再次 no-op 保持候选字节，其他 ZIP 成员逐字节保留，只有 slide1.xml 改变。

首轮 eZ2VRC 的 `background-linear/table` 失败属于测试叠加预期错误：x=745 位于线性背景 75% 以后的蓝色平台，表格透明区域应透出蓝色。保留该取点、纠正预期并新增左侧取点后通过。自定义几何测试也修正了源/候选文件名冲突、误用 wire 命令名代替 PPJ `quadraticTo/cubicTo`，以及把源修订身份误当成不变内容的断言。上述失败报告保留；没有删除源编辑检查或弱化原十二项删除失败。最终 22 项正式结果均为 requires-review，`radialGradientFailures=[]`；完整集成仍因原十二项失败退出 1，没有修改 C#、重建包或完成 Office/人类验收。

#### 渐变亮度插值：颜色与透明度分开计算

Office 对端点位于 0%/100% 的两色渐变，以及首尾 RGB 相同、位于 0%/100%、中间色标严格在内部的三色渐变，采用偏向较亮通道值的 1.875 次幂曲线；alpha 仍线性变化。共享 [`gradient()`](../src/ppj/preview-scene-svg.mjs) 已实现这一映射，依据[微软的两色与三色插值说明](https://learn.microsoft.com/en-us/answers/questions/2248059/non-preset-a-tilerect-behaves-strange-in-case-of-g)。普通色标继续沿用 [SVG 的逐通道线性插值](https://www.w3.org/TR/SVG2/pservers.html#Gradients)：包括不在两端的两色色标、重复端点、非首尾同色三色及四个以上色标。

每个特殊插值区间固定细分 32 段，只增加 SVG 色标：两色共 33 个、三色共 65 个，原始端点/中间色标与 alpha 保留。权重在模块内预计算，没有新增依赖；PPTX 和原生场景仍只含原来的两个或三个色标。数学检查得到曲线弦线误差小于 0.088 个 8 位通道单位，加上 RGB 取整后小于 0.589；密集采样验证实际生成的 SVG 色标插值满足该界限，alpha 不施加亮度曲线。`preview.scene.paint.gradient-interpolation` 继续为 partial，但说明已改为“32 段有限精度映射”，而非“尚未实现”。此界限不包含宿主色彩管理和栅格量化，不代表整张 PNG 与 Office 的误差界限。

WQsUZk 的 36 项 `brightnessCases` 全部通过：三类消费者 × 线性/径向两类，各有作者和去快照源 no-op，再分别从原始源独立请求反转颜色、全部 alpha=0、改为首尾同色三色、将末端色标移到 90%。检查原生色标数量/值、SVG 色标数量及是否采用特殊曲线、独立公式预测的实际合成像素、前景与正式入口；像素比较要求 alpha 误差至多 1、预乘 RGB／白底合成通道误差至多 2 个 8 位单位。24 项源编辑均重新投影并核对非目标 ZIP 内容，仅 slide1.xml 改变，候选/请求/重新投影文件均保存摘要。

全透明无文字形状采用既有投影规范：[`ProjectShape()` / `TryGetCompoundShapeOpacity()`](../native/OfficeKit/src/OfficeKit.Codec/PpjPresentationProjector.cs) 将相同的绘制透明度合并为 `compositing.opacity=0`，而不在每个色标重复列零。测试精确检查这一合并、其余填充字段不变，并从重新投影结果再次编译：候选字节相同，原生色标仍显式 alpha=0，像素仍透明。表格和背景保持逐色标零值。首轮 h2NK3z 修正了 fixture 误用 `position` 而非 PPJ `offset`；O5Yxl7 的两项“丢失零值”断言经上述代码和真实二次编译证明是忽略整体透明度的测试错误，不是新原生故障，也没有用加文字绕开合并场景。

新增 36 项正式结果均为 requires-review；整套仍因原四项变换、七项间距和一项共享背景删除失败退出 1，`brightnessFailures=[]`。实际三色径向 PNG 已由 Agent 查看；本轮没有原生改写或重建，不宣称人类校准、Office 色彩保真或完整 G-04 完成。

#### 直接图片背景：透明度、裁切和共享引用

直接图片背景现复用普通图片的资产、正负裁切及 alpha 绘制函数，不重新解释 PPJ 的 fit/focus。原生 `imagePaint.mode` 为 UNSPECIFIED/STRETCH 时按实际 canvas 与裁切绘制；源图片透明像素、显式 opacity=0、半透明及负裁切留白都保留透明度，不铺成白底。无额外 alpha 标志的旧 `imageAssetId` 有合成绘制回归；旧 `imageAlphaModulationFixed` 没有明确数值时保持 unavailable，不猜透明度。

oUjI1t 的 `backgroundImageCases` 有十二项：拉伸、opacity=0、正裁切加 0.5 透明度、负裁切加 0.5 透明度各有作者及去快照源 no-op，共八项；另从同一原始正裁切源分别修改裁切、设 opacity=0、删除 opacity、删除 crop，共四项。固定图片包含红/透明/绿三个区域，背景和一张前景图片共用同一资产。检查原生 crop/alpha、实际 RGBA、前景形状与图片全透明度独立性、正式入口内容区像素；四项修改重新投影正确，删除保留字段缺失而非显式默认值，完整 ZIP 清单不变且仅 `ppt/slides/slide1.xml` 改变，共享图片和关系文件逐字节保留。实际发布 PNG 已由 Agent 查看，不是人类校准。

平铺背景的作者与源 no-op 两例另列 `backgroundImageRejections`：不画成单张拉伸图，保留前景，输出不可用背景和红色警示；正式发布按既有契约抛出 `preview.output.incomplete`，留下 SVG/PNG 与最终 `render.json`。异常 receipt 与落盘一致，资产/scene/候选身份、产物摘要、正文及警示像素通过。它们是防误画和失败证据测试，不是平铺渲染成功。首次 Fa1mls 把该有意拒绝误当作正常发布，已按现有发布契约改为精确异常断言，没有放宽生产失败条件。

**新增未关闭失败：共享图片背景删除。** 同一原始源中仅删除整个 `page.background`，原生编译返回 `presentation_element_binding_mismatch: Presentation slide 1 element 2 changed its source capability contract.`，还未返回候选。oUjI1t 保留 `background-image-crop-source.pptx`、`background-image-source-delete-background.ppj` 及报告中的源/请求摘要，`backgroundImageFailures` 中此例仍使整套失败；没有通过换掉共享图片或去掉断言规避问题。原始源/投影/请求和资产保持不变。

源码显示一个需要原生回归确认的原因：[`PptxCodec`](../native/OfficeKit/src/OfficeKit.Codec/PptxCodec.cs) 先调用 `PptxBackgroundCodec.Apply`，再计算用于 `AssertElementBinding` 的源删除权限；[`PptxElementDeletionCodec`](../native/OfficeKit/src/OfficeKit.Codec/PptxElementDeletionCodec.cs) 按图片关系是否被对象外引用判断权限。删除背景改变共享引用后，重新计算的权限可能与原始投影不同。这是基于顺序和错误的定位，尚未用专门的原生诊断证明具体哪个权限字段发生变化。修复应在原始修订上验证权限，并保证操作后的关系清理安全，不能直接跳过绑定检查；涉及原生写回语义，待独立 change 范围确认，本轮没有修改 C#。

#### 形状图片填充：有界几何、独立轮廓与源编辑

[`shape()` 与 `shapeGeometry()`](../src/ppj/preview-scene-svg.mjs) 现消费完整原生 `imageFill`：使用已有图片资产/正负裁切/alpha 函数，将图片裁到同一形状几何，再独立绘制轮廓和文字。矩形及其别名、圆角矩形、椭圆、菱形及其别名、可求值 literal 自定义路径已有对应分支。`fillMode=NONE` 的自定义路径只参与轮廓，不加入图片遮罩；stroke=false 不绘制对应轮廓。图片完全透明时，轮廓和文字仍保留自身不透明度。

仅有旧 `imageFillAssetId` 的源对象仍报 `preview.scene.paint.shape-image` unavailable，因为资产身份不能证明完整 alpha/旋转等填充状态。未映射 preset/adjustment、未解析路径、平铺或未知 mode、非法 crop/alpha，以及图片与 solid/gradient/useBackgroundFill=true 冲突时也保守失败；可读取的文字和轮廓仍保留，不能为了显示图片将源对象改写为普通矩形。合成正例先复现原先无图片填充，再覆盖七种 preset/别名、带 stroke-only 路径的自定义几何、零 alpha、十二种反例与原场景字节不变。

首次 1A1hYE 新增十六项 `shapeImageCases`，7pSsUc 补上事实映射断言，本轮 WQsUZk 继续验证，不重复计数。rect、roundRect、ellipse、diamond、literal 自定义菱形各有作者和去快照源 no-op，共十项；另从原始 rect 源独立请求改为负裁切、设 opacity=0、删除 opacity、删除 crop、水平翻转、删除整个 fill，共六项。固定红/透明/绿图片铺在蓝色页面上，检查原生资产/crop/alpha、半透明混合像素、几何外部确实为空、独立洋红轮廓和黑色文字；正式入口内容区像素、候选身份和返回/落盘一致通过。六项源编辑均重新投影符合请求；成功候选另存 `shape-image-source-*.pptx`，源与请求摘要保留。

前五项编辑仅改变 `ppt/slides/slide1.xml`。删除整个 fill 则只移除候选中已无引用的 `ppt/media/image.png`、删除 slide1 的这一条图片关系，并修改目标 slide；测试核对被清理图片的原始字节恰为该测试资产，关系 XML 除这一条删除外完全相同，其他 ZIP 成员逐字节保留。原始 PPTX 和输入资产没有被删除，仍可恢复填充。这与前述“背景和前景共享图片”的删除失败是两个不同所有权案例，不能互相替代。

首轮 9zvTJi 的两个断言问题已定位：圆角外部取样恰落在轮廓抗锯齿边缘，蓝底多出 1 个红通道单位；取样改到经实测完全在轮廓外、又严格在矩形填充内的位置，仍用精确 RGBA 比较。整个 fill 删除时，初版要求 ZIP 清单完全相同，但实际正确清理了最后引用的图片；已改为上述精确所有权断言，没有允许任意部件变化，也没有绕过共享背景删除失败。

该组 `shapeImageFailures` 为空，整套仍因原十二项删除失败退出 1。初次 1A1hYE 保留了旧几何遗漏红色警示；随后 G-11 根据完整几何及变换记录解除这十六例的误报，正式结果变为 requires-review，其他限制仍在。已查看实际翻转候选 PNG，仅为 Agent 检查；复杂形状内部文字布局、完整 fit/focus、效果、任意 preset 和 Office/人类验收仍缺。

theme/master/layout、命名样式及继承状态仍未统一解析；非居中/其他路径渐变、图表/文字渐变、继承背景、剩余形状几何的图片填充、外阴影的剩余组合与宿主保真、内阴影、glow、reflection、soft edge，以及效果缩放/斜切仍缺实际消费或验收。特殊亮度曲线已有有限精度绘制，宿主色彩管理仍未验收；直接图片背景的完整 cover/contain/focus 与源背景删除也未验收，不能由已验证的 signed crop 推导完成。

风险不止外观：低对比度、透明边缘和阴影范围错误会让 reviewer 错判内容是否存在或是否越界。应消费确定的有效样式；未解析主题值不得静默替换后宣称正确。完成证据应包括直接值覆盖继承值、删除后恢复继承、显式 0/false、半透明叠加及效果边界。

### G-05：图片

**不要再把裁切、有限遮罩和边框整体记为“未实现”。** 它们已有第 3 节的限定实现及真实证据，现由正式入口共用；直接图片背景还新增了共享裁切/alpha 的十二项真实正例、两项平铺拒绝及一项背景删除失败，详见 G-04。默认包配套要求见第 2 节。

当前具体残留：

- tile 尚未实现实际重复绘制；内部 painter 已修正未裁切 tile 被画成单张 stretch 的错误，现在有/无裁切均明确失败并保留对象占位。真实 jHmDS0 报告验证作者/源 no-op 四例，源与 no-op 候选不变；三种裁切存在性另有合成回归。这是防误画修复，不是 tile 渲染通过，也未接入正式 CLI。
- 全部 fit/focus 输入到编译后状态的等价性尚未验证；不能只看最终 frame 就认定 cover/contain 等全部正确。
- 遮罩仅覆盖少量 preset 与可解析 literal 路径；其他 adjustments、复杂填充模式和公式引用仍缺。
- 边框需要已解析 RGB；未解析主题色会产生 unavailable 诊断，该边框不绘制，图片内容仍可保留。虚线长度和部分线条轮廓仍是有说明的近似。
- 阴影、反射、软边及复杂透明合成未完成。裁切几何正确不等于识别出了图片主体。

后续使用同一资产快照，不能重新读路径造成编译/显示不一致。最小案例要验证非居中裁切、负边留白、透明边缘、遮罩与旋转组合、阴影扩边、空字节和解码失败；主体范围不确定时不擅自裁切。

tile 的后续实现需要先补足尺寸语义：当前 PresentationImage 仅有参数为空的 tiled 布尔状态，scene 不提供已解析的平铺单元物理尺寸；源图片解析接受 useLocalDpi 的 0/1，但没有对应传输字段。不能把所有源图片一律按固定 DPI 或按图片框大小平铺。需要核对原生导入、有效 DPI/固有尺寸及缺省规则，再验证实际重复数、偏移、裁切、透明度和遮罩组合；不把当前失败占位视为此项的最终交付。

### G-06：表格和连接线

共享表格 painter 使用真实尺寸和合并范围，正式旧路线的均分网格已删除。剩余包括合并外围边框、冲突边优先级、继承/banding、单元格边距和排版。含可见文字的覆盖格也需要明确处理，不能通过合并静默抹掉内容。

内部连接线已使用原生端点，并消费 elbow 的 literal bendAdjustment，显式 0、负值和超过 100000 的值有合成覆盖。剩余是 curved、完整导入变换来源、任意 connection-site 求解/编辑、嵌套及组件锚点广泛回归。自动避障若不在当前支持范围，应明确限制；不能以 frame 中线或“最近对象”猜测关系。

对象锚点的两例通过只证明那两个场景。新 connection-site 接口或在途变更需要独立核验，不能提前记为 SVG 已覆盖。完成证据应包含目标移动、重定向/解绑、反向端点、折点变化、箭头方向，以及源文件和非目标部件保留。

### G-07：source-bound、opaque 和静态检查边界

原生媒体静态封面已有内部绘制（2026-09-10）：`PresentationMedia.posterAssetId` 绑定已验证图片字节，按现有 `PptxMediaCodec` 的 Stretch/FillRectangle 外框显示，保留透明度和外层变换；对象内增加可见 STATIC AUDIO/VIDEO POSTER 提示，诊断明确不验证播放、时序或载荷外观。未知媒体类型、空封面或非图片封面明确失败；未声明的资产引用由 scene 完整性检查拒绝，绝不把媒体载荷当图片使用。合成覆盖 audio/video、翻转、缺封面、缺资产和原 scene 不变。

真实报告 `tmp/officekit-native-scene-paint-rnrCq1/integration.json` 的 mediaPosterCases 有两个封面变体：作者输出的橙/蓝实色像素、透明半边背景和可见警示条均通过；每个作者文件去快照后重新投影，源媒体仍为 opaque，no-op 候选字节等于原源，未借作者快照伪造导入封面。源/输入/资产字节不变，已查看作者 PNG（Agent 检查，非人类校准）。媒体载荷仅为测试容器头，不是可播放视频。初次 h5iVdP 因测试资产缺必填 accessibility 被 schema 拒绝，补齐元数据后重跑；未放宽验证。沿用第 3.5 节原生包，Presentation 4/4、两项生成检查和 gate-policy 通过，整套仍有十一项删除失败。当前 opaque 描述符没有独立预览资产，OLE/导入媒体封面仍缺原生读取与传输，不在此宣称完成；未重建或切换正式入口。

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

正式路线已消费编译后的原生通道和生成图形，不再自行解释 categories/values 或 dataset。完整类型、复杂轴/通道绘制及对应旧事实规则的逐项解除仍未完成。

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

形状几何误报已按实际绘制记录修正（2026-09-10）。[painter](../src/ppj/preview-scene-svg.mjs) 仅为完整生成的预设/literal 路径记录 `shapeGeometryScenePaths`，不把图片 clip 定义、未知几何、未映射 adjustment 或部分成功路径当成完整形状。[输入检查](../src/ppj/preview-input-assessment.mjs) 还要求同一 owner 的全部绑定都是原生 shape，且节点与全部祖先变换均成功；只解除 `preview.fact.shape-geometry-omitted`。文字失败导致整个节点 SVG 被丢弃时，即使几何已经构造也保留错误。

合成回归先复现旧误报，后覆盖七种预设/别名与 literal 自定义路径，以及隐藏、未知几何、混合节点类型、未绘制兄弟、坏路径、未映射 adjustment、失败父节点和绘制后文字失败等十一项反例；另验证缺几何/自身/祖先记录、缺 registry 和多绑定 owner。真实形状图片的十六例在原有独立 RGBA、轮廓/文字、源编辑与发布检查外，逐例证明旧错误存在、映射后解除、删除记录恢复、缺 registry 拒绝，其他输入诊断完全相同、绘制诊断保留。其状态为 requires-review，不是 passed。

正式门禁 `ppj-preview-render-assessment.mjs` 原来用已能绘制的矩形制造几何失败；现在使用确实未映射的 star5，保留两页相同 ID、嵌套路径、opaque 不泄露与红色警示断言，另加矩形正例。字段级 partial、其他事实规则、未知字段与 opaque 边界不变。任务 4.1 尚有其他规则待逐项核验，没有据此勾选完成。

dataset 折线误报已按实际绘制证据修正（2026-09-10）：registry 的 `datasetLine` 只处理普通 line 的 `data.dataset`。同一输入 owner 的全部绑定必须是具有分类/系列的原生 LINE，实际完成折线 SVG 构造，且节点及全部祖先都完成外层变换；JS 不重新求值 dataset/encoding。隐藏、平滑未实现、通道长度错误、缺捕获、混合 owner 或失败祖先均不能解除原错误，其他图表也不能借用此映射。

最终 L4nZbV 的 `datasetProfileCases` 有两个正例和一个保留错误的 heatmap 反例。两个正例分别将固定数据首值设为 1、2，使用独立的显式数据侧验证完整原生载荷与整页像素相等；孤立点实际从 y=313 移到 y=271（x=160），旧位置无蓝色墨迹，真实零位置仍有墨迹、缺失类别的零位置无伪点。原输入不变，scene 开关候选字节相同，正式入口返回/落盘一致。删除捕获恢复旧错误，缺 registry 映射拒绝，其他输入限制逐项保留；正式折线结果为 requires-review，不是 passed。heatmap 仍为 failed，图表整体等级不变。合成正例先复现旧误报，再通过六种反例及祖先捕获检查；错误的“大角度必失败”测试假设已改成实际零子尺寸失败，没有修改旋转语义。任务 4.1 与整体目标继续开放。

折线事实映射新增有界证据（2026-09-10）：registry 的 isolatedLinePoints 要求同一输入 owner 的全部绑定都是完成外层变换的原生 LINE，具有实际可见孤立点捕获，且原始 literal categories、series 顺序/名称、values、null 与原生缺失索引逐项相等；不重算 dataset。符合时解除旧系列类型未继承错误；孤立点错误还须命中该 series/point 的可见捕获。只在完整 SVG 构造后记录，隐藏、透明、坐标范围外、平滑线未实现、数值不符或混入其他 owner 类型均保守保留。其他输入/绘制诊断继续合并，不改变默认 profile。

真实 `tmp/officekit-native-scene-paint-nPwMZz/integration.json` 的 isolatedLineProfileCases 有四项：作者、源 no-op、无关表格移动、源图值编辑，每项实际解除两条 series-type-not-inherited；移除捕获恢复原错误，移除 registry 映射拒绝，候选字节不变。原有 marker/缺口/零值实际像素和源部件保留断言继续通过。真实输入没有显式 series.chartType，因此本报告没有触发并解除 missing-observation-misrepresented；这条逐点解除目前是合成反例边界证据，不能与四项真实结果混算。首次 jiMcTB 为触发旧孤立点规则加显式系列 chartType，但被作者编译器拒绝；已恢复原输入契约，没有放宽编译器。smg21U 随后发现旧连接线测试把新折线映射也当成“其他规则不变”，现将两类已验证规则分别断言，其余诊断仍原样保留。沿用第 3.5 节包，整套仍有十一项删除失败，退出 1；未重建、提高完整等级或切换正式入口，4.1/full goal 继续开放。

显式连接线端点的旧事实误报已有逐字段映射：registry 的 connectorEndpoints 指向 `all-owner-direct-endpoints-painted`。painter 仅在实际生成直线/折线路径后记录 scenePath；assessment 还要求该节点完成外层变换，同一 owner 的全部绑定均为已画出的原生连接线。每个 from/to 分别核对纯 x/y 输入与安全 EMU 转换后的原生端点，带目标绑定或对象锚点的端点不据此豁免。接收到 scene、画出占位或只遍历节点都不算端点绘制证据。

合成正例覆盖 straight/elbow；坐标不匹配、hidden、curved 占位、混入非连接线 owner、对象锚点继续保留相应事实错误，另一端有独立证据时只解除另一端。真实 GjnuAZ 的 connectorProfile 六例覆盖直线作者/源 no-op/无关表格移动候选，以及折线作者/源 no-op/折点编辑候选；原有路径、箭头和折线 RGB 像素断言继续通过。每例验证两个旧端点错误解除、删除捕获记录后恢复、缺 registry 映射拒绝，以及原输入的无关诊断和合并后的绘制诊断保留。

本次未改原生包、绘制路线几何或整体等级，只按已有实际绘制证据修正 profile 误报。Presentation 4/4、两项生成摘要检查通过；完整 GjnuAZ 仍因四项 frame 删除失败退出 1。对象锚点、任意 connection-site、导入折线变换来源及正式可靠性接入仍未关闭，任务 4.1 保持开放。

支持诊断清单为 9/9，输出安全与证据清单为 8/8。已具备的限定能力包括支持状态/事实失败区分、图内警示、资产快照复用、非覆盖发布、失败产物和输入/输出摘要记录。

这些历史清单完成结论归属当时的旧路线。scene 发布随后完成任务 4.2：场景身份、作者/源候选、24 例操作失败及两例事实失败警示见第 3.4 节。正式接入及默认包新增证据见第 2、3.6 节；它不补足全部 profile 规则。

当前声明为 16 类元素中 9 类 partial、7 类 opaque；16 种 chartType 均 partial，没有整个类型提升为 supported。这是保守声明分布，不是完成率。

正式节点检查树已强制合并原始输入，仍需完成其他事实映射、更多归属边界及后续运行时配套复验。场景身份、源 no-op/文字候选发布、操作失败与事实失败警示已有证据；新类型或新错误分支仍须扩充回归。旧事实规则只能在有对应修复断言后按 profile 调整，不能为了让新路线变绿全部删除。

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

1. **维护验证快照**：以第 3.6 节的 035472e9 默认包作为当前回归基线；第 3.3、3.5 节及 G-14 保留旧构建和两项 AOT 修复证据。后续修改记录源码、dirty diff、二进制 manifest/hash 和测试脚本身份；新字段继续补视觉断言，原生语义改变时再更新精确包验证。当前基线的原十一项删除失败及新增共享图片背景删除失败仍须保留，不能称作全绿基线。
2. **补基础显示**：文字格式/布局、剩余几何、主题及图片效果；每次选一个字段或一组直接相关行为。
3. **补事实风险高的图表与关系**：源 X 编辑、X 缺失表达、散点连接语义、气泡 size、双轴、层级/流向/OHLC、缺失显示策略；保留已有数值 X/Y marker 回归，优先复用编译生成结果。
4. **持续核对默认包配套**：正式函数已接场景、profile 检查及发布保护，当前 linux-x64 默认包的无 preload 验证已完成。后续每次原生更新仍需构建和复验；未解决项继续明确失败或标注限制，接入不等于全面覆盖。
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
