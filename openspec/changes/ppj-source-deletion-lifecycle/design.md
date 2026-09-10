## Context

动机与授权见 [proposal.md](proposal.md)。本设计基于 2026-09-10 的只读源码检查和冻结回归，尚未实施原生修复。

### 已观察到的失败和版本边界

`tmp/officekit-native-scene-paint-ay3U2S/integration.json` 使用冻结 JS 副本 `locale-inset-snapshot-aKeyOX` 和本机默认 035472e9 原生包：506 个正式入口完成限定检查，整套 `failed`，本变更对应 12 项失败；另有 3 项语言/大小写失败不在授权内。

| 组 | 固定案例 | 实际失败 | 当前源码线索 |
| --- | --- | --- | --- |
| frame | rotate、flip-h、flip-v、combined | 候选重新投影仍有显式 `rotation: 0` / `flipH: false` / `flipV: false` | `PpjPresentationCompiler.TryCollectFrameLeafMutations` 比较解析后的值，将变更写为数字或 `1/0`；`ApplyFrame` 已有按 JSON 存在性清除字段的路径，但快速叶写回没有同样的删除语义 |
| 间距 | 三个倍数槽位的删除 mask 1～7 | 全部 `ppj.source.unsupportedMutation`，无候选 | 当前源码已有 `MaskTextValues`、逐字段授权及 `ApplyTextParagraphStyleMutation` 的间距处理；三个近期 lifecycle 测试覆盖单槽位删除/恢复，尚未在本轮用新包验证这些组合 |
| 背景 | `crop/delete-background` | `presentation_element_binding_mismatch`，slide 1 element 2 能力契约变化 | `PptxCodec` 先 `PptxBackgroundCodec.Apply`，后计算 `sourceDeletionPlans` 并调用 `AssertElementBinding`；图片删除能力分析会计算 slide 内关系引用次数 |

源码检查时 HEAD 为 `4630073a32d5cbd359df3c75ff9f8feaf3ba4985`，并有未提交/并行更新。这个 HEAD 不是完整源码快照。已安装 PPJ 可执行文件摘要为 `eeca85542b5758c567dff712cbf1d37874f9518b7471e42216c597e3fec4bd71`，Office profile 为 `a2217c0cc39fd70c5c0e739c75990d7e5e6602bbbed81fabc550ac485b1e4613`。实施时需冻结实际文件清单和摘要，不能把新源码测试、旧二进制结果混为一谈。

### 可复用的现有实现与测试

- `PpjParagraphSpaceBeforeLifecycleTests.cs` 的共享 helper 已检查两类 owner、点值/倍数切换、删除恢复、非目标 XML 和未知节点保留、精确能力字段及非法输入；line-spacing/space-after 测试复用该 helper。
- `PptxCodecTests.SourceBoundImageBackgroundCropAndOpacityEditOnlySlideAndReprojects` 覆盖背景裁切/透明度修改，但不覆盖背景与前景共享关系时删除整个背景。
- `test/ppj-preview-scene-native.mjs` 已保留四项 frame 删除、七份独立 spacing 请求和共享背景删除反例，不能削弱其断言使报告变绿。
- `PptxBackgroundCodec.Apply` 已在删除背景后调用 `RemoveIfUnreferenced`。本次共享案例应保留仍被前景引用的关系，不新增全包垃圾回收逻辑。

## Goals / Non-Goals

**Goals:**

- 将源绑定验证固定在原始状态，候选应用仅使用已经验证的权限与目标。
- 对属性存在性做最小局部写回，保留未知 XML、未改数值拼写和非目标 ZIP 部件。
- 用同一套可再生成案例连接托管测试、真实 NativeAOT 候选及只读预览证据。

**Non-Goals:**

- 不新增通用编辑事务框架、第二份 PPJ 解释器、协议字段或渲染引擎。
- 不开放原先不支持的 connector frame、未知间距拓扑或图片删除权限。
- 不将语言/大小写写回、完整字体排版、宿主像素一致性或默认包发布纳入本变更。
- 不重做已完成的单字段生命周期实现；新源码已修复的案例只补组合验收与版本证据。

## Decisions

### 1. 区分存在性变更与值变更，复用既有局部写回

`frame` 的 before/after 原始 JSON 决定字段是明确赋值还是删除，不能只比较解析后回落为零/false 的值。优先让涉及删除或存在性变化的请求退出不支持删除的 leaf fast path，经过现有 `ApplyFrame` 和原生 frame codec 的清除逻辑；普通坐标和显式值修改继续走快速路径。若既有局部写回不能保证目标 owner 的 XML 保留，则只为对应属性补有界删除处理，不能通过整体重建 slide 解决。

补充显式零/false → 缺省的 presence-only 反例，以及删除一个字段、显式重置另一个字段的混合请求。形状、图片、组是当前快速路径的主要影响面；chart/table 的既有 frame 行为做非回归检查，connector 的拒绝保持。

未采用“所有删除都写默认值”：XML 存在性会影响继承和后续编辑，即使当前 PNG 相同也不符合删除要求。未采用扩展 wire 删除操作：现有语义模型已有清除方法，先使用已有契约。

### 2. 间距先复验现有实现，再补组合缺口

将七种删除 mask 加入既有段落生命周期测试，覆盖普通 text/shape、三个倍数槽位、点值及混合单位、只含间距的 style 删除。每个原始删除请求重新投影同一源，避免前一次修改改变后一次测试前提。恢复案例则明确以删除后的候选为新源重新投影。

当前 `ApplyTextParagraphStyleMutation` 已能按槽位保留/清除；仅当新源码回归仍失败时，修正变化检测、精确能力检查或相应槽位写回。保留 before/after 两个单位字段的授权规则，保留非法输入和 source-owned 未建模间距的拒绝。

未采用放宽整体 rich-text 比较：其保护未知格式、run 拓扑与未授权修改。也不复制另一套段落写回入口，以免后续更新漂移。

### 3. 源绑定分析先于背景候选修改

当前调用顺序足以解释共享背景删除反例：背景和前景共用关系时，原始前景图片的删除能力被“有外部引用”限制；背景先删后，重新分析可能得到不同的权限及原因，随后严格比较原始绑定失败。这是源码支持的触发机制，具体不匹配字段需由新增托管反例断言确认。

在当前 slide 写回流程中，于背景修改之前捕获用于验证的原始元素、删除计划及其他受共享引用影响的能力结果。将绑定校验放在该原始基线上，再应用已验证的背景变更；不删除 `AssertElementBinding` 的任何检查，也不把候选新获得的独占资源能力当作本请求的授权。

范围保持在本次 slide 源写回，不推广为全项目事务框架。若同一请求还提出元素删除，必须满足原始源签发的删除能力；背景删除不能让原先被拒绝的图片删除突然获准。合法的新能力仅通过候选重新导入获取。

媒体清理继续使用现有 `RemoveIfUnreferenced`。固定共享背景案例要求只有 `ppt/slides/slide1.xml` 改变，前景图片子树、媒体、关系和其他 ZIP 条目保持。背景不共享的已有删除/替换行为做回归，不扩大其资源回收政策。

### 4. 分层验收，但不互相抵消失败

每个目标案例检查三层：

1. 结构：原生属性/槽位或直接背景真正移除，候选通过 Open XML 检查；目标部件内非目标 XML 与其他部件保持。
2. 编辑保真：候选新导入/新投影恢复正确存在性、值、单位及保留对象；源和请求摘要不变；no-op 原文件逐字节相同。
3. 有限视觉：真实候选 scene 的身份与 PPTX 摘要一致，scene 开/关候选字节一致；与独立作者参考比较限定内容区域，保留 `requires-review` 和已有不支持诊断。

12 项作为可单独执行/汇总的 native deletion 回归，不依赖其他 locale 案例先成功。全套仍执行并保留独立失败；局部 12/12 与整套状态分别报告。一次性产物写入新的 QA 目录，测试源码与 fixture 生成逻辑留在正式测试目录。

## Risks / Trade-offs

- [旧包结果与近期源码不一致] → 冻结源码及 SDK，先跑已有托管生命周期测试，再构建到独立输出目录；报告列出两种基线，不能将旧包失败冒充当前源码缺陷。
- [改变背景应用顺序影响能力或资产识别] → 针对共享/不共享、no-op、删除/替换及伪造前景能力补反例；原始源分析与候选资源清理分开。
- [退出快速路径导致邻近 XML 被规范化] → 对目标元素移除允许变化字段后比较剩余 XML，并比较非目标部件原始字节；失败时补局部写回，不放宽预期。
- [同一文件有并行生命周期改动] → 复用已有 helper，按当前实际 diff 合并，不覆盖其他变更；记录构建时源码摘要。
- [图片背景删除后的页面继承背景] → 只保证直接背景缺省；白底像素预期只用于没有继承背景的固定 fixture，不把任意删除解释为设置白底。
- [局部成功被误报为整套完成] → 原生 12 项、其他 3 项 locale 问题、整体回归和宿主验收分栏记录。

## Migration Plan

不需要数据迁移或 wire 版本变更。实施后使用仓库 `build:office-kit` 命令输出到新目录，生成并核验两个 NativeAOT profile、manifest、notices 和 SBOM；使用明确指定的新包运行窄门禁和真实集成。默认安装包在本变更中保持原样，包替换/发布另按仓库发布流程执行。

原有源输入不变，可随时继续使用原包；回退时只停用新包和本变更代码，不回滚用户的 PPTX。新旧包、输入、候选与报告都按摘要识别，不覆盖历史失败证据。
