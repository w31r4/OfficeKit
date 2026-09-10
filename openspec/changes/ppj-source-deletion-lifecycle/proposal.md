## Why

只读预览回归暴露了 12 项原生写回失败：4 项 frame 旋转/翻转删除留下显式 `0/false`，7 项段落间距删除组合被拒绝，1 项共享图片背景删除触发源能力绑定不匹配。用户已授权将它们单独纳入原生修复，避免以画面相同替代编辑保真，或突破 `ppj-preview-compiler-scene` 的只读边界。

## What Changes

- 让已获 `setFrame` 授权的 `rotation`、`flipH`、`flipV` 删除真正移除原生属性；保持显式零/false、未指定和删除三者的区别。
- 验证并补齐行距、段前、段后三种间距的 7 种非空删除组合。当前源码已有三个单字段生命周期变更，先用新构建复验并复用它们，仅修复仍存在的组合缺口。
- 修复共享图片背景删除的源校验时序：对未修改的源计算并验证绑定，不因本次背景删除改变前景图片的源能力契约；保留图片、共享媒体和关系。
- 每项均以原 PPTX 字节和独立重新投影的请求验证，检查目标 XML 真正删除、非目标内容保留、候选重新导入及对应预览；保留无权限、伪造绑定和 opaque 反例。
- 记录源码与 NativeAOT 包身份，建立这 12 项的独立验收结果，同时如实保留全套回归中的其他失败。

## Capabilities

### New Capabilities

- `ppj-source-deletion-lifecycle`: 在既有 source-bound 授权下完成 frame 属性、组合段落间距和共享图片背景的删除，约束原始绑定校验、局部写回与重新导入证据。

### Modified Capabilities

无。`openspec/specs/` 当前只有 `.gitkeep`；相关既有要求位于各变更的 delta specs。本变更补充删除组合和校验时序要求，保留 `ppj-frame-transforms`、三个段落间距 lifecycle 变更及只读 preview 变更的边界。

## Impact

- 原生实现：`PpjPresentationCompiler.cs` 的 frame 快速写回和段落样式变更；`PptxCodec.cs` 的背景应用、源能力分析与绑定校验；必要时使用现有 frame/background/paragraph codec 的局部修改入口。
- 测试：C# 的 frame、段落间距、背景源保留回归；`test/ppj-preview-scene-native.mjs` 中原 12 项真实 NativeAOT 失败。
- 文档：更新 PPJ 能力/Skill 的相关条目及渲染差距报告，只按验证结果标记完成，不提升整体视觉覆盖等级。
- 不增加渲染引擎、外部依赖或协议版本，不重写 opaque 内容，不自动替换本机默认包。新发现的 3 项语言/大小写写回失败不在本次授权范围内。
- 基线：`tmp/officekit-native-scene-paint-ay3U2S/integration.json` 为一次性 QA 证据，506 个正式入口检查、整体 `failed`；原生包仍来自 `035472e9`，不能据此认定当前持续更新的 C# 源码仍有全部相同缺陷。正式回归必须在 `test/` 和原生测试中可再生成，不能依赖保留 `tmp/`。
