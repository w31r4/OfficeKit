# 命名样式 / grammar 与显式样式对照

输入为同目录 `preview-style-grammar-equivalence.json`，以 `examples/ppj/minimum.ppj` 为基础，由 `test/ppj-preview-scene-native.mjs` 通过真实 NativeAOT 编译并绘制。

## 独立输入与比较范围

命名侧通过 shape/text `styleRef` 引用样式，grammar 提供橙色填充、蓝色文字、DejaVu Sans 字体和 32pt 字号。命名文本默认粗体 true，元素明确覆盖为 false。显式侧独立写出最终值，不从命名侧编译结果反向生成。两侧使用相同几何和文字 `Signal 0`。

两侧都包含 `[1,null,0]` 折线：两个孤立观测必须可见，中间缺失不能补零或连线。比较全部有序原生视觉载荷字节和整页 raw raster；外层节点身份由原始元素 ID、programPath 和几何断言单独验证，不将 scene 转回 PPJ。

测试另外检查：字体、字号、false 覆盖；矩形内部橙色像素；文字区域超过 100 个精确蓝色墨迹像素；实际 SVG 字体/字号；两侧各自 scene 开关候选字节一致；原输入不变；fixture 与实现摘要在运行前后相同。任一侧失败均留存报告，并令整套集成失败，另一侧与独立回归继续执行。

## 边界

这是作者输入中有限命名样式与 grammar token 的等价性测试，不是完整 theme/master/layout 继承、源编辑/删除、字体替代、全平台字体度量或人类视觉校准。两侧共享本机栅格环境；相同像素不证明 PowerPoint 像素一致。

grammar 文字颜色使用独立的 `label-ink` 名称。初稿误用基础主题已有的 `ink`，按 Catalog 的主题优先规则得到 `172B4D`，真实像素检查拒绝了错误的显式答案 `114477`。修正的是 fixture 名称，未修改解析优先级或放宽颜色断言。该次失败产物保留在实施报告指明的临时目录。
