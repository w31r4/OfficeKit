## Purpose

明确 source-bound PPJ 删除 frame 属性、组合段落间距和共享图片背景时的编辑保真要求，使原生属性存在性、源绑定授权和共享资源保留在候选写回及重新导入后保持一致，并提供可重复验证的验收边界。

## ADDED Requirements

### Requirement: Frame removal preserves attribute presence semantics

在既有可编辑 frame 和精确字段授权范围内，系统 SHALL 支持删除 `rotation`、`flipH`、`flipV`，并在候选原生 XML 和重新导入的 PPJ 中保留其缺省状态。删除 SHALL NOT 被转换成显式 `0/false`；显式赋值 `0/false` SHALL 保持显式存在。未修改的 frame 字段和其他对象 SHALL 保留。

#### Scenario: Four original transform deletions

- **WHEN** 从原始 PPTX 独立投影的请求分别删除旋转、水平翻转、垂直翻转，或同时删除三者
- **THEN** 候选对应的原生属性和重新导入的 PPJ 字段 SHALL 真正缺省，其他字段不变

#### Scenario: Presence-only and mixed changes

- **WHEN** 原始属性为显式零/false，或请求将一个属性删除并将另一个属性显式设为零/false
- **THEN** 系统 SHALL 区分值相同但存在性不同的请求，保留每个字段请求的存在状态
- **AND** 从删除后的候选重新投影并恢复该属性 SHALL 得到明确存在的原生值

### Requirement: Paragraph spacing combinations can be removed independently

对已建模且获字段授权的普通 text/shape 段落，系统 SHALL 支持同时删除行距、段前、段后的任意非空子集。每种间距现有的点值/倍数互斥、数值范围与单位精度规则 SHALL 保持；删除整个直接槽位 SHALL 不生成默认或零值占位。未删除的间距、未建模内容、runs 和相邻段落 SHALL 保留。

#### Scenario: Seven spacing subsets

- **WHEN** 三个间距均明确存在，分别从同一原始源独立请求删除七种非空子集
- **THEN** 候选 SHALL 只移除目标间距槽位，重新导入保留正确的单位、值及存在性
- **AND** 未删除槽位及其原生数值拼写 SHALL 保留

#### Scenario: Units and paragraph style container

- **WHEN** 请求对点值、倍数或混合单位段落执行删除、恢复，或删除只含目标间距的段落 style
- **THEN** 候选 SHALL 保留请求的直接属性存在性，不混入另一个单位或重写相邻段落
- **AND** 本次修复 SHALL NOT 放宽原本无效的行距零值或双单位输入

### Requirement: Shared image background removal retains the foreground resource

删除已授权的直接图片背景时，系统 SHALL 根据未修改原始源验证元素身份、语义摘要与能力契约，不得将本次候选修改引起的引用数量变化误判为绑定篡改。背景与前景图片共享媒体和同一关系的情况下，候选 SHALL 只移除直接背景，保留前景图片及其仍使用的媒体和关系，不扁平化或替换对象。

#### Scenario: Remove background sharing the picture relationship

- **WHEN** 源页面的直接图片背景与保留的前景图片共享图片关系，请求只删除背景
- **THEN** 写回应成功，候选原生直接背景 SHALL 缺省，重新导入的前景图片身份、内容和显示属性不变
- **AND** 共享媒体与关系文件 SHALL 逐字节保留；该固定案例只允许目标 slide XML 改变

#### Scenario: Removal does not grant foreground deletion authority

- **WHEN** 一个请求除删除背景外，还试图借引用变化伪造前景图片的源删除能力
- **THEN** 系统 SHALL 根据原始源契约拒绝未授权修改，不发布候选
- **AND** 后续若需使用候选产生的新能力，调用方 SHALL 先重新导入候选取得新的绑定

### Requirement: Deletion retains source integrity and fail-closed behavior

上述删除 SHALL 使用原始源文件和该源签发的字段/背景权限。缺失权限、篡改绑定、错误源字节或未支持的源拓扑 SHALL 拒绝且不产生输出。输入文件与请求 SHALL 保留，source-bound no-op SHALL 保持原文件逐字节相同；未修改部件及目标部件内的非目标内容 SHALL 保留。

#### Scenario: Tampered or unsupported request

- **WHEN** 删除请求缺少对应能力、修改绑定摘要，或试图替换未建模原生内容
- **THEN** 系统 SHALL 明确拒绝并保留输入，不通过重建整页、丢弃未知字段或放宽绑定校验完成请求

### Requirement: Verification separates edit fidelity from visual review

这 12 项删除的完成证据 SHALL 同时包括真实候选的结构检查、使用新投影的编辑保真检查和限定视觉检查。预览 SHALL 消费实际候选，并保留既有视觉限制；看起来相同 SHALL NOT 抵消属性删除失败。验收 SHALL 明确区分当前托管源码、新构建 NativeAOT 包、默认安装包和全套报告，其他失败不得因本变更通过而被隐藏。

#### Scenario: Bounded repair succeeds while other tests fail

- **WHEN** 12 项删除在标识明确的新原生包中全部通过，但其他独立回归仍失败
- **THEN** 报告 SHALL 分别列出这 12 项的通过证据和全套失败，不宣称全套通过或完整 PowerPoint 视觉验收

#### Scenario: Preview option does not change edit results

- **WHEN** 同一删除请求分别打开和关闭候选场景收集
- **THEN** 候选文件 SHALL 一致，场景身份 SHALL 对应真实候选，不修改原始源或制造额外编辑权限
