# 本地 PPJ 预览的输出和失败证据

`officekit ppj preview` 生成 OfficeKit 自有的静态 SVG/PNG 预览。输出完成只说明文件生成完成，不代表视觉正确、编辑保真或 PowerPoint 宿主验收通过。当前绘制限制仍见 [差距审计](ppj-svg-preview-gap-audit.zh-CN.md)。

## 调用与目录保护

```sh
officekit ppj preview deck.ppj -o preview-new --json
```

`preview-new` 必须不存在，父目录可以存在。已有目录（包括空目录）、文件、符号链接都拒绝复用；本命令没有覆盖选项。再次运行应选择新路径，不能删除或改写输入文件来绕过检查。

普通安全页 ID 继续生成 `<id>.svg` 和 `<id>.png`。包含路径分隔符、Windows 保留名或大小写冲突的页 ID 使用确定性安全文件名，原始 ID 保留在清单中。多个调用使用同一输出目录时，至多一个能够取得该目录。

PNG 使用可选 sharp 依赖。命令不会自动下载依赖或切换到 LibreOffice；`ppj render` 是独立的外部渲染入口。缺少 sharp 时，可生成的 SVG 会保留，但请求的 SVG+PNG 发布返回失败。

## 目录内的证据

| 文件 | 含义 |
| --- | --- |
| `render.pending.json` | 开始发布前写入的未完成标记；不能当作成功结果 |
| 每页 SVG/PNG | 独占写入的实际产物；失败页可能只有 SVG |
| `render.json` | 结束时写入的最终清单，包括成功与不完整两种状态 |

成功写入最终清单后，会尝试删除本次 pending 标记。删除失败时，最终清单仍为准，调用结果另带 cleanup warning。若没有最终清单，不能从“有若干图片”推断任务完成。

失败时保留已写文件，便于根据错误修复并重跑到新目录。不会递归删除输出目录。文件系统完全不可写、进程被杀或写清单时失败，可能只有 pending 或没有可落盘的失败记录；调用错误不能被当成成功。

## 最终清单字段

schema 为 `office-kit/ppj-svg-preview-output/v1`。

| 字段 | 说明 |
| --- | --- |
| `ok` / `output.status` | 发布是否完成；状态为 `complete` 或 `incomplete` |
| `output.directory` | 本次输出绝对路径 |
| `artifacts[]` | 实际成功写入的文件，含 `pageId/kind/file/sha256/bytes` |
| `pages[]` | 原始页 ID、`status`、`reliability`、`diagnostics` 和 `assessment`；只有实际写入时才存在 `file`/`png` 字段 |
| `assessment` | 整个程序的检查树，`children` 保留页面及嵌套元素；联合 `pageId/id/path` 和可选 `scenePath` 区分身份，不能只用局部元素 ID 或共享的 `$` 路径 |
| `status` | 按实际已检查状态聚合：`unavailable > opaque > partial > supported`，与诊断数量无关 |
| `reliability` | 独立可靠性门槛：`failed`、`requires-review` 或 `passed`，含错误级 `violations` |
| `diagnostics[]` | 含 `pageId/id`（适用时）、具体 `path`、可选原生 `scenePath`、`status/reason/severity/valueSummary/action`；源载荷不展开，值摘要限长 |
| `failures[]` | 失败阶段、可用的页面/元素/文件类型信息和错误消息 |
| `input` | 实际加载的 PPJ 路径与原始字节 hash |
| `compile` | canonical program hash 与实际编译结果字节 hash |
| `scene` | 仅场景路线：`version/origin/sha256/programSha256/candidateSha256`；origin 为 `authored-lowering` 或 `candidate-import`，不包含原生场景载荷或编辑权限 |
| `sourceBound` / `source` | 是否源绑定；有源字节时记录源路径和 hash，无源为 null |
| `assets[]` | 编译和预览共同使用的资产快照，含 ID/hash/字节数 |
| `implementation` | 预览实现版本和相关源文件 digest |
| `environment` | Node 版本、栅格后端可用状态与已加载的 sharp/libvips 等版本 |
| `renderEvidence` / `visualReview` | `local-svg-preview` / `requires-human`，不是 Office 文件渲染验收 |

支持与可靠性检查已经接入真实 SVG 绘制和发布。字段限制、继承/全局状态、未知视觉后代以及已登记的事实错误汇入同一份检查结果。分支自报 supported 不能覆盖字段限制；错误级事实诊断或生产失败使可靠性为 failed。仅有 partial/opaque 限制时为 requires-review；passed 只说明这一套自动检查没有发现限制，仍不等于人工视觉或编辑保真通过。

内部 scene painter 已复用上述发布基础，但正式 CLI 尚未切换。原生多页可以共享语义根 `$`，因此失败回填必须保留独立页身份与 `scenePath`；局部栅格/写盘失败不能把其他页的检查子树替换掉。全局依赖失败影响全部页面，以文档级地址记录。

场景路线的输入证据通过共享 receipt 验证器检查真实候选字节、程序、场景摘要和资产，然后提取 JSON 安全的 `scene` 身份；painter 同时记录自己消费的场景身份。publisher 比对二者与 compile 摘要、源绑定模式。缺场景、版本不兼容或不匹配时返回 `preview.output.scene`，可靠性为 failed；此时尚未创建输出目录或加载栅格后端，失败证据位于调用异常 receipt 中，而非落盘清单。原源 hash 与编辑候选 hash 分别记录，不能用未编辑源的场景证明编辑候选正确。

身份匹配只证明输出对应哪次场景，不证明全部绘制正确。内部警示和未支持字段诊断仍保留。真实集成已覆盖四个作者/源候选输入各六类操作故障，以及作者和源内容事实失败的红色警示；正式接入尚需继续验收。操作故障不反向修改已经生成的 SVG/PNG，最终发布状态以清单或调用异常为准。

能力声明描述真实类型的整体边界，不把“存在绘制分支”当作该类型完全支持。目前没有整个元素/图表类型被声明为 supported；个别受限状态可以通过，例如已验证的显式 contain PNG 映射。裁切、源绑定或其他未支持状态仍会降低实际结果的等级。诊断不修复绘制，已报告的文字、图表、图片等缺口仍需逐项实现。

## 图片警示与验收含义

原始输入检查默认使用 `canonical-svg` 规则。内部 `native-scene-svg` profile 必须提供匹配的已验证场景及实际绘制记录，目前对 registry 中有直接回归的可见性和旋转/镜像映射按字段解除旧 renderer 的误报。可见性要求全部 owner 绑定均隐藏且实际进入隐藏分支；变换要求全部 owner 绑定的原生值与该输入字段一致，且实际成功构造 SVG 变换。缺绘制记录仍保留错误；未知 profile 或身份不匹配明确拒绝。

内部调用 `paintPpjSceneSvg(receipt, { assessInput: true })` 已可将该输入检查合入实际场景检查树。先完成所有页绘制，再检查原始输入；保留原始 `path`，按最近语义 owner 添加真实 `scenePath`，一个 owner 的多个生成节点各保留诊断。未映射的输入限制归到真实页或全局，不静默丢弃；最终 SVG/PNG 警示及发布清单使用合并后的状态。公共资产 ID 按已验证 MIME/hash 对应场景中的资产字节，不重新读取路径。结果另保留 `inputAssessment` 供内部检查；最终清单持久化合并后的 assessment，并非额外复制整个输入树。

该选项目前默认关闭，供内部接入验收；正式 CLI 仍未切换。它不代表页面通过：其他旧事实规则没有被批量解除，未支持的输入或场景字段仍产生限制或失败。

组坐标另有 registry 的 `groupCoordinates` 映射：只有某输入组的全部 owner 绑定均为实际成功绘制的原生组，且外框和 childFrame 的 x/y/width/height 与安全转换后的原生 EMU 值逐项相等，才解除 `preview.fact.group-coordinates-ignored`。隐藏组、零子尺寸导致的绘制失败、缺记录或坐标不匹配仍保留错误；缺 registry 映射直接拒绝该 profile。字段整体 partial、未知属性和其他事实规则不因此解除。

内部 scene 路线的页面 assessment 下另保留真实原生节点层级。生成节点可以共用语义 ID，但 scenePath 独立；未访问的子节点明确记为 unassessed。节点局部状态不能抵消祖先、页面或全局限制，检查树不代表源编辑授权。正式路线尚未切换到这棵原生检查树。

事实错误或绘制不可用的页面显示红色 `UNRELIABLE PREVIEW`；只有未实现/不透明限制的页面显示棕色 `PREVIEW REQUIRES REVIEW`。警示在栅格化前加入 SVG，PNG 使用同一份 SVG，保留原画布尺寸与元素 ID。警示覆盖页面顶部的一小条区域，是审查标记，不是修改输入中的版式或新增可编辑页面对象。

SVG 警示包含 `data-officekit-assessment-path`、`data-officekit-diagnostic-path`、`data-officekit-diagnostic-reason`，可对应清单中的具体检查结果。PNG 只能显示简短提示；完整原因与建议看 `render.json`。警示反映绘制时已知的限制；随后发生的磁盘/清单写入失败仍须查看最终清单或异常 receipt，不能从已写图片反推整个发布成功。

以下组合是合法且有意保留的：

```json
{
  "ok": true,
  "output": { "status": "complete" },
  "status": "partial",
  "reliability": { "status": "failed" },
  "visualReview": "requires-human"
}
```

示例省略了实际错误内容；真实 failed 事实诊断会出现在 violations 中。含义是“文件写完，但事实表达错误，不能当正确性证据”。CLI 退出码当前仍只表达编译/绘制不可用/发布错误；事实错误诊断可以与退出码 0 并存。调用者必须读取 `reliability`，不能只看退出码或 `ok`。事实失败不能被外观评分抵消；requires-review/passed 也不自动替代结构检查、人工渲染检查或导入后的编辑保真检查。

输入/资源 hash 来自加载的同一份字节，不在编译后重读资产路径。无需输出目录的内部 `renderPpjToSvg` 调用仍返回内存 SVG，不加载 PNG 依赖。当前 CLI 仍保留原有返回布局；完整 CLI 摘要及页选择修复属于 G-13。

## 错误处理

发布错误使用 `PreviewOutputError`，包含 `code`、`outputPath` 和内存 `receipt`。最终清单无法落盘时，内存 receipt 仍提供已经捕获的失败信息；CLI 通过既有异常处理返回非零退出码。

| code | 处置 |
| --- | --- |
| `preview.output.exists` | 已有目标路径受到保护；选择新输出路径 |
| `preview.output.reserve` | 检查父目录、权限或路径错误 |
| `preview.output.pending` | 尚未发布页面；检查初始清单写入错误 |
| `preview.output.incomplete` | 查看最终清单中的 `failures`，修复依赖、资源、栅格或写入问题后重跑 |
| `preview.output.manifest` | 最终清单未成功写入；现有图片仅为未完成产物，检查磁盘/权限等错误 |

`raster-load` 与 `raster-render` 分开记录；后者不是缺依赖的同义词。生成 PNG 字节不等于成功写入文件，`artifact-write` 失败不会把文件名写进成功产物列表。

这里提供的是非覆盖文件发布和明确的完成记录，不是断电后的 fsync 持久性保证，也不承诺防御恶意进程在运行中替换整个目录树。

## 回归检查

```sh
npm run test:ppj-preview-output
node test/ppj-svg-preview.mjs
```

第一项是注入栅格/磁盘故障的轻量测试，包含非覆盖、并发、路径、缺依赖、部分写入、hash 和资源快照；已加入 fast/slow gate。其模拟 PNG 字节只测试发布，不用于证明栅格正确。第二项实际调用 codec 和 sharp，包含 authored 与移除私有 PPJ 快照后的 source-bound 投影、原输入保持和产物 hash 校验。

能力声明只编辑 `src/ppj/capability-registry.json` 的 `previewSupport`，再运行 `node scripts/generate-ppj-preview-capabilities.mjs` 与 `npm run docs:presentation-capabilities`；两项生成器均支持 `--check`，检查不应改写产物。诊断及集成回归用 `npm run test:ppj-preview-diagnostics` 验证；`npm run test:slow -- --segment presentation` 包含该测试、发布故障、真实预览和能力覆盖四项。真实预览测试核对返回/落盘检查一致性、源字节、带警示文件 hash 和 PNG 警示像素；字段级完整绘制覆盖仍需后续任务补齐。
