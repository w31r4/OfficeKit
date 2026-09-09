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
| `pages[]` | 原始页 ID 和显示诊断；只有实际写入时才存在 `file`/`png` 字段 |
| `failures[]` | 失败阶段、可用的页面/元素/文件类型信息和错误消息 |
| `input` | 实际加载的 PPJ 路径与原始字节 hash |
| `compile` | canonical program hash 与实际编译结果字节 hash |
| `sourceBound` / `source` | 是否源绑定；有源字节时记录源路径和 hash，无源为 null |
| `assets[]` | 编译和预览共同使用的资产快照，含 ID/hash/字节数 |
| `implementation` | 预览实现版本和相关源文件 digest |
| `environment` | Node 版本、栅格后端可用状态与已加载的 sharp/libvips 等版本 |
| `renderEvidence` / `visualReview` | `local-svg-preview` / `requires-human`，不是 Office 文件渲染验收 |

`status` 和 `diagnostics` 仍包含旧的元素支持信息，尚未代表全部视觉字段正确；G-11 的完整支持等级修复仍待完成。发布失败会覆盖为 unavailable，但发布成功绝不能用来证明画得正确。

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
