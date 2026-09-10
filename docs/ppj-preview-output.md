# 本地 PPJ 预览的输出和失败证据

文字外阴影复验（2026-09-10）：qCO4bS 使用冻结源码 `text-shadow-snapshot-OenB3z` 与同一 035472e9 原生包，新增 34 项形状文字阴影回归全部通过，正式入口共 473 项；原十二项源删除失败仍使整套 failed。外部九项直接阴影对照为七项通过、两项竖排位置失败，模糊与换行仍有限制。详见[当前文字阴影范围与失败证据](ppj-preview-current-gaps.zh-CN.md#形状文字外阴影字形投影与仍失败的外部定位)。未重建 C#，不宣称完整宿主或人类验收。

文字区域复验（2026-09-10）：已补上当前绘制预设的内部文字矩形和有限自定义边值；p4Au9s 在同一 035472e9 默认原生包下通过新增 78 项作者/源 no-op 回归，正式入口共 409 项。两组外部文字对照分别为 8/8、39/39，通过预设边界阈值；不是完整字体或人类验收。整套仍为 failed，原十二项删除失败保持不变。无 preload 的 presentation 门禁 4/4；详见[当前文字区域差距与证据](ppj-preview-current-gaps.zh-CN.md#形状内部文字区域预设定义与自定义坐标)。 TzaE9j 与旧轮廓对照保留为修复前的历史记录。

`officekit ppj preview` 生成 OfficeKit 自有的静态 SVG/PNG 预览。输出完成只说明文件生成完成，不代表视觉正确、编辑保真或 PowerPoint 宿主验收通过。当前绘制限制仍见 [差距审计](ppj-svg-preview-gap-audit.zh-CN.md)。

## 调用与目录保护

```sh
officekit ppj preview deck.ppj -o preview-new --json
```

`preview-new` 必须不存在，父目录可以存在。已有目录（包括空目录）、文件、符号链接都拒绝复用；本命令没有覆盖选项。再次运行应选择新路径，不能删除或改写输入文件来绕过检查。

普通安全页 ID 继续生成 `<id>.svg` 和 `<id>.png`。包含路径分隔符、Windows 保留名或大小写冲突的页 ID 使用确定性安全文件名，原始 ID 保留在清单中。多个调用使用同一输出目录时，至多一个能够取得该目录。

PNG 使用可选 sharp 依赖。命令不会自动下载依赖或切换到 LibreOffice；`ppj render` 是独立的外部渲染入口。缺少 sharp 时，可生成的 SVG 会保留，但请求的 SVG+PNG 发布返回失败。

本地 preview 现要求匹配的原生 codec 返回 `PresentationPreviewScene` v1，renderer 为 `officekit-native-scene-svg`。编译器负责组件展开、数据映射和有效状态；JS 只绘制该场景，不再用 canonical PPJ 猜测布局。普通 build/check 不请求 scene。旧包不支持时明确报 `preview.scene.missing`，必须配套重建/更新 codec，没有旧绘制回退。2026-09-10 已将冻结源码 `035472e9` 的构建交付到本机默认 linux-x64 包，无测试 preload 的正式 CLI 门禁通过。具体包身份见[当前差距第 3.6 节](ppj-preview-current-gaps.zh-CN.md#36-默认运行时交付与复验2026-09-10)；这不是其他平台或后续源码版本的验收。

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

正式 CLI 已复用 scene painter 和上述发布基础。原生多页可以共享语义根 `$`，因此失败回填保留独立页身份与 `scenePath`；局部栅格/写盘失败不能把其他页的检查子树替换掉。全局依赖失败影响全部页面，以文档级地址记录。

场景路线的输入证据通过共享 receipt 验证器检查真实候选字节、程序、场景摘要和资产，然后提取 JSON 安全的 `scene` 身份；painter 同时记录自己消费的场景身份。publisher 比对二者与 compile 摘要、源绑定模式。缺场景、版本不兼容或不匹配时返回 `preview.output.scene`，可靠性为 failed；此时尚未创建输出目录或加载栅格后端，失败证据位于调用异常 receipt 中，而非落盘清单。原源 hash 与编辑候选 hash 分别记录，不能用未编辑源的场景证明编辑候选正确。

身份匹配只证明输出对应哪次场景，不证明全部绘制正确。正式入口移除内部入口提示，保留实际输入/场景限制及对应警示。真实集成保留四个作者/源候选输入各六类操作故障，以及事实失败红色警示；正式函数现有 409 项限定接入案例（含七十八项文字区域、六十六项多边形、七十五项文字方向、三十六项文字翻转、四十七项文字旋转、三十六项亮度插值、二十二项径向渐变、十六项形状图片填充、九项直接线性背景渐变和十二项图片背景），并有 CLI 子进程回归，范围见当前差距第 2 节。操作故障不反向修改已经生成的 SVG/PNG，最终发布状态以清单或调用异常为准。

能力声明描述真实类型的整体边界，不把“存在绘制分支”当作该类型完全支持。目前没有整个元素/图表类型被声明为 supported；个别受限状态可以通过，例如已验证的显式 contain PNG 映射。裁切、源绑定或其他未支持状态仍会降低实际结果的等级。诊断不修复绘制，已报告的文字、图表、图片等缺口仍需逐项实现。

## 图片警示与验收含义

原始输入检查默认使用 `canonical-svg` 规则。内部 `native-scene-svg` profile 必须提供匹配的已验证场景及实际绘制记录，目前对 registry 中有直接回归的可见性和旋转/镜像映射按字段解除旧 renderer 的误报。可见性要求全部 owner 绑定均隐藏且实际进入隐藏分支；变换要求全部 owner 绑定的原生值与该输入字段一致，且实际成功构造 SVG 变换。缺绘制记录仍保留错误；未知 profile 或身份不匹配明确拒绝。

内部调用 `paintPpjSceneSvg(receipt, { assessInput: true })` 已可将该输入检查合入实际场景检查树。先完成所有页绘制，再检查原始输入；保留原始 `path`，按最近语义 owner 添加真实 `scenePath`，一个 owner 的多个生成节点各保留诊断。未映射的输入限制归到真实页或全局，不静默丢弃；最终 SVG/PNG 警示及发布清单使用合并后的状态。公共资产 ID 按已验证 MIME/hash 对应场景中的资产字节，不重新读取路径。结果另保留 `inputAssessment` 供内部检查；最终清单持久化合并后的 assessment，并非额外复制整个输入树。

上述选项仅控制内部 paint-only 入口。正式函数使用 `renderPpjSceneSvg(receipt)`，强制执行输入合并检查，没有关闭选项。其他旧事实规则未批量解除，未知/未支持状态仍产生限制或失败。公共返回仅含 JSON 安全的场景身份，不暴露 protobuf BigInt 对象图；原始 scene 只传给验证器和 publisher。

组坐标另有 registry 的 `groupCoordinates` 映射：只有某输入组的全部 owner 绑定均为实际成功绘制的原生组，且外框和 childFrame 的 x/y/width/height 与安全转换后的原生 EMU 值逐项相等，才解除 `preview.fact.group-coordinates-ignored`。隐藏组、零子尺寸导致的绘制失败、缺记录或坐标不匹配仍保留错误；缺 registry 映射直接拒绝该 profile。字段整体 partial、未知属性和其他事实规则不因此解除。

`datasetLine` 映射仅解除普通折线图 `data.dataset` 的旧通道忽略错误：全部 owner 绑定必须指向有分类与系列的原生 LINE 节点，完成实际折线构造，且自身与全部祖先的外部变换已成功。JS 不重新解释数据集或编码。缺记录、隐藏、混合归属、通道错误和不支持的平滑折线保留失败；其他图表（包括 vector heatmap）没有借用这条规则。真实输入的 1→2 修改已有像素移动、缺失断段/真实零及显式数据对照；文字、轴、样式、源编辑和其他事实限制继续单独检查，解除误报不等于整体通过。

正式路线的页面 assessment 保留真实原生节点层级。生成节点可以共用语义 ID，但 scenePath 独立；未访问的子节点明确记为 unassessed。节点局部状态不能抵消祖先、页面或全局限制，检查树不代表源编辑授权。

直接线性及居中径向背景渐变与形状/单元格共用色标绘制，SVG/PNG 保留透明度，不自动铺白底。径向映射使用实际几何范围；literal 自定义路径复用现有路径求值与曲线/圆弧极值，未知或退化范围保守失败。特殊两色/首尾同色三色插值保留 `preview.scene.paint.gradient-interpolation` partial 提示。非法或未支持的背景渐变记录页面级 `preview.scene.paint.background` 及原生背景地址，输出可辨认的 unavailable 背景和失败警示；前景及其他页面仍保留。二十二项真实径向案例与剩余范围见[当前差距 G-04](ppj-preview-current-gaps.zh-CN.md#径向渐变实际路径范围和源编辑)，不代表非居中焦点、继承背景或宿主颜色保真已完成。

特殊亮度曲线现已实际绘制：每区间使用 32 段 SVG 插值，RGB 与线性 alpha 分开处理，PPTX 色标不增加。partial 提示描述有限采样精度，不再表示曲线未实现。36 项真实案例验证颜色反向、零透明度、两色改三色和恢复普通插值；无文字形状的统一透明度合并到 `compositing.opacity=0` 已通过原生零值与字节不变的二次编译检查。精度、源编辑范围和剩余限制见[亮度插值记录](ppj-preview-current-gaps.zh-CN.md#渐变亮度插值颜色与透明度分开计算)。

直接图片背景复用图片的资产、裁切与透明度绘制，保留素材透明像素和负裁切留白。平铺及缺少明确数值的旧 alpha 状态仍报不可用。共享图片的整个背景删除目前被原生 capability 校验拒绝；正式入口接入和其余绘制成功不代表该删除已完成。

两项平铺背景测试单独验证失败发布：页面级 `preview.scene.paint.background` 和具体 mode 地址保留，正式调用抛出 `preview.output.incomplete`，留下前景、红色警示、SVG/PNG 和最终清单，异常 receipt 与落盘一致。这些拒绝案例不计入 331 项正常发布路径，也不算平铺渲染成功。完整 cover/contain/focus、非居中渐变及继承背景不能由 signed crop 或直接渐变正例推导为已完成。

文字自身旋转先作用于已锚定的文字块，再叠加外层形状/组变换，不旋转填充和轮廓。显式零与删除分开保留；文本框、带文字形状及单元格有 47 项实际原生/像素/正式入口案例，其中九项独立源修改、清零、删除通过重新投影和非目标保留。非零角度与未解析的 upright 或文字扭曲组合明确 unavailable。完整字体排版、真实文字区域和溢出仍需 review。表格单 run 内换行的两个 opaque 反例另列，不能用其原样保留结果证明已渲染；范围与证据见[当前 G-02](ppj-preview-current-gaps.zh-CN.md#文字自身旋转独立于形状的变换)。

文字绘制还会在独立旋转外层抵消自身及祖先组造成的奇数次镜像，字形可以旋转，但不会随外框变成反字；填充和轮廓仍按原变换绘制。新增 36 项实际形状/单元格及嵌套源编辑检查通过，旧 47 项旋转案例中的四项镜像期望已纠正。这不是完整 upright 或全部原生图表标签支持，也不消除现有文字排版限制；详见[文字翻转范围](ppj-preview-current-gaps.zh-CN.md#文字翻转抵消镜像但保留外层变换)。

horizontal/vertical/vertical270 先确定阅读区域、物理边距映射和锚定，再叠加方向角、独立文字角度及外层变换。显式横排和缺省分别保留，75 项真实方向及源编辑回归通过；多列、其他竖排、upright、扭曲、完整换行和 AutoFit 仍有明确限制，见[文字方向范围](ppj-preview-current-gaps.zh-CN.md#文字方向阅读坐标与物理边距)。

triangle、rtTriangle、trapezoid、parallelogram、chevron 在普通形状、形状图片填充和图片遮罩之间共享调整后的真实轮廓。66 项作者/源回归含 36 项独立修改、清零和删除，保留原生非目标字段与 ZIP 部件；不覆盖其他预设、文字内部区域或继承效果。可见笔画没有直接 join 时，`preview.scene.paint.line-join-inherited` 明确提示 miter 审查近似的尖角范围不确定。外部五个缺省连接案例有两个失败，显式 round 的独立五例通过，不能互相抵消；详见[多边形与外部反例](ppj-preview-current-gaps.zh-CN.md#五种多边形共享轮廓与尚未解决的文字区域)。

形状图片填充沿同一几何边界裁切，不把 stroke-only 路径当作遮罩，轮廓和文字独立于图片 alpha。只有完整原生 imageFill 和已映射几何可走此分支；旧资产身份字段、平铺、未知几何或冲突填充保留 unavailable。十六项作者/源回归验证像素和六项源编辑生命周期，正式结果为 requires-review，仍保留文字等限制。

registry 的 `shapeGeometry` 只解除 `preview.fact.shape-geometry-omitted`：同一输入 owner 的全部绑定都必须是实际完成几何绘制的原生 shape，并且自身及全部祖先变换完成。painter 的 `shapeGeometryScenePaths` 记录完整预设或 literal 路径构造；clip 定义本身不算，未知几何、未映射 adjustment、任何一条失败路径均不产生完整记录。几何先构造、随后文字失败的节点也不能通过。缺记录恢复旧错误，缺 registry 映射拒绝该 profile；canonical profile 不变，填充、文字、效果、opaque 和其他事实限制不受此映射影响。

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
| `preview.scene.missing` / `preview.scene.version` | 原生包缺少场景或版本不匹配；配套重建/更新，禁止回退猜测 |
| `preview.scene.*-mismatch` | 检查程序、候选、scene 或资产身份；绘制前拒绝，无输出目录 |
| `preview.output.scene` | publisher 的场景身份交叉检查失败；查看异常 receipt，此时尚未发布文件 |
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
