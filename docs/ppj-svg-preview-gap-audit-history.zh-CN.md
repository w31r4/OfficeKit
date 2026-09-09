# OfficeKit 预览差距审计：历史验证记录

归档日期：2026-09-09。本文保存本次文档整理前的轮次记录，**所有“本轮”“当前”“最新”均只指原记录发生时，不代表现在的工作区状态**。临时目录可能被清理；表中的测试通过不能直接沿用到变更后的脚本。

当前结论、最新实际运行结果和待办请读 [主审计文档](ppj-svg-preview-gap-audit.zh-CN.md)。本次归档没有重新执行以下历史测试，没有改写历史结果，也不构成发布或人类验收。

原记录中的第 2～9 节指主审计文档的同号章节，并非本附录的章节；其中“第 1.3 节”指当时的最新验证区，现已整体归档于此，不应拿主文档的新第 1.3 节反向解释旧记录。

## 整理前的提交说明

原子提交复核（2026-09-09，基线 `3f390e71`）：本次保存折线图修正，以及内部饼图/环图比例绘制和合成回归。当前 SVG 专项、presentation 分段 4/4、能力矩阵及 strict OpenSpec 检查通过；同步了生成的 PPJ 参考文档摘要。本次没有重跑 NativeAOT 集成，下文两个对象锚点失败仍为已知未解决项。此前真实折线图证据不能用于证明新饼图/环图的编译器集成或宿主保真；生产 CLI 切换与任务完成状态不变。

实施复核（2026-09-09）：基于 HEAD `d891a173` 和已有折线图修正，继续 G-01 3.4。内部 line 新增直接 marker 轮廓映射，修复图框边缘观测点被裁半的问题；增加缺失/多系列/可见性回归及源图表编辑的非目标 ZIP 保留断言，并同步 registry、tasks 和本文。未修改 C#、生产入口、任务勾选、安装包或 Skill 路由。内部 SVG 专项及 presentation 4/4 通过，真实 NativeAOT 集成仍因两个对象锚点反例退出 1。完整执行边界见第 1.3 节。


## 历史原文：归档前的审计基线与各轮验证

以下第一张表是本轮 3.4 实施实际执行结果。折叠区中的文档复核及实施记录均为历史；其中“本次”“当前”“下一项”只指对应轮次。历史脚本通过不代表添加新断言后的当前脚本仍然通过。

| 本轮 3.4 实施 | 实际结果与证明范围 |
| --- | --- |
| 工作区 | HEAD `d891a173`；保留已有 line 实现及文档改动，在其上修复 marker 边缘/轮廓、补回归并同步 registry/tasks/本文；没有改编译器或生产 CLI |
| 内部 SVG 专项 | `node test/ppj-preview-scene-svg.mjs` 退出 0；新增首尾/连续缺失、多系列各自断线、轴隐藏/轴线/标签间隔、marker 轮廓与显式零值、超界观测证据断言；仍为合成场景，不是实际 codec 全覆盖 |
| 常规门禁 | `npm run test:slow -- --segment presentation` 退出 0，4/4 通过；gate-policy、预览摘要与能力矩阵生成器 `--check` 均通过。单跑的内部 SVG 专项也被诊断分段包含，不重复计作两份功能覆盖 |
| 真实生产输出 | `/tmp/officekit-ppj-preview-vl2H8D`；authored `preview/`、`cli-preview/` 各两页，均 complete/partial/failed；`source-preview/` 一页，complete/opaque/requires-review。三者 visualReview 均 requires-human。此路线仍为 canonical JSON，不调用内部 painter |
| 显式 NativeAOT 集成 | `node test/ppj-preview-scene-native.mjs /tmp/officekit-preview-runtime-xZSfO9` 退出 1；最终报告 `status: failed`。两个对象锚点案例均失败，anchors 0/2，实际都为 `(0,0.5) → (1,0.5)` |
| 集成中实际通过的独立检查 | minimum/canonical、1 组组件/显式配对、6 条生成 Bézier 路径、源文字修改、显式坐标连接线、表格网格/像素、源表格移动及重新投影均执行通过；这些不是整条集成通过 |
| 折线图实际通过范围 | authored、source-bound no-op、表格移动后图表不变、源图表数值编辑与重新投影、marker 填充/轮廓/边缘及缺失区间像素均通过。第一系列 `[2,null,0,4,5]→[2,null,3,4,5]`；第二系列 `[5,4,null,0,1]` 不变，各自缺失点独立断线 |
| 保真证据的区别 | 图表数值编辑新增精确 ZIP 修改范围断言：仅 `ppt/slides/charts/chart1.xml` 改变，文件集合、其他 part 和原始源字节不变；表格移动的非目标保留回归继续通过。图表是 literal-data fixture，明确不含 workbook，不能推广为 workbook 联动或删除生命周期通过 |
| 内部产物与看图 | 最终产物 `/tmp/officekit-native-scene-paint-4x0ZVv` 保存 SVG/PNG/diagnostics JSON 和 `integration.json`；后者保留整体失败、独立通过项及源/候选/二进制身份。同轮此前 `/tmp/officekit-native-scene-paint-rAFBXC/source-line-edited-0.png` 已看图确认边缘圆点完整且有直接轮廓；最后只纠正了虚线诊断字段路径并重跑，不把同样案例重复计数。字体、标签和图例布局仍不完整，不是人类校准或宿主保真验收 |
| 二进制边界 | 沿用前轮显式构建包，没有重建当前 HEAD；脚本核验 manifest/hash，不能据此证明此包包含最新 C# 文字效果等功能 |
| 维护检查 | 生成器只读检查、相关 JS 语法和 strict OpenSpec 通过；70 个本地链接存在，任务仍为 7/15，差异格式检查通过。registry 补记有限 line，tasks 追加证据，不提升支持等级或勾选。共享线型限制现区分图表/表格的 dashStyle 与连接线的 lineStyle，marker 虚线有精确路径断言。上一轮独立读者复核不计为本轮功能或视觉验收 |
| 本轮未执行 | 全仓测试、全量 smoke、C# 专项、NativeAOT 重建及双构建验证、proto、Skill 全套检查、外部 Office、第三方 fixture 全覆盖、人类校准与性能基准；未重新核验独立 Skill 评测工作树 |

本轮代码身份如下，后续工作区变动后应重新核验；临时产物可被清理，并非仓库归档的可移植证据包。

| 文件 | SHA-256 |
| --- | --- |
| 生产 `svg-preview.mjs` | `f98c59fb95478610748e4377b906c922156fef0e64898eb1659cad2d1d7986ff` |
| 内部 `preview-scene-svg.mjs` | `2a73506b29f37b7b83d56d4ee052a1a672c649649a7aa12e5aba06acdf09eb92` |
| 内部合成测试 | `d87024bd259e535bf4ee58e347fbeab9c787907c272a875909811df4987fd6df` |
| 真实 NativeAOT 集成测试 | `81a38d62cc8b0ed9cce415674a07b6ec18513c6c8579ffaab37e79c10efa5e30` |
| 指定包 PPJ 二进制 | `9c97c6605c36c9e2b8218fe208a380745cc25b3b94742c4a620eb4ed21a496e9` |

<details>
<summary>历史证据：提交快照、3.4 表格/连接线、基础绘制、场景实施、G-11 和初次审计（不计入本轮新测）</summary>

#### 历史：line 边缘修复前的文档复核

上一轮仅修改本文，内部专项、presentation 4/4、gate-policy 和两个生成器检查通过；独立读者复核后修正了轮次及生产/内部范围的措辞，70 个本地链接有效。生产产物为 `/tmp/officekit-ppj-preview-ZSYGEe`，内部为 `/tmp/officekit-native-scene-paint-CXYpwN`。真实集成整体因锚点失败；当时单系列 source line 编辑仅核对值和原始字节，没有独立非目标 ZIP 断言，看图发现 marker 在 plot 边缘被裁半。该轮未重建、未跑 strict OpenSpec，也不是人类校准。上述 marker 和限定 ZIP 缺口已由本轮 3.4 处理，不能沿用旧状态覆盖最新结果。

#### 历史：内部 painter 提交快照

`d891a173` 保存内部 painter 的在途实现，包含尚未完成的折线图绘制。该提交快照的内部 SVG 专项曾因 marker fixture 的 `ForeignFieldError` 退出 1，之前 presentation gate 也曾在图表诊断断言处失败。当时能力矩阵、同步参考文档和 strict OpenSpec 检查通过。这些是提交时记录，不是目前带工作区修正后的测试结果。

#### 历史：3.4 表格与连接线实施

| 本轮 3.4 验证 | 实际结果与证明范围 |
| --- | --- |
| 修改范围 | 保留既有 table/connector painter，实现无关用例继续执行但锚点失败仍致最终退出 1；增加显式坐标/source-bound 连接线及源表格像素检查；纠正错误注释并同步 registry。没有修改 C#、生产入口或任务勾选 |
| 表格真实绘制 | authored 与 source-bound no-op 均验证 120/240 列宽、40/80 行高、360×40 合并头、可见文字 `0`、右边框及绿色/橙色精确 RGB 像素；不是仅检查文件存在 |
| 表格源编辑 | 从原始源投影重新构造请求，x 从 40 改到 64；候选 native、SVG、颜色像素及再次投影一致；ZIP 文件集合及非目标部件字节不变，原始源字节不变 |
| 连接线有限正例 | 独立显式坐标输入 `(750,320) → (510,120)` 的 native/SVG/终点箭头通过；source-bound no-op 和表格移动后的候选保留该方向。此正例不替代对象锚点反例 |
| 关系硬失败 | 对象锚点原例及目标 y+40 两例全部执行、全部失败：实际仍为 `(0,0.5) → (1,0.5)`。脚本输出 `status: failed`、anchors 0/2，最后抛 AggregateError，退出 1 |
| 其他实际集成 | minimum/canonical、组件/显式配对、6 条生成 Bézier 路径、source-bound no-op 与文字编辑均执行；同一显式运行时，不把这些部分通过算成整条集成通过 |
| 产物与看图 | 最终内部产物 `/tmp/officekit-native-scene-paint-5vZYXO`；看图复核 source-table-moved 页面中的合并表格和有向连接线。只是限定测试输入，不是完整视觉质量或人类校准 |
| 常规门禁 | pure scene SVG、presentation 4/4、gate-policy、两个生成器 `--check`、strict OpenSpec 通过；旧生产路线产物 `/tmp/officekit-ppj-preview-95r73S` |
| 版本身份 | PPJ 二进制沿用 `9c97c6605c36c9e2b8218fe208a380745cc25b3b94742c4a620eb4ed21a496e9`；painter SHA-256 `de853ee2e1da8d5f05c7037afcf238c4a79b7a2dcf2429020c6bc5f222740073`；集成测试 SHA-256 `e39c9ec82ead7b412e24c07ab4ebd31ccdf3b339ccc1eb4eae1994963314ffbc` |
| 未执行/未完成 | 未重建 NativeAOT、运行 C#/proto 专项、完整配对组、全仓测试、外部 Office 或性能验收；G-01 3.3/3.4/4.x/5.x 继续开放。编译锚点修复超出原场景消费计划，已提出扩项确认，尚未修改规划契约或 C# |

#### 历史：锚点失败后的文档复核

| 本轮文档复核 | 实际结果与证明范围 |
| --- | --- |
| 工作区 | 起点 HEAD `5837c51e`；本文、G-01 tasks、registry 和预览测试已有改动，内部 painter/专项测试仍是未跟踪文件。本轮保留这些在途实现，只更新本文，不改代码、任务勾选或安装包 |
| 生产执行链 | `svg-preview.mjs` 仍以 `includeNodeMap: false` 编译，再解析 `compiled.programJson`；没有请求场景或调用内部 painter。3.3、3.4 和后续任务仍未勾选 |
| 常规门禁 | `npm run test:slow -- --segment presentation` 退出 0，4/4 通过；其中诊断套件包含新增的表格/连接线合成 SVG 断言。这不是显式 NativeAOT 集成，也不是全仓测试 |
| 真实生产输出 | `/tmp/officekit-ppj-preview-qA2tmL`；authored `preview/` 与 `cli-preview/` 各两页，发布 complete、支持 partial、可靠性 failed；`source-preview/` 一页，complete、opaque、requires-review。三者 `visualReview` 均为 requires-human |
| 显式 NativeAOT 集成 | `node test/ppj-preview-scene-native.mjs /tmp/officekit-preview-runtime-xZSfO9` 退出 1；对象锚点预期 `(750,320) → (510,120)`，实际 `(0,0.5) → (1,0.5)`。断言位置为当前脚本第 181 行；输入已编译并生成场景，不是 compilerRejected、渲染异常或依赖安装失败 |
| 本次集成实际执行边界 | 失败前已执行 minimum/canonical、组件/显式配对及生成路径检查；第一组 topology 已调用内部绘制/栅格化，随后端点断言失败。此后的表格 SVG/像素断言、第二组目标移动、source-bound no-op/文字/表格移动检查均未执行，不得计为本轮通过 |
| 维护检查 | gate-policy、两个生成器的 `--check` 通过；没有重新生成文件来消除漂移 |
| 文档检查 | 70 个本地链接存在，G-01 计数为 7 已勾选/8 未勾选，`git diff --check` 通过；独立读者能区分整体状态、两条 painter 路线、通过/失败/未执行及锚点根因。按复核意见修正了历史范围、配对覆盖和测试执行顺序的措辞；这不是功能或人类视觉验收 |
| 本轮未执行 | C# 专项、NativeAOT 重建/双构建验证、proto 生成、全量 smoke、全仓测试、外部 Office 渲染、人工视觉校准、Skill 四路线实验及性能基准 |

本次核验的文件身份如下。指定临时二进制沿用前轮构建，不能认为它包含当前 HEAD 的所有图表文字新字段；锚点根因则另由当前 C# 源码核对，见第 3.5 节。

| 文件 | SHA-256 |
| --- | --- |
| `src/ppj/svg-preview.mjs` | `f98c59fb95478610748e4377b906c922156fef0e64898eb1659cad2d1d7986ff` |
| `src/ppj/preview-output.mjs` | `0d5aca1b4c89a5ed43e599250414dedc0312f7076cc928359112759d05f374df` |
| `src/ppj/preview-scene-svg.mjs` | `7cf360d0320f654838bba88d93734869b305a0f01c5a7f5821e744619400e21d` |
| `test/ppj-preview-scene-native.mjs` | `9c1df639a41aecdacfa9096ccc5de59710571a018efee16bdd8fca4ff0f1f354` |
| 指定包 `bin/officekit-ppj-codec` | `9c97c6605c36c9e2b8218fe208a380745cc25b3b94742c4a620eb4ed21a496e9` |

#### 历史：3.3 基础绘制实施

| 本轮 3.3 基础绘制 | 实际结果与证明范围 |
| --- | --- |
| 代码 | 新增内部 `preview-scene-svg.mjs` 和测试，registry 登记并接入诊断 gate；没有改动 `svg-preview.mjs` 的生产路线，也未勾完 3.3 |
| 已映射 | rect/textbox/process、ellipse、decision；形状和文字同时保留；内联 run 的字号、颜色、粗斜体及显式换行；图片原始字节/透明度；group 子坐标、旋转、镜像及 hidden；literal M/L/C/Q/Z 路径按 native viewport 映射 |
| 明确限制 | 文本度量、换行/AutoFit/继承、更多 preset、裁切/mask/effects、路径引用及 arc 命令仍有限制；未知路径整条失败，不跳段连线。原生 chart/table/connector/opaque 等专门绘制、G-11/G-12 正式集成仍未完成 |
| 合成专项 | `node test/ppj-preview-scene-svg.mjs` 通过：实际几何/文字属性、路径坐标、嵌套矩阵、图片字节与显式 0、hidden、未知字段及效果/路径/零子范围失败；冻结输入不变；新进程实际内存绘制不加载 native/raster/旧 facade |
| 真实运行时 | 指定包集成通过，PPJ 二进制 SHA-256 仍为 `9c97c6605c36c9e2b8218fe208a380745cc25b3b94742c4a620eb4ed21a496e9`，不是本轮新构建；minimum/canonical、组件/显式配对、Sunburst/Sankey 生成组和源绑定文字编辑实际进入内部 SVG/PNG |
| 直接绘制断言 | 一对组件/显式输入得到相同 native 位置、颜色、文字 `0` 和 RGB 像素 `[204,85,0]`；共享风险图表保留 `[1,0,0]` 与 missing index `[1]`，该 native 图表尚未绘制；6 条实际生成的 Bézier 路径及 owner 进入 SVG，候选编辑文字进入新画面 |
| 产物及看图 | 最终内部 SVG/PNG `/tmp/officekit-native-scene-paint-yMtf47`；已看图复核 vectors 和 canonical 第 2 页。初版 `textbox` 误报未知几何，按 codec 矩形语义修正后重跑；看图不算两轮盲评或人类校准 |
| 运行中纠正 | 测试最初使用非法颜色对象并漏了矢量图表必需 style，被编译器拒绝；按真实契约补正，没有放宽 compiler 或记成绘制成功。合成测试共享路径被前次冻结，改为每例独立克隆，没有放松输入不可变性 |
| 常规门禁 | presentation 4/4 通过，生产旧路线产物 `/tmp/officekit-ppj-preview-8TMpOD`；gate-policy、两个生成器 `--check`、语法、strict OpenSpec、差异检查通过 |
| 未验收 | 没有本轮 C#/proto 修改或 NativeAOT 重建；未做完整 paired suite、生产新路线发布、全量 smoke/全仓测试、宿主验收、人类校准和性能。3.3/3.4/4.x/5.x 与 G-01 整项继续开放 |

#### 历史：上一轮仅文档复核

| 本轮文档复核 | 实际结果与证明范围 |
| --- | --- |
| 代码与任务基线 | 起点 `aee69e35`，复查读到 `8b5b6771`；G-01 清单 7 项完成、8 项未完成。已有适配层由另一工作流提交，本轮没有实现、提交或推送代码，也未修改任务勾选 |
| 真实执行链 | `renderPpjToSvg` 调用 compile 时仅给 `includeNodeMap: false`，然后解析 `compiled.programJson`；没有请求 `includePreviewScene`，也没有导入 `createPpjSceneView`。组件猜测和旧图表分支仍执行 |
| 适配边界 | workspace/native 已有默认关闭的场景转发；`preview-scene-view.mjs` 提供单位、身份、资产和树结构视图，明确返回 `paintAssessment: unassessed`。字段被保留不等于字段已画到 SVG |
| 常规预览回归 | `npm run test:slow -- --segment presentation` 4/4 通过，覆盖诊断（含场景传输与适配）、发布故障、真实 SVG/PNG、能力声明。这不是全部 slow、全仓测试或全部视觉功能通过 |
| 实际输出 | 临时证据根目录 `/tmp/officekit-ppj-preview-ATnliz`；`preview/` 和 `cli-preview/` 各两页，均 `output.status: complete`、`status: partial`、`reliability.status: failed`；`source-preview/` 一页，发布 complete、opaque、requires-review。三份清单均 `visualReview: requires-human` |
| 维护检查 | `node test/gate-policy.mjs` 和两个能力生成器的 `--check` 均通过；没有通过重新生成或修改 registry 消除漂移 |
| 文档复核 | 63 个本地链接全部存在，任务计数与 7/15 一致，`git diff --check` 通过；独立读者复核通过，已修正历史验收范围与过时下一步，未把读者复核当作功能测试 |
| 绘制文件身份 | renderer SHA-256 仍为 `f98c59fb95478610748e4377b906c922156fef0e64898eb1659cad2d1d7986ff`，publisher 为 `0d5aca1b4c89a5ed43e599250414dedc0312f7076cc928359112759d05f374df`；与下方 G-11 快照一致 |
| 本轮未执行 | C# 专项、NativeAOT 构建或指定临时包集成、proto 生成检查、全量 smoke、全仓测试、外部 Office 渲染、人类视觉校准、Skill 实验及性能基准。下方 39/39、真实场景传输等为历史证据，不记成本轮新测 |

上述目录只是本机临时证据，不是已归档的跨环境复现包。`preview/` 与 `cli-preview/` 使用同一个 authored 案例，不是两个新增视觉场景；测试能正确断言 failed，也应算测试通过，不能把这种“通过”记为画面已正确。

#### 历史：3.2 轻量适配实施

| 历史 3.2 实施复核 | 实际结果与证明范围 |
| --- | --- |
| 代码范围 | 新增 `preview-scene-view.mjs` 和专项测试，在共享 registry 登记适配模块、测试及生成 descriptor 来源，并接入常规诊断测试；没有第二份手写内容类型台账，也没有更改 SVG 路由 |
| 全字段与单位 | 直接引用冻结后的完整 native 对象；保留 runs、数据通道、缺失索引、effects 和 opaque 描述。只转换 EMU/角度/透明度/字号，保留显式 0/false 与未给出状态；不安全的整数转换明确报错 |
| 身份与树 | 场景路径与语义 owner 路径分开，保留生成节点归属；group 子坐标和 diagram 缓存绘制树不扁平化。空页没有可证明的语义 ID 时明确 unmapped，不按页序号猜身份 |
| 专项测试 | `node test/ppj-preview-scene-view.mjs` 通过：全部 9 类 native content、有向端点、资产映射、原消息字节不变、未知字段/oneof/枚举及空页限制；root 和实际内存适配在两个全新进程验证惰性依赖 |
| 真实运行时适配 | `node test/ppj-preview-scene-native.mjs /tmp/officekit-preview-runtime-xZSfO9` 通过，沿用下表 PPJ 二进制摘要。minimum/canonical 和 source-bound no-op/文字编辑进入适配，核对实际位置、候选文字、资产和 native 对象身份；不是 SVG 像素断言 |
| 常规回归 | presentation 4/4 通过，旧路线真实产物 `/tmp/officekit-ppj-preview-hMJ6U2`；补充空页限制后又完整重跑 4/4，通过产物 `/tmp/officekit-ppj-preview-O9lE46`，不重复计算覆盖数。真实 NativeAOT 适配再次通过；gate-policy、两个能力生成器 `--check`、语法、strict OpenSpec、63 个本地链接和差异检查通过 |
| 状态与未执行 | 3.2 已勾选，7/15；未修改 C# 或重新构建运行时，本轮未重跑原生 39 项、proto 生成、全量 smoke、全仓测试、宿主验收、人类校准或性能。真实 SVG、完整 G-11/G-12 场景对接及等价仍待 3.3 以后完成 |

#### 历史：3.1 传输验收

| 历史 3.1 实施复核 | 实际结果与证明范围 |
| --- | --- |
| 重建状态 | 重新轮询此前构建句柄，确认仓库 `build:office-kit` 命令退出 0；SDK 8.0.128，独立输出 `/tmp/officekit-preview-runtime-xZSfO9`，没有覆盖安装包。manifest 为 schema 2、package 2.0.0、wire/transport 2、linux-x64 |
| 实际 PPJ 可执行文件 | `bin/officekit-ppj-codec`，SHA-256 `9c97c6605c36c9e2b8218fe208a380745cc25b3b94742c4a620eb4ed21a496e9`；测试加载器核验 manifest 和实际文件身份，不以当前源码 HEAD 代替二进制身份 |
| 真实跨语言回归 | `node test/ppj-preview-scene-native.mjs /tmp/officekit-preview-runtime-xZSfO9` 通过。minimum/canonical 两个 authored 输入，以及带资产的 source-bound no-op、文字叶编辑；轻量和生成 wire 读法场景相等，实际 canonical/candidate/scene/asset 摘要通过；源输入不变，普通 build/check 无场景 |
| 路由边界纠正 | 初稿调用打包 `office` profile 编译 PPJ，按设计收到 `unsupported_operation`。该包只处理 Word/Excel，不等于 C# 通用 `CodecProtocol`。修正测试去实际 PPJ 二进制，并断言 Office 拒绝；没有扩大 Office 包依赖，也移除了此前为该错误路线添加但不需要的解码覆盖 |
| 原生专项 | SDK 8.0.128 重跑 `FullyQualifiedName~PpjPreview`，39/39 通过、0 跳过，覆盖通用 `CodecProtocol` 与 `PpjCodecProtocol` 库入口；此源码测试与临时包集成是两项独立证据 |
| JS 与维护回归 | presentation 分段 4/4 通过，真实旧路线产物 `/tmp/officekit-ppj-preview-GmzKQ2`；gate-policy、两个能力生成器 `--check`、集成脚本语法检查和 strict OpenSpec 通过；`npm run proto:check` 完成 lint、生成及差异检查，退出 0 |
| 任务状态 | 3.1 已勾选，6/15；3.2～5.4 仍未勾选。此结果不证明新场景已进入 SVG，不证明任务 5.2 的双构建复现及整条新绘制链路完成 |
| 修改及未执行范围 | 修正真实集成测试，保留传输负例，并更新本文和 tasks；没有提交、推送或更换安装包。未运行完整 smoke、全仓测试、实际 scene-to-SVG、宿主视觉验收、人类校准或性能基准 |

#### 历史：仅文档复核（HEAD `ae8f9404`）

该轮只编辑本文，未修改运行时代码、任务勾选、生成文件或 Skill，也未提交或推送；当时没有核验后台构建结果。以下保留当时记录。

| 本轮文档复核 | 实际结果与证明范围 |
| --- | --- |
| 代码基线 | HEAD `ae8f9404`；工作区另有 `src/codecs/office-kit-runtime.mjs`、`test/ppj-preview-scene-transport.mjs` 的已有修改及未跟踪的 `test/ppj-preview-scene-native.mjs`，全部保留。本文描述此工作区，不推断 origin/main 或安装包同步状态 |
| 当前绘制输入 | `renderPpjToSvg` 调用 `compile(workspace, { includeNodeMap: false })`，随后解析 `compiled.programJson`；没有开启 `includePreviewScene`。绘制代码仍含组件启发式和第 3～4 节的语义缺口 |
| 当前传输代码 | workspace/native 转发默认关闭的 scene 选项；专用 wire 保留场景消息字节，场景模块延迟解码。校验版本、来源、实际 canonical/candidate 字节摘要、场景摘要、节点绑定和资产 MIME/hash；拒绝源编辑授权及 opaque XML 进入只读场景 |
| G-01 清单 | 5/15 勾选；3.1 有实现和合成传输测试，但真实重建运行时验收仍待确认；3.2～5.4 未勾选，不在本次文档任务中提前改状态 |
| presentation 专项 | `npm run test:slow -- --segment presentation` 4/4 通过，包含场景传输、诊断、发布故障、真实 SVG/PNG 和能力声明检查；不是全仓测试，也不是新场景绘制验收 |
| 本轮实际图像产物 | `/tmp/officekit-ppj-preview-9zAt7n`，来自实际 codec + sharp 的 authored/source-bound 回归；仍为旧绘制路线的本机临时证据，不是第三方 PPTX 验收包 |
| 维护检查 | `node test/gate-policy.mjs` 通过；预览能力摘要和演示能力矩阵两个生成器的 `--check` 通过；声明覆盖测试确认 16 类元素、16 类图表 |
| 未执行或未确认 | 本轮未运行 C# 专项、NativeAOT 构建/双构建复现、真实 scene-native 集成、全量 smoke、全仓测试、proto 生成、外部 Office 视觉对照、人类校准或性能基准。不判断其他后台构建是否成功，测试文件存在也不算执行通过 |

#### 历史：G-01 2.3 实施复核

以下记录的是原生归属任务完成时的结果，早于 `ae8f9404` 的 JavaScript 传输提交。

| 历史 2.3 实施复核 | 2026-09-09 记录 |
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

初次审计的测试使用当时运行时能加载的 codec，未重建 NativeAOT，不能将那次预览测试视为工作区 C# 改动已经生效的证据。后续指定包重建与传输验证见本折叠区的 3.1 记录；本文不声称所有轮次都已固定字体、原生二进制和资源环境。

</details>

### 历史：2026-09-09 提交快照复核

基于 `654cb3e5` 的独立提交快照已验证：原生预览专项 39/39 通过；JS 预览诊断及新增场景传输测试、`proto:check`、gate-policy、OpenSpec 严格校验和差异格式检查通过。该次包含原生归属实现和 JS 默认关闭的场景传输、身份与资产校验。传输实验使用合成回执和替代原生调用，只证明传输契约；当时尚未重建 NativeAOT，也未验证新场景的 SVG 绘制，任务 3.1 与 G-01 保持开放。以上仅属于该次提交检查，不属于第 1.3 节的最新文档复核；前文“未提交”等描述保留为各轮历史快照。

## 历史：a3673cea 文档复核（早于 circular fixture 修正及径向分离）

下表是上一轮文档任务的原记录。全缺失 fixture 和分段失败此后已修复，当前结果请读主文档；独立读者复核只验证当时的文档表达。

### 当时的基线与结果

本表只记录这次文档任务实际核对和运行的结果，覆盖此前“4/4 通过”等最新状态表述。[历史验证记录](ppj-svg-preview-gap-audit-history.zh-CN.md) 保留旧轮次的命令、产物和摘要，不将其计入本次执行。

| 检查项 | 2026-09-09 本次实际结果 | 证明范围与限制 |
| --- | --- | --- |
| 代码快照 | HEAD `a3673cea`；已有未提交修改仅涉及内部 painter 和其合成测试 | 本文按这个工作区核对；不推断 origin/main、发布包或已安装运行时同步状态。本次只修改审计文档及历史附录 |
| G-01 任务 | 1.1、1.2、2.1、2.2、2.3、3.1、3.2 已勾选，7/15 | 3.3、3.4、4.1、4.2、5.1～5.4 未勾选；任务数不是绘制覆盖率 |
| 正式入口 | `renderPpjToSvg` 编译时仍只传 `includeNodeMap: false`，随后解析 `compiled.programJson` | 未请求 `includePreviewScene`，不调用内部 painter；正式 CLI 缺陷见第 3～4 节 |
| 内部 SVG 合成专项 | `node test/ppj-preview-scene-svg.mjs` 退出 1 | 第 363 行全缺失饼图断言失败；详见第 2.6 节。执行到失败点之前的断言成功，不等于整个脚本通过；后续测试未执行 |
| presentation 常规分段 | `npm run test:slow -- --segment presentation` 退出 1，停在第 1/4 步诊断套件 | 同一合成测试失败；第 2～4 步未由此分段执行。不能继续写“本轮 4/4 通过” |
| 单独执行被阻断的检查 | `test/ppj-preview-output-evidence.mjs`、`test/ppj-svg-preview.mjs`、`test/ppj-preview-capability-coverage.mjs` 分别退出 0 | 这是单独运行的三项结果，不把失败分段重记为通过；发布保护和现有正式路线回归保持通过 |
| 正式路线产物 | `/tmp/officekit-ppj-preview-2Gqkj2` | 真实 codec + sharp；authored 和 CLI 样例为 complete / partial / failed，简单 source-bound 为 complete / opaque / requires-review；人工视觉状态仍 requires-human |
| 显式 NativeAOT 集成 | `node test/ppj-preview-scene-native.mjs /tmp/officekit-preview-runtime-xZSfO9` 退出 1 | 报告 `status: failed`；两个对象锚点案例均失败，0/2。实际都为 `(0,0.5) → (1,0.5)`，详见第 3.5 节 |
| 集成实际通过的基础项 | minimum/canonical、1 组组件/显式配对、6 条生成 Bézier 路径、源文字修改、显式坐标连接线、合并表格像素及源表格移动重新投影 | 有 native/SVG/像素或重新投影断言；不同项目证据不同，不等于所有案例都有完整视觉与编辑生命周期覆盖 |
| 集成实际通过的普通 line | authored、source no-op、表格移动后图表不变、源数值修改、重新投影、两系列独立缺失、marker/轮廓/边缘与缺失区间像素 | 第一系列 `[2,null,0,4,5]→[2,null,3,4,5]`，第二系列不变；literal-data 输入，不含 workbook |
| 集成实际通过的 pie / doughnut | 每类 authored、比例交换、旋转、source no-op、源数值修改和再次投影；比例颜色像素通过，doughnut 透明孔像素通过 | `[1,null,0,9]→[9,null,0,1]`；角度 0/90，环孔 60/40。两类各只改变 `ppt/slides/charts/chart1.xml`，其余 ZIP 成员和源字节不变；详见第 2.6 节 |
| 内部产物 | `/tmp/officekit-native-scene-paint-l89Na5` | 保存 SVG、PNG、逐案例 diagnostics 和 `integration.json`；报告同时保留独立通过与整体失败。临时产物不是已归档的可移植 fixture 包 |
| 使用的二进制 | 沿用显式临时构建包，加载器检查 manifest/hash | 本次没有重建 HEAD；当前 C# 新接口不能仅凭这份旧包的成功用例视作已验收 |
| 维护检查 | gate-policy、两个能力生成器 `--check` 分别退出 0 | 证明声明/生成文件一致；registry 的内部进展文字仍只列 line，tasks 也未补这次 circular 集成结果，见第 6.3 节 |
| 文档检查 | 主文档 77 个、历史附录 1 个本地链接均存在；差异空白检查通过；独立读者复核后修正历史导航和编译扩项的行动边界 | 读者复核只验证文字清楚、范围一致，不是功能测试或人类视觉校准；已有源码/测试差异保持原样 |
| 本次未执行 | 全仓测试、全量 preview smoke、C# 专项、NativeAOT 重建与双构建验证、proto、Skill 全套门禁、外部 Office、人类校准和性能基准 | 不沿用历史结果来填补；也未重新核验独立 Skill 评测工作树 |

实际代码身份如下。之后文件发生变化，必须重新运行对应检查；只记录 HEAD 不足以覆盖未提交修改。

| 文件 | SHA-256 |
| --- | --- |
| 生产 `svg-preview.mjs` | `f98c59fb95478610748e4377b906c922156fef0e64898eb1659cad2d1d7986ff` |
| 内部 `preview-scene-svg.mjs` | `61dfc4af7304e4c15b0b6c2ee45873849b9a5355bd3f9aef74c2e996f447534d` |
| 发布 `preview-output.mjs` | `0d5aca1b4c89a5ed43e599250414dedc0312f7076cc928359112759d05f374df` |
| 内部合成测试 | `c61312350a7129f1c9dc5c5d2770942658f8255b682decfa4510b093d3043246` |
| 真实 NativeAOT 集成测试 | `93c1f9b229e8a4d867faef31d89375f3e7d82f6def4e24572895cf72af8b4826` |
| 指定包 PPJ 二进制 | `9c97c6605c36c9e2b8218fe208a380745cc25b3b94742c4a620eb4ed21a496e9` |

证据分三类：**运行确认**表示指定案例在本轮实际执行；**代码确认**表示可直接定位的实现；**待验证**表示没有足够的案例证明。负例按预期被拒绝可以是测试通过，但不是正向绘制成功；反之，合成 fixture 不合法造成测试失败，也不能直接推断所有合法输入都画错。

术语：`authored` 指从结构化输入创建文稿；`source-bound` 指编辑绑定原始 PPTX、原生对象与所有权证据；`opaque` 指无法完整安全建模的原生内容；`canonical JSON` 是规范化程序文本，不承诺已展开布局；`fixture` 是固定测试输入；`smoke` 只检查基本运行路径。
