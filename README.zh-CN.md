# OfficeKit

## 把一句需求变成可以交付的 Office 文件

[English](README.md) | **简体中文**

给 Agent 数据、现有文件和一句需求，拿回能打开、能继续编辑、能交付的
Word、Excel、PowerPoint 或 PDF。

```text
→ 一个入口，自动选择 Word、Excel、PPT 或 PDF 工作流
→ 复用合适模板，也能从零设计
→ 修改已有文件时保留未触及的复杂内容
→ 交付前重新打开、渲染并检查
```

## 看它怎么工作

```text
你：
使用 OfficeKit。用 sales.xlsx 做一份给管理层看的 Q2 经营复盘 PPT，
必须包含收入、毛利、区域差异、主要风险和三项决策；模板由你判断。

OfficeKit：
读取数据 → 确定 PPT 工作流 → 选择模板或不用模板
→ 生成文件 → 重新打开 → 渲染检查 → 返回 q2-review.pptx
```

也可以直接提出单一格式任务：

```text
使用 OfficeKit，把这些 CSV 做成带公式、图表和异常标记的 Excel 经营看板。
使用 OfficeKit，沿用 template.pptx 和 data.xlsx 完成一份客户汇报。
使用 OfficeKit，更新 contract.docx 的日期和条款，保留目录、批注和页眉。
使用 OfficeKit，检查 report.pdf 的表单、签名和可访问性并输出审计结果。
```

## 不浪费上下文的交付复核

每个最终文件都会重新打开，并经过 OfficeKit 原生模型、文件结构、渲染和交付
哈希检查。能理解图片的 Agent 会直接复核页面或幻灯片；不能理解图片时，可以
按需调用随包分发的紧凑文本阅读视图（由 AnyDoc 提供解析），把标题、段落、表格
和跨格式内容整理成 Markdown，无需把所有截图都塞进上下文。

AnyDoc 解析器只在任务确实需要文本阅读视图时懒加载。它不能判断字体、裁剪、对比度、图表外观或页面
构图；无法直接查看而又涉及设计质量的结果，会明确标记为需要人工复核。

## Quick Start

安装一次 OfficeKit；不需要预装 Node.js 或 npm。

macOS（Apple 芯片）和 Linux x64：

```sh
curl -fsSL https://github.com/w31r4/OfficeKit/releases/latest/download/install.sh | sh
```

Windows PowerShell：

```powershell
irm https://github.com/w31r4/OfficeKit/releases/latest/download/install.ps1 | iex
```

新开一个终端，进入要处理 Office 文件的项目：

```sh
cd your-project
officekit init
```

`officekit init` 会识别项目里的 Agent 配置，并让你选择将 9 个 OfficeKit Skill
写入哪些目录。直接回车接受识别结果；需要明确指定时：

```sh
officekit init --tools claude,cursor
```

Claude Code 也可以通过仓库里的 marketplace 发现同一套 Skill：

```text
/plugin marketplace add w31r4/OfficeKit
/plugin install office-kit@officekit
```

如果只需要某个格式，也可以直接安装 `documents`、`spreadsheets`、
`presentations`、`pdf` 或 `template-creator`。对于 Claude Code 和其他
支持的 Agent，通用的项目初始化入口仍然是 `officekit init`。

也可以直接对正在使用的 Agent 说：

> 在这个项目安装并配置 OfficeKit。

它会使用同一安装器，运行初始化，并按项目配置选择目标。项目里有多个 Agent，
或目标无法判断时，才需要你确认。

安装新版后，在项目里刷新已安装的 Skill：

```sh
officekit update
```

Skill 里的 JavaScript 任务这样运行：

```sh
officekit run task.mjs -- input.docx output.docx
```

`officekit run` 使用已安装的同版本 API；任务自己的第三方依赖仍从任务所在项目解析。

需要连续完成检查、编辑、复核的任务，可以保持一个本地 JavaScript 会话：

```sh
officekit repl --workspace "$PWD" --task-root "$PWD/.officekit-task"
```

每行发送一个 JSON 请求，例如
`{"id":"inspect","code":"const {PdfFile}=await ctx.import('office-kit'); return await PdfFile.inspectPdf(ctx.inputRoot + '/input.pdf');"}`。
会话提供 `ctx.state`、`ctx.publish`、`ctx.recordEvidence` 和有类型的
`ctx.excel` facade。可以用
`officekit repl --resume /absolute/path/to/checkpoint.json` 恢复；恢复只加载
安全状态，不会重放可能产生副作用的代码。

## 直接操作当前打开的 Excel 工作簿

工作簿已经在 Microsoft Excel 里打开、甚至还没有保存时，走这条路径。OfficeKit 通过
自己的 Excel Add-in 把当前工作簿连接到本机 CLI：

```sh
officekit excel install
officekit excel doctor --json
```

`install` 会先征求是否信任用户级本地证书，然后输出 manifest 路径。第一次在 Excel
桌面版中按下面的路径上传：

```text
Home > Add-ins > My Add-ins > Upload My Add-in
```

在 Home Ribbon 打开 **OfficeKit** 并点击 **Connect OfficeKit**。随后 Agent 用
`officekit excel sessions --json` 找到目标工作簿，通过有类型的范围、格式、图表、表格、
PivotTable、截图和保存操作执行任务，并在完成前读回验证。

Excel Live Control V1 面向 Windows 和 macOS 的 Microsoft Excel 桌面版。第一次加载
Add-in 需要访问微软的 Office.js 运行时；工作簿内容和 OfficeKit 的请求审计都留在本机。

## 操作当前已经打开的 PowerPoint

如果演示文稿已经在桌面版 PowerPoint 中打开，使用 Live 路径，不要把未保存的
内容改走文件级 PPTX 流程：

```sh
officekit live install --app powerpoint --yes --json
officekit live doctor --app powerpoint --json
officekit live sessions --app powerpoint --json
officekit live execute request.json --json
```

安装命令会生成 manifest。第一次在 PowerPoint 中通过 **Home > Add-ins > My Add-ins
> Upload My Add-in** 上传它，再从 Home 功能区打开 OfficeKit 并连接目标演示文稿。
PowerPoint Live 提供有类型的幻灯片、选区、文本、形状、图片、单页预览和显式保存
操作；修改后会重新读取，遇到 `maybeApplied` 或 `unsupported-capability` 会明确报告，
不会偷偷改写已关闭的文件。第一轮真实宿主验收是 Windows x64 桌面版 PowerPoint；
macOS 当前只跑构建、mock 和打包检查。

## 用一句需求创建演示文稿

创建新 deck 时，Presentations 路径会先判断受众看完后应发生什么变化、这份材料
如何传递和会后使用、属于哪类演示场景，以及什么视觉方向最适合这个任务。随后再规划
叙事、完成静态构图、只在有意义时加入动效，并在交付前复核结果。用户提供的 Template
Skill、品牌规范或参考文件始终是设计权威；没有这些材料时，Agent 会为当前任务选择
独立的视觉方向。

自定义设计会先制作开场页、证据页和最高风险页进行校准，再扩展整份演示文稿。
Presentation Editorial Trim 会在构图前和首轮渲染后分别整理文案，保留事实、来源、
用户锁定措辞和局部编辑边界。

完整方法见 [OfficeKit 所说的演示文稿是什么](docs/what-is-a-presentation.zh-CN.md)，
其中说明了沟通任务、生命周期、六层质量模型和原生产物边界。

## 为演示文稿找图并保留来源

页面需要图片时，OfficeKit 可以搜索 Openverse、Wikimedia 或离线 Lucide 图标，
由 Agent 从候选中选择，再把图片字节和权利凭证保存到当前任务。用户素材和模板素材
仍然优先。

```sh
officekit image search "institutional bitcoin trading" \
  --task <task-id> --kind photo --purpose evidence \
  --orientation landscape --max 5 --json

officekit image add --task <task-id> --candidate <candidate-ref> --json
officekit image audit deck.pptx --task <task-id> \
  --sources-output deck.pptx.sources.json --json
```

搜索命令只返回候选，选择由 Agent 完成。登记后的图片按内容寻址，`tasks → resume`
后可以继续使用。审计命令会按 SHA-256 核对 PPTX 中实际嵌入的媒体，并列出可见署名
义务。Openverse 元数据标记为 provider-declared；Wikimedia 机器元数据与 Lucide ISC
包许可证分别保留自己的证据类型。

## 一个总入口，也保留直接入口

普通任务直接使用 OfficeKit。它会检查输入、确定输出格式、判断是否需要模板，
再把文件交给对应 Skill 完成。

| 入口 | 适合的任务 |
| --- | --- |
| [OfficeKit](skills/office-kit/skills/office-kit/SKILL.md) | 从目标直接开始，或处理跨格式、多交付物和模板判断。 |
| [Documents](skills/documents/skills/documents/SKILL.md) | 已确定要创建或修改 Word。 |
| [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md) | 已确定要处理 Excel、CSV、公式、模型或图表。 |
| [Excel Live Control](skills/spreadsheets/skills/excel-live-control/SKILL.md) | 通过本机 OfficeKit Add-in 操作 Microsoft Excel 桌面版里已经打开的工作簿。 |
| [Presentations](skills/presentations/skills/presentations/SKILL.md) | 已确定要创建或修改 PowerPoint。 |
| [Presentation Editorial Trim](skills/presentations/skills/presentation-editorial-trim/SKILL.md) | 只优化演示文稿文案，同时保留事实、来源、设计和局部修改范围。 |
| [PowerPoint Live Control](skills/presentations/skills/powerpoint-live-control/SKILL.md) | 操作桌面版 PowerPoint 中已经打开的演示文稿。 |
| [PDF](skills/pdf/skills/pdf/SKILL.md) | 已确定要读取、创建、检查或处理 PDF。 |
| [Template Creator](skills/template-creator/skills/template-creator/SKILL.md) | 把 DOCX 或 XLSX 保存为可复用的源文件模板。 |
| [Presentation Template Creator](skills/presentation-template-creator/skills/presentation-template-creator/SKILL.md) | 把演示参考资料提炼成可复用的风格指导和原创视觉示例。 |

入口不同，底层文件能力和检查规则相同。直接使用领域 Skill 会跳过格式路由，
并继续执行源文件保护、渲染和验证。

## 能做什么

| 文件 | 常见任务 |
| --- | --- |
| Word / DOCX | 报告、函件、合同草稿、样式、分节、页眉页脚、表格、图片、字段、批注和有界局部修改。 |
| Excel / XLSX | 数据整理、公式、样式、表格、验证、条件格式、图表、sparklines、有界 PivotTable 和财务模型。 |
| PowerPoint / PPTX | 演示文稿、模板套用、富文本、图片裁剪、表格、图表、连接线、备注、批注和 Master/Layout 保真。 |
| PDF | 创建、提取文本/表格/图片/链接、表单、批注、页面操作、渲染、rewrite 脱敏和有界签名。 |

完整边界见 [coverage](docs/coverage.md)。

## 模板按需使用

OfficeKit 用同一条搜索命令管理两类目录：
[Default Template Library](skills/default-template-library/README.md) 保存 13 套
MIT 授权的 DOCX/XLSX 源文件模板；
[Presentation Template Library](skills/presentation-template-library/README.md)
保存 39 套演示 Template Skill：包括 30 套 Kimi 风格方向、8 套 Codex 对齐风格，
以及 Evidence Ledger。每套包含风格指导、检索元数据和视觉校准图；只有明确声明并
绑定许可与哈希的 PPJ/PPTX 才会随包提供。`officekit init` 安装工作流 Skill，两个目录继续留在包内。

目标明确且模板未指定时，OfficeKit 把需求归一成英文检索词，再执行本地 BM25F
搜索。Agent 选择 0 或 1 个结果；选中演示模板后，读取风格指导、查看示例，形成当前
deck 的 Design Grammar，再自由构图。选择 `none` 也是正确结果。

```sh
officekit template search \
  --kind presentation \
  --purpose "quarterly business review" \
  --audience executive \
  --json
```

用户上传的 DOCX、XLSX 或 PPTX 默认只用于当前任务。DOCX/XLSX 由 Template Creator
保存；PPTX、图片、文字说明或已有 OfficeKit task 由 Presentation Template Creator
提炼，源参考文件不会进入发布模板。

## 文件交付前会再检查一遍

```text
读取原件 → 创建或修改 → 导出 → 重新打开 → 渲染页面 → 检查结果
```

- DOCX、XLSX 由公开 JavaScript 对象模型配合对应平台的 OfficeKit C#
  NativeAOT codec 读写。演示文稿使用唯一的严格 `.ppj` 程序；NativeAOT
  直接校验、投影并在 PPJ 与 PPTX 之间编译。
- OfficeKit 先确定复杂 Office 内容的可编辑范围，再修改受支持的部分；其余内容
  保持原样并报告具体限制。
- PDF 默认通过 MuPDF.js 读取、编辑、检查和渲染。qpdf、OCR、严格清理、
  pyHanko 签名和 veraPDF 等重型能力由项目显式授权后按任务加载。

PDF provider 的策略和限制见
[Provider Setup](skills/pdf/skills/pdf/tasks/provider_setup.md)。

## JavaScript API

Skills 和应用代码使用同一个 API。Skill 任务通过 `officekit run` 使用全局包；
应用开发者也可以把 `office-kit` 加入自己的项目依赖后直接调用：

```js
import { SpreadsheetFile, Workbook } from "office-kit";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Summary");
sheet.getRange("A1:B2").values = [
  ["Metric", "Value"],
  ["Revenue", 42.5],
];

const file = await SpreadsheetFile.exportXlsx(workbook, { recalculate: true });
await file.save("summary.xlsx");
```

演示文稿创建和第三方 PPTX 续写统一使用 PPJ，不再暴露 JavaScript
Presentation 对象模型：

```sh
officekit ppj import input.pptx -o deck.ppj --json
officekit ppj check deck.ppj --json
officekit ppj build deck.ppj -o deck.pptx --json
officekit ppj render deck.ppj -o previews/ --json
officekit ppj review deck.ppj --json
```

Agent 直接修改严格 JSON。第三方文件无修改 build 会逐字节返回源包；无法证明
安全的 native mutation 会明确拒绝。设计理由见
[OfficeKit 为什么使用 PPJ](docs/why-ppj.zh-CN.md)。

可运行示例：

- [创建 DOCX 报告](examples/create-docx-report.mjs)
- [创建 XLSX 仪表盘](examples/create-xlsx-dashboard.mjs)
- [Evidence Ledger PPJ](skills/presentation-template-library/skills/artifact-template-evidence-ledger/assets/references/reference.ppj)
- [解析与渲染 PDF](examples/parse-render-pdf.mjs)

需要直接访问底层 Office codec 时，使用 `office-kit/codec`；生成的 wire
类型位于 `office-kit/codec/wire`。

## 文档与开发

- [API 参考](https://github.com/w31r4/OfficeKit/blob/main/docs/api.md)
- [参考 Skill 兼容性](https://github.com/w31r4/OfficeKit/blob/main/docs/reference-skills.md)
- [全部能力边界](https://github.com/w31r4/OfficeKit/blob/main/docs/coverage.md)
- [发布状态](https://github.com/w31r4/OfficeKit/blob/main/docs/release.md)

```sh
npm test
npm run test:pack
npm run docs:api
npm run release:check
```

`OfficeKit` 是产品名。独立安装包发布在
[GitHub Releases](https://github.com/w31r4/OfficeKit/releases)；JavaScript API
继续面向需要将 OfficeKit 嵌入应用的开发者。

## 许可证

[GNU AGPL v3 或更高版本](LICENSE)。网络部署、修改和分发必须遵守 AGPL
的对应义务。第三方运行时、MuPDF 和专项 provider 的许可证与来源见
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。
