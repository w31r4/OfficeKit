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

`officekit init` 会识别项目里的 Agent 配置，并让你选择将 7 个 OfficeKit Skill
写入哪些目录。直接回车接受识别结果；需要明确指定时：

```sh
officekit init --tools claude,cursor
```

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
| [PDF](skills/pdf/skills/pdf/SKILL.md) | 已确定要读取、创建、检查或处理 PDF。 |
| [Template Creator](skills/template-creator/skills/template-creator/SKILL.md) | 把自己的 DOCX、XLSX 或 PPTX 保存为可复用模板。 |

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

[Office Template Library](skills/default-template-library/README.md) 提供 20 套
MIT 授权模板，随已安装的 OfficeKit 保存一份。`officekit init` 只安装 Skill，模板继续
留在包内。目标明确且模板未指定时，OfficeKit 把需求归一成英文检索词，
再执行本地 BM25F 搜索；Agent 查看少量候选后选择一个、询问用户或明确
不用模板。

```sh
officekit template search \
  --kind presentation \
  --purpose "quarterly business review" \
  --audience executive \
  --json
```

用户上传的 DOCX、XLSX 或 PPTX 默认只用于当前任务。明确要求以后复用时，
再交给 Template Creator 保存。

## 文件交付前会再检查一遍

```text
读取原件 → 创建或修改 → 导出 → 重新打开 → 渲染页面 → 检查结果
```

- DOCX、XLSX 和 PPTX 统一通过 OfficeKit C#/.NET WASM 读写；导入、编辑、
  导出和二次校验沿用同一条路径。
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

可运行示例：

- [创建 DOCX 报告](examples/create-docx-report.mjs)
- [创建 XLSX 仪表盘](examples/create-xlsx-dashboard.mjs)
- [使用 Compose 创建 PPTX](examples/create-pptx-compose.mjs)
- [解析与渲染 PDF](examples/parse-render-pdf.mjs)

需要直接访问底层 Office codec 时，使用 `office-kit/codec`；生成的 wire
类型位于 `office-kit/codec/wire`。

## 文档与开发

- [API 参考](docs/api.md)
- [参考 Skill 兼容性](docs/reference-skills.md)
- [全部能力边界](docs/coverage.md)
- [发布状态](docs/release.md)

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
