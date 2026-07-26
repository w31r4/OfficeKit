# OfficeKit

## 把一句需求变成可以交付的 Office 文件

**简体中文** | [English](README.en.md)

给 Agent 数据、现有文件和一句需求，拿回能打开、能继续编辑、能交付的
Word、Excel、PowerPoint 或 PDF。

```text
→ 一个入口，自动选择 Word、Excel、PPT 或 PDF 工作流
→ 有合适模板就复用，没有就从零开始
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

需要 Node.js 22 或更新版本。在 Agent 工作的项目中执行：

```sh
npm install github:w31r4/open-office-artifact-tool
npx skills add w31r4/open-office-artifact-tool --skill '*' --yes
```

第一条命令安装文件运行库，第二条命令安装 OfficeKit、四个文件类型 Skill、
Excel Live Control、Template Creator 和开源模板。无需 clone 仓库，也不要求
本机安装 Microsoft Office、.NET SDK 或 Python。

只安装核心 Skills：

```sh
npx skills add w31r4/open-office-artifact-tool \
  --skill officekit documents spreadsheets excel-live-control presentations pdf template-creator \
  --yes
```

当前正式 npm 包尚未发布，因此使用 GitHub 安装源。发布后第一条命令可改为：

```sh
npm install open-office-artifact-tool
```

## 一个总入口，也保留直接入口

普通任务直接使用 OfficeKit。它会检查输入、确定输出格式、判断是否需要模板，
再把文件交给对应 Skill 完成。

| 入口 | 适合的任务 |
| --- | --- |
| [OfficeKit](skills/officekit/skills/officekit/SKILL.md) | 不想先研究工具，或需要跨格式、多交付物和模板判断。 |
| [Documents](skills/documents/skills/documents/SKILL.md) | 已确定要创建或修改 Word。 |
| [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md) | 已确定要处理 Excel、CSV、公式、模型或图表。 |
| [Excel Live Control](skills/spreadsheets/skills/excel-live-control/SKILL.md) | 操作已经在 Excel 中打开的工作簿。 |
| [Presentations](skills/presentations/skills/presentations/SKILL.md) | 已确定要创建或修改 PowerPoint。 |
| [PDF](skills/pdf/skills/pdf/SKILL.md) | 已确定要读取、创建、检查或处理 PDF。 |
| [Template Creator](skills/template-creator/skills/template-creator/SKILL.md) | 把自己的 DOCX、XLSX 或 PPTX 保存为可复用模板。 |

入口不同，底层文件能力和检查规则相同。直接使用领域 Skill 只会跳过格式路由，
不会绕过源文件保护、渲染或验证。

## 能做什么

| 文件 | 常见任务 |
| --- | --- |
| Word / DOCX | 报告、函件、合同草稿、样式、分节、页眉页脚、表格、图片、字段、批注和有界局部修改。 |
| Excel / XLSX | 数据整理、公式、样式、表格、验证、条件格式、图表、sparklines、有界 PivotTable 和财务模型。 |
| PowerPoint / PPTX | 演示文稿、模板套用、富文本、图片裁剪、表格、图表、连接线、备注、批注和 Master/Layout 保真。 |
| PDF | 创建、提取文本/表格/图片/链接、表单、批注、页面操作、渲染、rewrite 脱敏和有界签名。 |

完整边界见 [coverage](docs/coverage.md)。

## 模板是加速器，不是前提

[Office Template Library](skills/default-template-library/README.md) 提供 20 套
MIT 授权模板。OfficeKit 只在目标明确、模板未指定时检索紧凑 metadata；
Agent 查看少量候选后选择一个、询问用户或明确不用模板。

用户上传的 DOCX、XLSX 或 PPTX 可以只用于当前任务，不会自动注册或覆盖。
只有明确要求以后复用时，才交给 Template Creator 保存。

## 文件交付前会再检查一遍

```text
读取原件 → 创建或修改 → 导出 → 重新打开 → 渲染页面 → 检查结果
```

- DOCX、XLSX 和 PPTX 统一通过 OpenChestnut C#/.NET WASM 读写；包内没有第二套
  JavaScript Office writer 或静默 fallback。
- 无法安全修改的 Office 内容会原样保留或明确拒绝修改，不会为了“看起来成功”
  而破坏原文件。
- PDF 默认通过 MuPDF.js 读取、编辑、检查和渲染。qpdf、OCR、严格清理、
  pyHanko 签名和 veraPDF 等重型能力由项目显式授权后按需启用，不会静默下载。

PDF provider 的策略和限制见
[Provider Setup](skills/pdf/skills/pdf/tasks/provider_setup.md)。

## JavaScript API

Skills 和应用代码使用同一个包。Agent 可以按 Skill 完成任务，应用也可以直接调用
API：

```js
import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

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

旧的 `codecs/openxml-wasm` 导入在 0.x 中仍是 OpenChestnut 的弃用别名；
新代码应使用 `codecs/open-chestnut`。详见 [release notes](docs/release.md)。

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

`OfficeKit` 是产品名；当前包名仍为 `open-office-artifact-tool`。版本 `0.3.0`
是 release candidate，尚未正式发布到 npm。

## 许可证

[GNU AGPL v3 或更高版本](LICENSE)。网络部署、修改和分发必须遵守 AGPL
的对应义务。第三方运行时、MuPDF 和专项 provider 的许可证与来源见
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。
