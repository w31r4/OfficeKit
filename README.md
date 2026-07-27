# OfficeKit

## Turn one request into an Office file ready to hand off

**English** | [简体中文](README.zh-CN.md)

Give an agent your data, existing files, and a short request. Get back a Word,
Excel, PowerPoint, or PDF file that opens, remains editable, and is ready to use.

```text
→ one entry point for Word, Excel, PowerPoint, and PDF work
→ reuse a fitting template or design from scratch
→ preserve complex source content that was not changed
→ reopen, render, and check every deliverable
```

## See it in action

```text
You:
Use OfficeKit. Turn sales.xlsx into a Q2 business review deck for management.
Include revenue, margin, regional differences, major risks, and three decisions.
Choose a template only if one clearly fits.

OfficeKit:
read the data → choose the presentation workflow → select a template or none
→ build the deck → reopen it → render and check it → return q2-review.pptx
```

Single-format requests are just as direct:

```text
Use OfficeKit to turn these CSV files into an Excel operating dashboard with formulas, charts, and exception flags.
Use OfficeKit with template.pptx and data.xlsx to complete a customer presentation.
Use OfficeKit to update the dates and clauses in contract.docx without disturbing its TOC, comments, or headers.
Use OfficeKit to inspect the forms, signatures, and accessibility of report.pdf and return an audit.
```

## Quick Start

Install OfficeKit once. Node.js and npm are not prerequisites.

macOS on Apple silicon and Linux x64:

```sh
curl -fsSL https://github.com/w31r4/OfficeKit/releases/latest/download/install.sh | sh
```

Windows PowerShell:

```powershell
irm https://github.com/w31r4/OfficeKit/releases/latest/download/install.ps1 | iex
```

Open a new terminal, then enter the project where you want to work with Office files:

```sh
cd your-project
officekit init
```

`officekit init` finds the Agent configurations in the project and lets you
choose which directories receive the seven OfficeKit Skills. Press Enter to
accept the detected targets, or name them explicitly:

```sh
officekit init --tools claude,cursor
```

You can also tell your current Agent:

> Set up OfficeKit in this project.

It uses the same installer, runs initialization, and selects targets from the
project configuration. It asks you to confirm only when several Agent targets
are present or none can be identified.

After installing a newer release, refresh the installed Skills from the project:

```sh
officekit update
```

Run JavaScript tasks referenced by the Skills with:

```sh
officekit run task.mjs -- input.docx output.docx
```

`officekit run` supplies the matching installed OfficeKit API. A task's other
dependencies continue to resolve from the task's own project.

## One front door, with direct routes when you want them

For ordinary work, start with OfficeKit. It inspects the inputs, decides the
output route, considers templates, and hands each file to its owning Skill.

| Entry | Good for |
| --- | --- |
| [OfficeKit](skills/office-kit/skills/office-kit/SKILL.md) | Starting from the requested outcome, or handling cross-format and multi-deliverable work. |
| [Documents](skills/documents/skills/documents/SKILL.md) | Creating or changing a Word document when the format is already known. |
| [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md) | Excel, CSV, formulas, models, data preparation, and charts. |
| [Excel Live Control](skills/spreadsheets/skills/excel-live-control/SKILL.md) | Working with a workbook already open in Excel. |
| [Presentations](skills/presentations/skills/presentations/SKILL.md) | Creating or changing a PowerPoint presentation. |
| [PDF](skills/pdf/skills/pdf/SKILL.md) | Reading, creating, inspecting, or processing a PDF. |
| [Template Creator](skills/template-creator/skills/template-creator/SKILL.md) | Saving your own DOCX, XLSX, or PPTX reference as a reusable template. |

Every route uses the same file capabilities and checks. Invoking a domain Skill
directly skips format routing while retaining source protection, rendering, and
verification.

## What it handles

| File | Common work |
| --- | --- |
| Word / DOCX | Reports, letters, contract drafts, styles, sections, headers and footers, tables, images, fields, comments, and bounded edits. |
| Excel / XLSX | Data preparation, formulas, styles, tables, validation, conditional formats, charts, sparklines, bounded PivotTables, and financial models. |
| PowerPoint / PPTX | Decks, templates, rich text, reversible image crops, tables, charts, connectors, notes, comments, and master/layout fidelity. |
| PDF | Authoring; text, table, image, and link extraction; forms; annotations; page operations; rendering; rewrite redaction; and bounded signing. |

See [coverage](docs/coverage.md) for the complete supported boundary.

## Use templates when they fit

The [Office Template Library](skills/default-template-library/README.md) provides
20 MIT-licensed templates stored once in the installed OfficeKit runtime. `officekit init`
installs Skills and leaves the template assets in place. When the goal is clear
and no template has been specified, OfficeKit normalizes the intent into
English search terms and runs local BM25F retrieval. After reviewing a small
shortlist, the agent selects one, asks, or proceeds without a template.

```sh
officekit template search \
  --kind presentation \
  --purpose "quarterly business review" \
  --audience executive \
  --json
```

An uploaded DOCX, XLSX, or PPTX stays scoped to the current task by default.
Template Creator saves it when the user explicitly wants future reuse.

## Files are checked before handoff

```text
read the source → create or change → export → reopen → render pages → check the result
```

- DOCX, XLSX, and PPTX use OfficeKit C#/.NET WASM. Import, editing, export, and
  second-pass verification follow the same path.
- OfficeKit identifies the editable scope of complex Office content, applies
  supported changes, preserves the rest, and reports the exact boundary.
- PDF uses MuPDF.js for normal reading, editing, inspection, and rendering.
  qpdf, OCR, strict cleanup, pyHanko signing, veraPDF, and other heavy
  capabilities load by task after project authorization.

See [Provider Setup](skills/pdf/skills/pdf/tasks/provider_setup.md) for PDF
provider policy and operational limits.

## JavaScript API

Skills and application code use the same API. Skill tasks use the global
package through `officekit run`; application developers can also add
`office-kit` to a project and call it directly:

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

Runnable examples:

- [Create a DOCX report](examples/create-docx-report.mjs)
- [Create an XLSX dashboard](examples/create-xlsx-dashboard.mjs)
- [Create a PPTX deck with Compose](examples/create-pptx-compose.mjs)
- [Parse and render a PDF](examples/parse-render-pdf.mjs)

For direct access to the low-level Office codec, use `office-kit/codec`.
Generated wire types are available from `office-kit/codec/wire`.

## Documentation and development

- [API reference](docs/api.md)
- [Reference Skill compatibility](docs/reference-skills.md)
- [Complete capability boundary](docs/coverage.md)
- [Release status](docs/release.md)

```sh
npm test
npm run test:pack
npm run docs:api
npm run release:check
```

OfficeKit's standalone installers are published through
[GitHub Releases](https://github.com/w31r4/OfficeKit/releases). The JavaScript
API remains available for developers who embed OfficeKit in an application.

## License

[GNU AGPL v3 or later](LICENSE). Network deployment, modification, and
redistribution must meet the applicable AGPL obligations. Third-party runtime,
MuPDF, and specialist-provider licenses and provenance are recorded in
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
