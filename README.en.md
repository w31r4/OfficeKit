# OfficeKit

## Turn one request into an Office file ready to hand off

[简体中文](README.md) | **English**

Give an agent your data, existing files, and a short request. Get back a Word,
Excel, PowerPoint, or PDF file that opens, remains editable, and is ready to use.

```text
→ one entry point for Word, Excel, PowerPoint, and PDF work
→ reuse a template when it fits; start clean when it does not
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

Requires Node.js 22 or newer. Run these commands in the project where the agent
will work:

```sh
npm install github:w31r4/OfficeKit
npx skills add w31r4/OfficeKit --skill '*' --yes
```

The first command installs the file runtime. The second installs OfficeKit, the
four file-type Skills, Excel Live Control, Template Creator, and the open-source
templates. No repository clone, Microsoft Office installation, .NET SDK, or
Python setup is required.

To install only the core Skills:

```sh
npx skills add w31r4/OfficeKit \
  --skill office-kit documents spreadsheets excel-live-control presentations pdf template-creator \
  --yes
```

The formal npm package has not shipped yet, so the runtime currently uses the
GitHub source. After publication, the first command becomes:

```sh
npm install office-kit
```

## One front door, with direct routes when you want them

For ordinary work, start with OfficeKit. It inspects the inputs, decides the
output route, considers templates, and hands each file to its owning Skill.

| Entry | Good for |
| --- | --- |
| [OfficeKit](skills/office-kit/skills/office-kit/SKILL.md) | Starting from the requested outcome without first choosing tools, or handling cross-format and multi-deliverable work. |
| [Documents](skills/documents/skills/documents/SKILL.md) | Creating or changing a Word document when the format is already known. |
| [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md) | Excel, CSV, formulas, models, data preparation, and charts. |
| [Excel Live Control](skills/spreadsheets/skills/excel-live-control/SKILL.md) | Working with a workbook already open in Excel. |
| [Presentations](skills/presentations/skills/presentations/SKILL.md) | Creating or changing a PowerPoint presentation. |
| [PDF](skills/pdf/skills/pdf/SKILL.md) | Reading, creating, inspecting, or processing a PDF. |
| [Template Creator](skills/template-creator/skills/template-creator/SKILL.md) | Saving your own DOCX, XLSX, or PPTX reference as a reusable template. |

Every route uses the same file capabilities and checks. Invoking a domain Skill
directly skips format routing, not source protection, rendering, or verification.

## What it handles

| File | Common work |
| --- | --- |
| Word / DOCX | Reports, letters, contract drafts, styles, sections, headers and footers, tables, images, fields, comments, and bounded edits. |
| Excel / XLSX | Data preparation, formulas, styles, tables, validation, conditional formats, charts, sparklines, bounded PivotTables, and financial models. |
| PowerPoint / PPTX | Decks, templates, rich text, reversible image crops, tables, charts, connectors, notes, comments, and master/layout fidelity. |
| PDF | Authoring; text, table, image, and link extraction; forms; annotations; page operations; rendering; rewrite redaction; and bounded signing. |

See [coverage](docs/coverage.md) for the complete supported boundary.

## Templates accelerate the work; they are not a requirement

The [Office Template Library](skills/default-template-library/README.md) provides
20 MIT-licensed templates. OfficeKit searches compact metadata only when the
goal is clear and no template has been specified. After reviewing a small
shortlist, the agent selects one, asks, or proceeds without a template.

An uploaded DOCX, XLSX, or PPTX can be used for one task without registration
or overwrite. Template Creator is used only when the user explicitly wants to
save that reference for reuse.

## Files are checked before handoff

```text
read the source → create or change → export → reopen → render pages → check the result
```

- DOCX, XLSX, and PPTX use OfficeKit C#/.NET WASM. The package has no second
  JavaScript Office writer or silent fallback.
- Office content that cannot be changed safely is preserved or rejected
  explicitly instead of being damaged to produce an apparently successful file.
- PDF uses MuPDF.js for normal reading, editing, inspection, and rendering.
  qpdf, OCR, strict cleanup, pyHanko signing, veraPDF, and other heavy
  capabilities are enabled on demand only after project authorization.

See [Provider Setup](skills/pdf/skills/pdf/tasks/provider_setup.md) for PDF
provider policy and operational limits.

## JavaScript API

Skills and application code use the same package. An agent can follow a Skill,
while an application can call the API directly:

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

`OfficeKit` is the product name, and the npm package is `office-kit`.
Version `0.3.0` is a release candidate and has not
been formally published to npm.

## License

[GNU AGPL v3 or later](LICENSE). Network deployment, modification, and
redistribution must meet the applicable AGPL obligations. Third-party runtime,
MuPDF, and specialist-provider licenses and provenance are recorded in
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
