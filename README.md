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

## Review without wasting the context window

Every final file is reopened and checked through its native OfficeKit model,
package structure, render pipeline, and delivery hash. An Agent that can inspect
images reviews the rendered pages or slides directly. An Agent without image
understanding can request the bundled compact text reading view, powered by
AnyDoc: Markdown for checking headings, paragraphs, tables, and cross-format
content without loading every screenshot into context.

The AnyDoc parser is loaded only when the task benefits from the text reading
view. It does not judge
typography, crop, contrast, chart appearance, or composition, so design-sensitive
work that cannot be viewed is clearly marked for human review.

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
choose which directories receive the nine OfficeKit Skills. Press Enter to
accept the detected targets, or name them explicitly:

```sh
officekit init --tools claude,cursor
```

Claude Code can also discover the same canonical Skill trees through the
repository marketplace:

```text
/plugin marketplace add w31r4/OfficeKit
/plugin install office-kit@officekit
```

Install the direct `documents`, `spreadsheets`, `presentations`, `pdf`, or
`template-creator` plugin only when you want that route without the OfficeKit
coordinator. `officekit init` remains the portable project initializer for
Claude Code and other supported Agent tools.

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

For a task with several inspect/edit/verify steps, keep one local JavaScript
session instead of rebuilding the object graph for every command:

```sh
officekit repl --workspace "$PWD" --task-root "$PWD/.officekit-task"
```

Send one JSON request per line, for example
`{"id":"inspect","code":"const {PdfFile}=await ctx.import('office-kit'); return await PdfFile.inspectPdf(ctx.inputRoot + '/input.pdf');"}`.
The session exposes `ctx.state`, `ctx.publish`, `ctx.recordEvidence`, and a
typed `ctx.excel` facade. Checkpoints can be resumed with
`officekit repl --resume /absolute/path/to/checkpoint.json`; resume restores
safe state without replaying side effects.

## Work in the Excel workbook already on screen

Use this route when the workbook is open in Microsoft Excel and may contain
unsaved work. OfficeKit connects the workbook to the local CLI through its own
Excel Add-in:

```sh
officekit excel install
officekit excel doctor --json
```

`install` asks before adding a user-local certificate trust, then prints the
manifest path. In Excel desktop, upload it once through:

```text
Home > Add-ins > My Add-ins > Upload My Add-in
```

Open **OfficeKit** from the Home ribbon and click **Connect OfficeKit**. The
Agent then discovers the intended workbook with `officekit excel sessions
--json`, uses typed range, formatting, chart, table, PivotTable, screenshot,
and save operations, and reads changes back before reporting completion.

Excel Live Control V1 targets Microsoft Excel desktop on Windows and macOS.
The first Add-in load uses Microsoft's Office.js runtime; the workbook and
OfficeKit's request audit remain on the local machine.

## Work in the PowerPoint deck already on screen

For a presentation that is open in desktop PowerPoint, use the companion live
route rather than the file-level PPTX workflow:

```sh
officekit live install --app powerpoint --yes --json
officekit live doctor --app powerpoint --json
officekit live sessions --app powerpoint --json
officekit live execute request.json --json
```

Upload the printed manifest once through **Home > Add-ins > My Add-ins > Upload
My Add-in**, open OfficeKit from the Home ribbon, and connect the intended deck.
PowerPoint Live uses typed slide, selection, text, shape, image, slide-preview,
and explicit-save operations. It can work with unsaved changes, rereads after
mutations, and reports `maybeApplied` or `unsupported-capability` instead of
editing a closed file behind the user's back. The first real host acceptance
matrix is Windows x64 desktop PowerPoint; macOS currently has build, mock, and
package checks only.

## Create a presentation from one request

For a new deck, the Presentations route turns a clear request into a working
draft by deciding what must change for the audience, how the deck will be used,
which presentation scenario applies, and what visual direction serves that
job. It then plans the story, composes the pages, adds motion only when useful,
and checks the result before delivery. A supplied template, brand guide, or
reference deck remains the design authority; without one, the Agent chooses a
task-specific direction. Grid Layout is an explicit scaffold, not a hidden
fallback.

Read [What OfficeKit Means by a Presentation](docs/what-is-a-presentation.md)
for the communication, lifecycle, quality, and native-artifact principles that
drive this workflow.

The plan is stored with the task, so a later session can reopen the reviewed
revision, see the pending decision, make a local change, and review it again.
For self-directed decks, OfficeKit calibrates an opening page, an evidence page,
and the densest page before expanding the full story. Route C is the default;
Grid remains an explicit scaffold when the user or plan requests it.

## One front door, with direct routes when you want them

For ordinary work, start with OfficeKit. It inspects the inputs, decides the
output route, considers templates, and hands each file to its owning Skill.

| Entry | Good for |
| --- | --- |
| [OfficeKit](skills/office-kit/skills/office-kit/SKILL.md) | Starting from the requested outcome, or handling cross-format and multi-deliverable work. |
| [Documents](skills/documents/skills/documents/SKILL.md) | Creating or changing a Word document when the format is already known. |
| [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md) | Excel, CSV, formulas, models, data preparation, and charts. |
| [Excel Live Control](skills/spreadsheets/skills/excel-live-control/SKILL.md) | Working with a workbook already open in Microsoft Excel desktop through the local OfficeKit Add-in. |
| [Presentations](skills/presentations/skills/presentations/SKILL.md) | Creating or changing a PowerPoint presentation. |
| [Presentation Editorial Trim](skills/presentations/skills/presentation-editorial-trim/SKILL.md) | Polishing slide copy while preserving facts, sources, design, and local edit scope. |
| [PowerPoint Live Control](skills/presentations/skills/powerpoint-live-control/SKILL.md) | Working with a presentation already open in desktop PowerPoint through the local OfficeKit bridge. |
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

- [API reference](https://github.com/w31r4/OfficeKit/blob/main/docs/api.md)
- [Reference Skill compatibility](https://github.com/w31r4/OfficeKit/blob/main/docs/reference-skills.md)
- [Complete capability boundary](https://github.com/w31r4/OfficeKit/blob/main/docs/coverage.md)
- [Release status](https://github.com/w31r4/OfficeKit/blob/main/docs/release.md)

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
