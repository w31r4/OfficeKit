---
name: "Spreadsheets"
description: "Create, edit, analyze, and verify standalone spreadsheet files or Google Sheets-ready workbooks, including .xlsx, .xls, .csv, and .tsv. Do not use for live controlling Microsoft Excel app or a live Excel session."
---

# Spreadsheets skill (Create • Edit • Analyze • Visualize)

Use the shared `../office-kit/references/workspace.md` contract for paths and
results. Standalone workbook work uses `workspaceRoot`, `inputRoot`,
`assetRoot`, `outputRoot`, and `evidenceRoot`; it does not require a host
workspace loader. After the final workbook is reopened, follow
`../office-kit/references/review.md`; use the optional text reading view
(`contentView: "anydoc"`) only to close an identified sheet-value or table
coverage gap, never as proof of formulas, chart appearance, cell formatting,
or layout.

## Run the workbook workflow in one task

For a multi-step workbook task, use `officekit repl` and the portable guidance
in `../office-kit/references/repl.md`. Import the public API with
`await ctx.import("office-kit")`, keep the workbook and reusable calculations
under `ctx.state`, and use inspect → edit → recalculate → reimport/render →
verify before `ctx.publish`. Register render or inspect files with
`ctx.recordEvidence`; return the final absolute path, SHA-256, and
`visualReview` status. A currently open workbook belongs to Excel Live Control,
not this file workflow.
Use this skill when you need to work with spreadsheets (.xlsx, .csv, .tsv) to do any of the following:
- Create or modify a new workbook/sheet with proper formulas, cell/number formatting, and structured layout
- Read or analyze tabular data (filter, aggregate, pivot, compute metrics) directly in a sheet
- Visualize data with in-sheet charts/tables and sensible formatting
- Recalculate/evaluate formulas to update results after changes

## Decision Boundary

- For Google Sheets-targeted outputs, such as creating or editing a Google Sheet, follow the additional instructions here: `routing/google_sheets.md`.

Do not follow those routing instructions if irrelevant to the task. Default is to create/edit spreadsheets with artifact tool.

## Tools + Contract Requirements
- Use the public `office-kit` JS library for all spreadsheet authoring. Run
  bundled examples with `officekit run` or the active OfficeKit installation;
  do not assume a host-specific dependency loader.
- Never import or use `@oai/artifact-tool`. It is a different host-bundled
  runtime, not an OfficeKit alias or fallback, and its output must never be
  attributed to OfficeKit.
- If the runtime or `office-kit` is unavailable, report a blocker. Do not guess or search for paths, install packages, use resolution hacks, or import bundled internals.
- Work in a writable `taskRoot` or system temporary directory. Keep generated
  sources and intermediate files out of the managed installation directory.
- Prefer one executable `.mjs` builder and patch/rerun it. Do not use heredocs or duplicate builders.
- Use the provided API reference. Do not inspect package internals or prototypes. If blocked, run at most one targeted `workbook.help("<api_or_feature>")` query.
- Do not use alternate workbook creation/editing libraries such as `openpyxl`, `xlsxwriter`, or `pandas.ExcelWriter` unless the user explicitly asks.
- For supporting analysis or data processing outside workbook authoring, use JS or spreadsheet formulas when sufficient. If Python is necessary, prefer the bundled python libraries, save JSON/CSV intermediates, and have the JS builder create the workbook. Use existing system Python or user-provided libraries only when the bundled environment lacks a required capability. Keep auditable and user-editable calculations in the workbook as formulas.
- Use `update_plan` for complex spreadsheet work.

### Final Response
- Include a short user-visible summary and standalone Markdown link(s) only to final `.xlsx` artifact(s), one per line: `[Revenue Model - MNST.xlsx](/absolute/path/to/revenue_model_mnst.xlsx)`.
- Do not mention or link builders, previews, or other support files unless requested.

Other documents:
- `style_guidelines.md`: REQUIRED for formatting requirements
- `artifact_tool_docs/API_QUICK_START.md`: REQUIRED API documentation for `artifact_tool` JS library, which exposes methods to read, manipulate, edit, recalculate, render, import and save spreadsheets. You must read it entirely to get started.
- `features/charts.md`: Read when creating or editing charts.
- `features/pivot-tables.md`: Read when creating, importing, or verifying native PivotTables.
- `examples/officekit-range-workflow.mjs`: Run or adapt this public-package example for R1C1 formulas, block writes, range navigation, formula inspection, bounded native line-chart trendlines and percentage error bars, OfficeKit roundtrip, native XML, and SVG verification.
- `examples/officekit-sparkline-workflow.mjs`: Run or adapt this canonical example for line/column sparkline authoring, Range aliases, inspect/render/verify, and source-bound OfficeKit edits.
- `examples/officekit-data-table-workflow.mjs`: Run or adapt this canonical example for one-variable and two-variable What-If data-table authoring, inspect, OfficeKit roundtrip, and source-bound imported topology.
- `examples/officekit-data-validation-workflow.mjs`: Run or adapt this canonical example for list/whole/custom validation authoring, input prompts, error alerts, blank policy, intuitive drop-down visibility, OfficeKit roundtrip, fixed-range edits, and native render QA.
- `examples/officekit-worksheet-protection-workflow.mjs`: Run or adapt this canonical example for unlocked input cells, hidden locked formulas, passwordless worksheet editing restrictions, source-bound permission edits, second import, and native render QA. Protection is not encryption or access control.
- `examples/officekit-pivot-table-workflow.mjs`: Run or adapt this canonical example for bounded native PivotTable authoring with exact item and absolute whole-day date filters, source and cached-value audit, inspect/render/verify, OfficeKit roundtrip, and the one source-bound `refreshOnLoad=true` → `false` hardening operation on a uniquely owned imported cache. Its native-host check records that LibreOffice resave drops the advanced date filter and recalculates that Pivot unfiltered; it does not claim the filter survived. Other imported Pivot topology remains read-only.
- `examples/officekit-pivot-refresh-hardening-workflow.mjs`: Run or adapt this no-overwrite imported-workbook transaction only when one named PivotTable on one named worksheet proves the uniquely owned explicit `refreshOnLoad=true` capability. It changes only its cache-definition root to false, compares every other package part byte-for-byte, reimports/renders, and writes a byte-bound audit; it never refreshes data or becomes a general Pivot editor.
- `examples/officekit-financial-returns-workflow.mjs`: Run or adapt this three-sheet `Inputs` / `Returns` / `Checks` example for visible finance/reinvestment assumptions, finance color conventions, NPV/XNPV/IRR/XIRR/MIRR/PMT outputs, guarded return checks, SVG review, and canonical OfficeKit roundtrip.
- `examples/officekit-loan-amortization-workflow.mjs`: Run or adapt this three-sheet `Inputs` / `Amortization` / `Checks` example for a visible loan schedule using PMT/IPMT/PPMT plus RATE/PV/FV/NPER inverse audits and CUMIPMT/CUMPRINC schedule reconciliations, payment-timing-aware first-period checks, finance color conventions, SVG review, and canonical OfficeKit roundtrip.
- `examples/officekit-asset-depreciation-workflow.mjs`: Run or adapt this three-sheet `Inputs` / `Depreciation` / `Checks` example for visible fixed-asset assumptions, SLN/DB/DDB schedules, salvage-floor and depreciable-basis checks, finance color conventions, SVG review, and canonical OfficeKit roundtrip.
- `examples/officekit-scatter-chart-workflow.mjs`: Run or adapt this canonical example for formula-backed numeric X/Y scatter series, dual value axes, marker styling, SVG verification, and fixed-topology OfficeKit edits.
- `examples/officekit-bubble-chart-workflow.mjs`: Run or adapt this canonical example for one formula-backed numeric X/Y/Size bubble series, dual value axes, proportional SVG/native bubbles, and fixed-topology OfficeKit edits.
- `examples/officekit-threaded-comment-reply-workflow.mjs`: Run or adapt this imported-workbook example when an existing canonical Excel threaded-comment root needs one or more direct replies plus an explicit resolve/reopen transition; it preserves existing IDs, people, dates, text, and direct-reply order, performs a second import and all-sheet SVG review, and writes a byte-bound audit. Nested/branched reply graphs remain source-bound and fail closed.
- `examples/officekit-growth-assumption-edit-workflow.mjs`: Run or adapt this bounded imported-workbook example when exactly one named growth assumption may change. It protects a second-sheet canary, formula topology, a separate margin assumption, sheet identity/order, source bytes, and writes a byte-bound OfficeKit rewrite audit after recalculation, second import, and all-sheet SVG review.
- `examples/officekit-connection-refresh-hardening-workflow.mjs`: Run or adapt this source-bound transaction only to turn one imported connection's explicit `refreshOnLoad=true` request off. It refuses an absent/already-false/ambiguous connection, preserves every other modeled connection field, protects the input, publishes without overwrite, reimports, renders all sheets, and writes a byte-bound audit. It does not run external data or stop manual, macro, PivotTable, or other host-triggered refreshes.
- `examples/officekit-operating-plan-workflow.mjs`: Run or adapt this forward workflow for a locked 24-month actuals CSV plus three-scenario assumptions JSON. It creates Sources, Assumptions, Forecast, Dashboard, and Checks sheets with formula-backed values, scenario validation, cash warnings, root threaded comments, line and pie charts, frozen headers, second import, model/native render review, and a byte-bound audit; never hard-code derived values or mutate the input files.
- `examples/officekit-accessibility-audit-workflow.mjs`: Use this read-only existing-XLSX workflow to bind the source SHA-256, run `workbook.auditAccessibility()` plus `workbook.verify()`, and publish one no-overwrite JSON report with `savePolicy.strategy: none`. Review `artifact_tool_docs/ACCESSIBILITY_AUDIT.md`; a green modeled drawing check is not Excel Accessibility Checker or WCAG conformance.
- `examples/officekit-drawing-accessibility-edit-workflow.mjs`: Use this source-bound existing-XLSX transaction to repair exactly one uniquely named worksheet image or chart after binding its complete prior accessibility state and `accessibilityCapability`. It permits exactly one worksheet-drawing XML part to change, reimports the result, proves the non-target drawing projection and visual SVG projection unchanged, and publishes no-overwrite XLSX plus byte-bound audit outputs. It never edits reading order or claims whole-workbook conformance.

## Domain Requirements
You must read these domain rules when the request clearly relates to the domain, but do not load domain guidance for unrelated tasks unless asked:
- Finance and investment banking: `domain_guidance/financial_models.md`
- Corporate finance and FP&A: `domain_guidance/corporate_finance_fpa.md`
- Healthcare: `domain_guidance/healthcare.md`
- Marketing and advertising: `domain_guidance/marketing_advertising.md`
- Scientific research: `domain_guidance/scientific_research.md`

Instruction precedence for workbook content, layout, and formatting is: user request > reference/template > domain and formatting defaults.

## Making edits on a spreadsheet or using an uploaded reference or template.
- Before modifying, study and match the existing format, style, conventions,
  related values, and formulas. Render the relevant sheets. Inspect those
  images when visual input is available; otherwise use structural evidence and
  report the visual-review limitation.
- For visual fix requests, start with the smallest plausible local change. Do not apply sheet-wide autofit, wrapping, or restyling unless requested.
- Ensure existing formulas, layouts, structures, and patterns are consistent. For example, if asked to add another column or row to a table and there is conditional formatting applied to the whole table, it should extend to the new column or rows as well.
- Keep edits targeted unless a broader change is clearly necessary. Exceptions are when there's dependencies, e.g. a dynamic chart that is based on the range of values in a table and a new row is added, the chart should also update.
- Extend conditional formatting if needed to keep style consistent for an area or table.
- Never overwrite formatting for spreadsheets with established formats, unless requested or to extend an added range.

## Importing or extracting data from screenshots or reference images
- When a reference image or screenshot is provided, use appropriate data formats (e.g. number/date formats) based on the workbook topic, audience and purpose instead of trying to recreate the rendered format with just text. Preserve numeric/date usability even when the screenshot shows locale-specific punctuation or currency symbols.
- Use formulas when appropriate and correct: For screenshot recreation, do not bulk-write numeric tables as all static values until you have separated any clearly formula-derived ranges; test adjacent numeric rows/columns for exact repeated relationships such as sums, differences, products, ratios, or constant multiples, then keep inputs hardcoded and write derived ranges as formulas.
- Match visible styling, but do not infer intentional formatting from ambiguous image artifacts such as zoom, antialiasing, or compression. Infer font weight only from relative contrast or clear semantics; if all visible text has the same apparent weight, use normal weight.

## Handling queries and questions
- The user may ask questions about the sheet instead of requesting an edit or a change. Simply answer those questions about the spreadsheet based on the context available rather than making an edit the user didn't intend for. Use the selected workflow's read tools to inspect relevant values, formulas, tables, and objects.
- For a read-only question, do not modify or export the workbook.
- Locate the requested output by its row and column labels and period, inspect its displayed value and formula, and trace formula precedents to labeled assumptions or raw inputs instead of stopping at an intermediate total.
- Explain calculations with the workbook's displayed values and preserve units and period conversions. For broad questions about assumptions or drivers, rank the inputs that actually drive the requested output rather than inferring from nearby labels.

## Error Recovery
On first tool or API error:
1. Read error text.
2. Consult the selected workflow's targeted help or schema discovery only if needed.
3. Retry with minimal patch (not full rewrite).
4. Continue from existing workbook state.

Do not loop indefinitely on similar failures.

## Formula Rules
- Place assumptions and raw data in dedicated cells or clearly delineated input ranges, following the reference workbook's organization when one is provided.
- Keep lookup, mapping, scoring, and quality-control rules in visible cells or tables and reference them from formulas instead of hardcoding the logic.
- Derived values must be formulas (not hardcoded) and legible.
- Keep calculations formula driven, and prefer consistent formula patterns across a range where possible for readability. For example, formulas should be consistent across all projection periods.
- Use absolute/relative references correctly for fill/copy behavior.
- Use references instead of hardcoded or magic numbers inside formulas e.g. Use `=A5*(1+$A$6)` instead of =A5*1.05
- Formulas should be simple, legible and **easily auditable**. Use helper cells for intermediate values rather than performing complex calculations in a single cell. Users should be able to trace the model from inputs to outputs easily.
- No harcoded numbers inside calculation areas unless explicitly allowed. Always ensure color formatting conventions are properly applied.
- For any complex formulas or important assumptions, add comments to cells to explain.
- Always reference cells on other Excel sheets using the format ='Sheet Name'!A1, wrapping the sheet name in single quotes every time since quotes are required for any spaces or special characters.

### Ensure formulas are correct
- Checklist: No formula errors, all cell references are correct, no off-by-one errors in ranges, edge cases (zero values, negative numbers) are handled, no unintended circular references.
- For source-backed analyses and summaries, spot-check representative outputs and reconcile key totals with source definitions.

## Data Formatting Rules
- Store numbers, percentages, currency, and dates as typed spreadsheet values, not preformatted strings. Use text only for true identifiers such as ZIP codes, account IDs, SKUs, or labels.
- Use Excel-invariant number/date format codes, not locale-specific display strings. Examples include `#,##0`, `#,##0.0`, `0.0%`, `0.00%`, `"$"#,##0`, `"$"#,##0.00`, `yyyy-mm-dd`, `mmm yyyy` but choose the format that best fits the data.
- Percentages: When not specified or no reference is provided, use 1 decimal for most internal/analytical cells, 0 decimals for user-facing/dashboard outputs, and 2 decimals where small differences in rates matter.
- Do not swap `.` and `,` in format codes to mimic locale separators; separators are controlled by spreadsheet/render locale. Use `0.0%`, not `0,0%`, and `#,##0`, not `#.##0`.
- Choose the appropriate format for readability. Match precision to meaning: counts use `#,##0`; rates usually use `0.0%` or `0.00%`; currency uses whole units unless cents matter.

## Quality Guidelines
- Build correct, readable workbooks for the intended audience with clear structure, consistent formatting, reliable formulas, and useful outputs. Keep them as simple as practical.
- After autofit and wrapping, cap oversized column widths and row heights.
- Make workbooks easy for another person to update, trace, and audit without the original author.

## Completion Criteria
### Criteria for Question / Read only requests
- Answer from the available workbook context. Do not edit or overwrite unless the user asks for a workbook change.

### Criteria for all create and edit requests
Complete only when:
- Workbook content is populated and formulas compute.
- No obvious formula errors in key scanned ranges (no bad refs/off-by-one/circular errors).
- Save final `.xlsx` files under `outputRoot` (normally
  `workspaceRoot/outputs`) unless the user explicitly gives another path.
- Visual render verification passes:
  - Layout is organized, legible, and aligned to request style (or default/existing formatting baseline for edits).
  - Important numbers and callouts are all visible.
  - Numbers, text, charts and content is not clipped or awkwardly wrapped.

## Verification Rules
Before final response, verify values/formulas and visual quality.

Before delivering a workbook that contains images or charts, run
`workbook.auditAccessibility()`. Resolve machine issues, review every manual
check, and do not infer image descriptions from filenames or visible chart
titles. Use the packaged read-only workflow when the request is an audit of an
existing XLSX.

1. Inspect key ranges:
```js
const check = await workbook.inspect({
  kind: "table",
  range: "Dashboard!A1:H20",
  include: "values,formulas",
  tableMaxRows: 20,
  tableMaxCols: 12,
});
console.log(check.ndjson);
```

2. Scan formula errors:
```js
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);
```

3. Render sheets/ranges to verify visual output (skip if already verified and no style changes):
```js
const blob = await workbook.render({ sheetName: "Sheet1", range: "A1:H20", scale: 2 });
```
Make sure you do at least one visual pass of all the sheets in the workbook before the final export.

Visual requirements:
- Fix severe defects before finalizing: blank/broken charts, clipped key headers or numbers, unreadable colors, obvious formula errors, default blank sheets, or content outside the visible working area.
- Ensure logical labels or titles appear once, and merged ranges exist where labels or content intentionally span multiple columns.
- Ensure texts are all clearly visible and NOT clipped, columns and appropriately sized
- Do focused visual repair pass(s) after the initial render. Limit looping/time sinks for minor polish: stop once the workbook is correct, legible, and exported; note any minor limitation briefly and finalize.

4. Keep verification compact:
- Inspect key ranges.
- Avoid huge NDJSON dumps.

5. Export:
```js
await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/output.xlsx`);
```

6. Finalize immediately after successful export + compact verification.
- Do not export extra `.xlsx` variants unless asked.

## Citation Requirements
### Cite sources inside the spreadsheet
- Use plain-text URLs in spreadsheet cells.
- For financial models, cite model-input sources in cell comments.
- For researched row-wise data tables, include source URLs in a dedicated source column.

## Result and evidence

Return the final workbook as an absolute path with `kind: "workbook"` and its
SHA-256. Include the narrowest inspected sheet/range or object locator for
material claims, plus render/inspect/verify evidence paths when available.
Report `visualReview: "complete"` only after the required renders were
understood; otherwise use `"unavailable"` or `"requires-human"`. Do not emit
a host-specific citation directive. See `../office-kit/references/workspace.md` for
the result envelope. Semantic review must cover the requested ranges, formulas,
errors, tables, charts, and sheet identities; structural review must reopen the
actual XLSX. The text reading view may reduce reading cost for a large workbook,
but it cannot
validate formula topology or rendered workbook geometry.

## Comment Author
- If the authenticated/user profile or env context provides a user display name, use it as the threaded comment display name unless the user requests another name. Default to `User`.


## Source, PDF, and Attachment Processing
- Keep source notes compact: record file name, section/table label, and enough context to audit the number. Do not paste large PDF excerpts into the workbook unless requested.
- Bundled Python libraries available in the bundled runtime environment for extraction/analysis include `pandas`, `numpy`, `pypdf`, `python-docx`, and `reportlab`. You may read/extract in separate scripts if needed.
- Bundled JS libraries available for document/PDF work include `docx`, `pdf-lib`, and `pdfjs-dist`.
