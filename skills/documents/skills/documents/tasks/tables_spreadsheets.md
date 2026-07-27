# Task: Tables ↔ spreadsheets (import/export)

## Goal
Move tabular data between Excel and Word reliably, without hand-copying.

## Import XLSX → DOCX table (simple)
Use the helper to convert a sheet into a Word table:

```bash
python scripts/xlsx_to_docx_table.py /mnt/data/input.xlsx /mnt/data/table.docx --title "Table: Results"
```

What it preserves (best-effort)
- cell values (data_only)
- basic alignment (left/center/right)
- header rows as bold
- column widths (heuristic)

What it does **not** preserve
- merged cells, formulas, charts, conditional formatting, complex number formats

## Export DOCX table → CSV

```bash
python scripts/docx_table_to_csv.py /mnt/data/input.docx --table_index 0 --out /mnt/data/table0.csv
```

## Edit one imported fixed-layout table without rebuilding it

For a request such as “make the first column wider but keep this Word table's
overall width,” inspect the imported table first. Record its **document block
index** and its complete `columnWidthsDxa` array. Do not select by nearby text
or by a guessed visual position.

Use the bounded public transaction:

```bash
officekit run examples/officekit-table-column-widths-edit-workflow.mjs \
  input.docx output.docx audit.json 1 \
  '[2100,4500,2700]' \
  '[3000,3600,2700]'
```

The workflow is intentionally narrow:

- It accepts one imported flat, rectangular, unmerged table with a complete
  recognized fixed-layout direct-formatting profile.
- The replacement retains both the column count and the exact total table
  width. It synchronizes `w:tblGrid/w:gridCol` with every physical cell's
  `w:tcW`; it does not only change the visible grid.
- It preserves text, table style, borders, margins, indentation, row/cell
  topology, and every other package part. Only `word/document.xml` may differ.
- It protects the source, refuses output/audit overwrites, reimports the table,
  verifies a raw XML residual outside the owned width leaves, and records a
  byte-bound audit.

Merged tables, nested tables, content-control cells, irregular widths,
unrecognized formatting, a changed total width, or a stale source value fail
closed. This is a width redistribution operation, not automatic table design:
inspect Word or LibreOffice plus Poppler output for wrapping, clipping, and
page-flow changes before delivery.

## Edit one imported direct table-formatting profile without rebuilding it

For a request such as “make this table's header softer, use a thinner blue
grid, and add a little cell padding,” keep the imported table's geometry fixed.
Inspect its document block index and record the entire current direct profile;
do not submit a partial style patch or a visual guess.

```bash
officekit run examples/officekit-table-formatting-edit-workflow.mjs \
  input.docx output.docx audit.json 1 \
  '{"indentDxa":120,"cellMarginsDxa":{"top":80,"bottom":80,"start":120,"end":120},"borderColor":"445566","borderSize":8,"headerFill":"E2E8F0"}' \
  '{"indentDxa":240,"cellMarginsDxa":{"top":100,"bottom":120,"start":160,"end":180},"borderColor":"224466","borderSize":12,"headerFill":"DDEBF7"}'
```

The workflow accepts one imported, flat, rectangular, unmerged fixed-layout
table only when OfficeKit recognized the complete direct profile. It changes
exactly `w:tblInd`, all six uniform `w:tblBorders` leaves, four
`w:tblCellMar` leaves, and each first-row cell's canonical `w:shd` fill. Table
width, grid and physical-cell widths, text, style, rows, cells, merge state,
and all non-document parts remain bound. It protects the source, refuses
output/audit overwrites, reimports, verifies the raw residual, and writes a
byte-bound audit.

Do not use it for a custom table-style graph, mixed border treatment,
conditional/header-row styling, merged or nested tables, content-control cells,
or auto-fit/reflow. Those cases remain source-bound or require an explicit
package/design workflow. Review a native Word or LibreOffice plus Poppler
render before delivery: the transaction preserves geometry, but it does not
calculate host wrapping or pagination.

## Edit one imported repeat-header prefix without changing table styling

`headerFill` is a visual fill on the first row. Repeat-header semantics are a
different native property: `headerRowCount` counts contiguous leading physical
rows marked with `w:tblHeader`, which Word can repeat at page breaks. For a new
table, set it directly with `document.addTable({ ..., headerRowCount: 1 })`.

For an imported correction, bind the inspected block index and the complete
source/replacement counts:

```bash
officekit run examples/officekit-table-header-rows-edit-workflow.mjs \
  input.docx output.docx audit.json 1 1 2
```

The workflow accepts only one imported flat rectangular unmerged table whose
row properties have the canonical optional grid-offset profile and a contiguous
no-`w:val` `w:tblHeader` prefix. It changes only those marker leaves in
`word/document.xml`; text, visual `headerFill`, widths, style, rows/cells, and
all other package parts remain bound. It protects the source, refuses overwrite,
checks the raw residual, reimports the full table projection, verifies,
model-renders, and writes a byte-bound audit. Non-prefix, duplicate,
explicit-value, extension-bearing, merged, nested, content-control, stale, and
no-op sources fail closed. Review a native Word or LibreOffice plus Poppler
render before delivery because changed repeat headers can alter visible page
breaks even though the transaction does not calculate pagination.

## Keep one imported physical table row intact across a page boundary

`keepTogetherRows` is a zero-based set of physical row indexes. Each selected
row writes native `w:cantSplit`, which tells Word not to split that row across
pages. It does not keep several rows as a group, and it does not calculate
where a page break will land. For a new table, use
`document.addTable({ ..., keepTogetherRows: [1, 4] })` or
`table.setRowKeepTogether(1, true)`.

For one imported correction, bind one table block, one physical row, and the
complete old/new boolean state:

```bash
node examples/officekit-table-row-break-policy-edit-workflow.mjs \
  input.docx output.docx audit.json 1 2 false true
```

The workflow accepts only one imported flat rectangular unmerged table with
canonical optional `w:gridBefore`/`w:gridAfter`, no-`w:val` `w:cantSplit`, and
no-`w:val` `w:tblHeader` leaves in native order. It changes one selected
`w:cantSplit` leaf in `word/document.xml`, keeps any repeat-header state,
content, visual styling, widths, style, topology, and every other package part
bound, then reimports, verifies, model-renders, and emits a no-overwrite audit.
Duplicate, explicit-value, reordered, extension-bearing, merged, nested,
content-control, stale, and no-op sources fail closed. Review a native Word or
LibreOffice plus Poppler render before delivery because page flow can change.

## Render → PNG review checklist (tables)
- Table fits within margins (no clipped columns)
- Header row is visually distinct
- Numbers align consistently (esp. decimals)
- No unexpected wrapping that hurts readability

## Common pitfalls
- Word tables do not auto-match Excel column widths; you must verify visually.
- Multi-line cells and merged cells round-trip poorly.
