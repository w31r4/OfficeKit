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
node examples/officekit-table-column-widths-edit-workflow.mjs \
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

## Render → PNG review checklist (tables)
- Table fits within margins (no clipped columns)
- Header row is visually distinct
- Numbers align consistently (esp. decimals)
- No unexpected wrapping that hurts readability

## Common pitfalls
- Word tables do not auto-match Excel column widths; you must verify visually.
- Multi-line cells and merged cells round-trip poorly.
