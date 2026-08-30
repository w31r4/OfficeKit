# Design

`table.style.headerRows` remains the semantic count. For each cell, authored
style precedence is:

```text
cell fill/textStyle
→ headerCellFill/headerTextStyle when row < headerRows
→ defaultCellFill/defaultTextStyle
```

The compiler writes direct cell fills and rich DrawingML text bodies, so every
header row remains editable. It sets the native `firstRow` table flag whenever
the semantic count is non-zero. Exact counts above one are recovered from the
embedded PPJ; a third-party import without that program conservatively projects
only the native first-row fact.

No protocol field is required because header styling is lowered into existing
`PresentationTableCell.fill` and `PresentationTableCell.text_body` state before
the canonical PPTX writer runs.
