# Change: PPJ multi-header authored tables

## Why

High-density financial, research, and operating-review slides regularly use two
or more semantic header rows. PPJ already accepts `table.style.headerRows` up to
100, but the authored compiler rejects every value above one. This makes valid
programs fail at build time and forces Agents to fake tables with loose shapes.

## What changes

- Make `headerRows` style the first N physical rows of a source-free native
  PowerPoint table.
- Add optional `headerCellFill` and `headerTextStyle` fallbacks to `tableStyle`.
- Preserve cell-local fill and text style as the highest-precedence authority.
- Reject a header count larger than the physical row count or header styling
  without at least one header row.
- Keep imported table topology and styling source-owned.

## Scope

This change does not add a second table model, infer headers from formatting,
or claim that PowerPoint exposes a native multi-row header flag. Authored PPJ
retains the exact semantic count through the embedded program; the PPTX carries
the intended appearance as direct editable cell formatting and uses the native
first-row flag for ordinary Office behavior.
