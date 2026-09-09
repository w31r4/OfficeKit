## Why

F-07 still lists trendline labels as source-owned. PPJ can request an equation or R-squared display, but cannot express the associated label text or appearance; importing a label also prevents analytics edits.

## What Changes

- Add `data.series[].trendlines[].label` with optional literal text, number format, text style, fill and line, reusing the existing chart-label value grammar.
- Support authored output and source-bound label insertion, replacement, deletion and recreation on ordinary bar/column/line and categorical combo series.
- Preserve automatic label content when text is omitted; an empty object owns the default native label container.
- Keep formula text, nonempty manual layouts, unsupported effects and extensions source-owned, with no silent flattening.

## Capabilities

### New Capabilities
- `ppj-chart-trendline-label`: A persistent native trendline label with complete bounded lifecycle and fresh projection.

### Modified Capabilities
None.

## Impact

Additive wire field and label message, generated JS bindings, PPJ schema, shared ChartML codec, authored/source-bound compilers, projection and documentation. Reuse existing text/fill/line codecs; no new runtime dependency. This advances F-07 without claiming automatic label layout or PowerPoint visual acceptance.
