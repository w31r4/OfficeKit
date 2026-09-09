## Why

F-07 still cannot express custom error-bar data in PPJ, although the shared native codec already reads and writes it. Exposing literal plus/minus data completes the local per-point uncertainty fields after the scalar lifecycle change.

## What Changes

- Add `valueType: "custom"` and `plus` / `minus` objects containing `values` and optional `formatCode` to PPJ errorBars.
- Support authored/imported literal data, side selection, value/format edits, custom/scalar conversion, deletion and recreation for column/bar/line and categorical combo series.
- Preserve unprojected formula-backed error data and reject attempts to overwrite it through the literal fields.
- Reuse the explicit partial SVG diagnostic for error bars and update the gap/reference evidence.

## Capabilities

### New Capabilities
- `ppj-chart-custom-error-data`: Literal per-point plus/minus error-bar data with full local ownership.

### Modified Capabilities
None.

## Impact

PPJ schema and semantic validation, authored/source-bound compilers, projection, shared native presence policy, focused tests and generated references. Existing protobuf data fields are reused. Formula/workbook closure remains a separate F-07 gap; this change does not flatten formula sources.
