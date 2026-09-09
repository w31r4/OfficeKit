## Why

F-07 still lacks formula-backed custom error data in PPJ. Literal fields are now complete, but editing an imported formula cache without its embedded worksheet would leave conflicting values.

## What Changes

- Project existing local range formulas as source-bound `errorBars.plus/minus.formula` alongside their complete values arrays and formats.
- Keep formula identity fixed while allowing error-value edits to update both the ChartPart cache and the owned embedded worksheet cells in one export.
- Reject missing/shared workbook ownership, mismatched values, overlapping chart consumers and dependent worksheet graphs.
- Preserve no-op and style-only behavior; keep formula creation/removal/retargeting and source-free workbook authoring explicitly separate.

## Capabilities

### New Capabilities
- `ppj-chart-error-data-workbook-sync`: Atomic cache/worksheet updates for existing error-bar formula bindings.

### Modified Capabilities
None.

## Impact

PPJ schema, projection and compilers, PPTX chart export, shared range and worksheet splice helpers, focused regression and chart documentation. Existing protobuf data fields are reused. This advances F-07 without claiming arbitrary formula or workbook topology editing.
