## Why

F-07 chart text still rejects direct baseline attributes, so superscript/subscript placement in otherwise supported chart text cannot be expressed or edited. Ordinary PPJ text already defines baseline as a signed percentage from -400 to 400.

## What Changes

- Add chartTextStyle.baseline with the existing ordinary-text percentage range.
- Preserve explicit zero separately from omission, using native thousandths of a percent for wire state and native edits.
- Propagate through chart style owners, trendline rich text, nested precedence and vector text; more-specific title-run baseline wins.
- Verify creation/change/reset/removal/recreation, precision conversion, invalid values and unrelated wire preservation with minimal fixtures.

## Capabilities

### New Capabilities
- `ppj-chart-text-baseline`: Direct chart character baseline percentage and its edit lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema/protobuf, shared ChartML style codec, authored/source-bound builders, projector, vector text, focused tests and chart documentation. No new runtime dependency or host measurement requirement.
