## Why

The native Office codec already authors, imports, and edits literal scatter and
bubble charts, but PPJ cannot express their numeric X channel or bubble sizes.
The schema advertises both chart families while source-free compilation remains
unusable, which is a discoverability and contract defect rather than a missing
chart engine.

## What Changes

- Add bounded `xValues` and `bubbleSizes` vectors to PPJ chart series.
- Validate vector lengths and chart-family applicability before compilation.
- Lower numeric channels into the existing native chart series owner.
- Project recognized scatter and bubble caches back into PPJ.
- Keep source-bound numeric X and bubble-size caches immutable unless a future
  capability explicitly owns them.
- Update the generated PPJ manual and focused chart guidance.

## Capabilities

### New Capabilities

- `ppj-numeric-chart-series`: Defines finite numeric X and bubble-size channels
  for source-free scatter and bubble charts.

### Modified Capabilities

None.

## Impact

- Affected code: PPJ schema/models/validator, authored compiler, projector,
  source-bound chart-data boundary, generated PPJ reference, focused chart
  guidance, and the existing integrated PPJ chart test.
- The Office wire remains version 2 because `x_values` and `bubble_sizes`
  already have native owners.
- No new chart family, generic ECharts layer, or raw property bag is introduced.
