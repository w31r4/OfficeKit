## Why

PPJ now owns the main native chart frame and plot styling, but several high-value
analytical semantics already implemented by the Office codec remain hidden from
the language: axis bounds and titles, rich marker state, label content,
trendlines, and error bars. This makes the authored surface look smaller than
the underlying OfficeKit capability and encourages agents to approximate data
graphics with shapes.

## What Changes

- Add bounded PPJ axis, marker, data-label, trendline, and error-bar state.
- Lower that state through the existing typed chart codec rather than a generic
  style property bag.
- Project recognized native chart state back into PPJ for inspection.
- Prevent source-bound `setChartData` from silently accepting style changes.
- Generate the exhaustive PPJ reference and focused chart guidance from the
  same capability facts.

## Capabilities

### New Capabilities

- `ppj-analytical-chart-primitives`: Defines deterministic authored and projected
  chart evidence semantics without expanding the raw OOXML surface.

## Impact

- Affected code: PPJ schema/models/compiler/projector/source-bound differ,
  existing chart protobuf state, generated PPJ reference, and one integrated
  PPJ test.
- Office wire remains version 2; no new wire message is required for the first
  slice because the bounded native owners already exist.
- Unsupported chart types and irregular imported chart graphs remain
  source-preserved or fail closed.

