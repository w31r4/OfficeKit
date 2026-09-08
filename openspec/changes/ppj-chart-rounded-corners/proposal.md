## Why

F-07 (P0) still lacks ChartML settings that PPJ can author and edit independently.
Chart-space rounded corners currently have no PPJ field or native round-trip owner.

## What Changes

- Add optional boolean `chart.roundedCorners` with absent/false/true presence.
- Read and write direct `c:chartSpace/c:roundedCorners` for ordinary and bounded combo charts.
- Allow source-bound addition, replacement and removal through `setChartPlot`.
- Add one focused round-trip experiment covering both chart paths and reject unsupported topology.

## Capabilities

### New Capabilities
- `ppj-chart-rounded-corners`: Typed chart-space corner setting with bounded round-trip editing.

### Modified Capabilities

None.

## Impact

PPJ schema/capability metadata, additive presentation and spreadsheet chart wire fields,
shared ChartML codec, PPJ compiler/projector/validator, focused tests and gap documentation.
No protocol version bump or new dependency.
