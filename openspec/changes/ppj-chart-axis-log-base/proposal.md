## Why

F-07 still lacks logarithmic axes. XlsxChartAxisCodec currently rejects c:scaling/c:logBase as uneditable, and PPJ has no field for its base.

## What Changes

- Add optional chart axis logBase, a finite number from 2 through 1000 or a size grammar token.
- Author, project, add, change and remove logarithmic scaling on native value axes, including numeric x axes and combo secondary value axes.
- Support the same value-axis field as radar spokeAxis.logBase so projection and deletion preserve its editing route.
- Reject category-axis use and invalid/ambiguous owners; preserve unrelated package parts.

## Capabilities

### New Capabilities
- `ppj-chart-axis-log-base`: Presence-aware native logarithmic axis scaling.

### Modified Capabilities
None.

## Impact

PPJ schema/compiler/projector/validator, shared additive axis wire field, ChartML axis codec, XLSX JavaScript preservation during unrelated edits, focused tests and reference/backlog docs. No new dependency or wire version bump.
