## Why

PPJ can author pie and doughnut charts, but the native compiler fixes their
orientation at zero degrees and every doughnut hole at 50 percent. Agents
therefore cannot place the first slice deliberately, make room for a center
label, or retain those ordinary properties when a third-party chart is
projected and edited.

## What Changes

- Add `style.startAngle` for pie and doughnut charts.
- Add `style.holeSize` for doughnut charts.
- Preserve both values through authored compilation, PPTX projection and
  source-bound local edits.
- Issue one bounded `setChartPlot` capability for imported circular charts.
- Teach the Presentation Skill when the geometry carries information rather
  than decorative rotation.

## Capabilities

### New Capabilities

- `ppj-circular-chart-geometry`: bounded native first-slice angle and doughnut
  hole-size semantics across PPJ, ChartSpace and source-bound continuation.

### Modified Capabilities

None. The PPJ schema ID and Office wire version remain unchanged.

## Impact

The PPJ schema and generated manual, additive wire-v2 chart messages, shared
ChartSpace codec, PPTX chart adapter, PPJ authored/projected/source-bound
compiler, capability registry, chart guidance, coverage and one existing
comprehensive PPJ contract are affected.

