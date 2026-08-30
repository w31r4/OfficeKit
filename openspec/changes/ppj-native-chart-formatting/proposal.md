## Why

PPJ already owns native chart data, typography, paint and circular geometry,
but several ordinary formatting decisions are still fixed by the compiler or
lost when an editable third-party chart is projected. Agents cannot format
data-label values, reverse an axis, style axis/grid lines or control how native
bubble sizes are interpreted.

## What Changes

- Add `style.dataLabels.numberFormat` for plot-level native labels.
- Add native bubble `style.bubbleScale` and `style.bubbleSizeMode` controls.
- Add `reverse`, `axisLine` and `gridLine` to PPJ chart axes.
- Preserve the new values through authored compilation, PPTX projection and
  source-bound local edits.
- Document where PPJ intentionally differs from renderer-only pixel-radius
  controls.

## Capabilities

### New Capabilities

- `ppj-native-chart-formatting`: bounded label, bubble and axis formatting
  semantics across PPJ, ChartSpace and source-bound continuation.

### Modified Capabilities

None. The PPJ schema ID and Office wire version remain unchanged.

## Impact

The PPJ schema and generated manual, additive wire-v2 chart messages, shared
ChartSpace codecs, PPTX chart adapter, PPJ authored/projected/source-bound
compiler, capability registry, chart guidance, coverage and one existing PPJ
contract are affected.
