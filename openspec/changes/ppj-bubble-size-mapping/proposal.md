# Change: PPJ bounded bubble size mapping

## Why

PPJ can author native bubble charts with Office's global scale and area/width
semantics, but it cannot say that visible radii must occupy a deliberate range
or that order-of-magnitude evidence needs logarithmic size mapping. Agents must
currently accept host-dependent bubble geometry or draw circles manually.

## What changes

- Add chart-level `bubbleSizeScale` and `bubbleRadiusRange` state.
- Apply one shared size domain across every bubble series in the chart.
- Lower explicit sizing to stable editable DrawingML circles with native axes,
  labels and legend expressed as ordinary editable elements.
- Keep ordinary bubble charts on the native ChartPart path when the new state is
  absent.
- Preserve exact authored PPJ through the embedded program; do not infer bubble
  semantics from arbitrary third-party groups.

## Impact

- Additive PPJ v1 schema and NativeAOT authored compiler state; Office wire 2 is
  unchanged.
- The existing comprehensive authored PPJ contract gains one bounded example.
- No provider, fixture matrix, snapshot farm or new test file is added.
