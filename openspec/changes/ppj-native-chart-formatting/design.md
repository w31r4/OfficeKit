## Context

DrawingML already has stable native representations for these decisions:
`c:dLbls/c:numFmt`, `c:bubbleScale`, `c:sizeRepresents`,
`c:scaling/c:orientation`, and bounded `c:spPr/a:ln` graphs on axes and major
gridlines. OfficeKit owns the surrounding chart profile and can extend it
without changing chart data or relationships.

Kimi's `sizeRange` is not a native chart property: it is a pixel-radius rule in
its renderer. PPJ must not claim that a native editable PowerPoint chart can
recover a value that DrawingML cannot represent.

## Decisions

### 1. Expose native bubble semantics, not fake pixel geometry

`chart.style.bubbleScale` is an integer from 0 through 300 and maps to native
`c:bubbleScale`. `chart.style.bubbleSizeMode` is `area` or `width` and maps to
`c:sizeRepresents`. The former scales every bubble relative to the plot; the
latter chooses whether source values represent area or diameter. A Kimi-style
log transform or exact pixel radius must be performed explicitly on data or by
a non-native visual, not hidden in the compiler.

### 2. Keep label formats presence-aware

`chart.style.dataLabels.numberFormat` is an optional stable format code up to
255 characters. Projection preserves an existing direct `c:numFmt` only when
`sourceLinked=false`; authored and source-bound output write that canonical
form. Documentation recommends the portable subset `0`, `0.0`, `0%`, `0.0%`,
`#,##0`, and `0.0E+00` without rejecting other control-free native codes.

### 3. Model axis direction and line intent separately

`axis.reverse` maps to `orientation=maxMin`; omission preserves the native
default or absence. `axisLine` and `gridLine` each accept a boolean or a bounded
stroke object. `false` hides the line, `true` requests the native default and a
stroke object writes one direct RGB/width/dash/cap/join line. Grid-line absence
and visibility continue to use the existing presence-aware major-gridline
contract.

### 4. Source-bound edits remain local

Editable imported charts receive `setChartLabels`, `setChartPlot` and
`setChartAxis` only for semantics the codec actually proved. A PPJ diff may
change the modeled scalar or bounded line graph; chart type, series, axes,
relationships and unknown siblings are re-proved before the existing ChartPart
is patched.

## Lean verification

Extend one existing PPJ chart contract with a native bubble and one axis-rich
chart. Prove authored native XML, projection, capability issuance, a bounded
source edit and reimport. Do not create a property matrix, fixture farm or new
test file.
