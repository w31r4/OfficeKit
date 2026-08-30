## Context

The native codec already owns bounded category/value axes, direct marker style,
trendlines, and error bars. PPJ currently exposes only axis visibility, one
marker symbol, and value-label visibility. The missing layer is semantic
language mapping, not a new chart engine.

## Decisions

### Keep the authored vocabulary bounded

Axes own visibility, title, number format, tick interval, numeric bounds, major
unit, and tick-label font size. Markers own symbol, size, solid fill, and direct
stroke. Labels own value/category/series/percent visibility and position.
Trendlines and non-formula error bars reuse the codec's closed native enums.

### Keep compatibility spellings

The existing marker string and `showDataLabels`/`dataLabelPosition` fields remain
valid. Structured marker and `dataLabels` objects add depth without changing
old programs. Conflicting compatibility and structured fields fail closed.

### Separate authored state from imported authority

Projection may describe recognized state, but a third-party chart may change
only through an issued capability. `setChartData` continues to own names and
values; it cannot smuggle marker, line, trendline, error-bar, or axis changes.

## Verification

Extend the existing canonical PPJ build/reimport test with one axis, marker,
trendline, and error-bar example. Run that test, schema/manual sync, and strict
OpenSpec validation. Do not create a chart matrix.

