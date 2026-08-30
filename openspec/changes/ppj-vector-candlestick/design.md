## Context

Standard DrawingML ChartParts do not provide a portable authored candlestick
profile in the current OfficeKit writer. Rasterizing a financial chart would
make PPJ appear broad while losing native editability. PPJ already has ordered
chart categories, semantic series, editable group/shape/text writers and exact
embedded-program recovery.

## Decisions

### Reuse `values` for close and add three aligned channels

One candlestick series owns `highValues` and `lowValues`; `values` remains the
close channel. Optional `openValues` selects OHLC mode. Omitting it selects HLC
mode. Every channel must align with ordered string categories, be finite and
obey `low <= open/close <= high`.

The bounded profile accepts one series and at most 64 observations. It excludes
mixed overlays, volume, logarithmic scales, date interpolation and missing
points. Those can be added only with a separate truthful semantic contract.

### Lower to a native vector group

The compiler emits one `p:grpSp`. Wicks are thin editable rectangles. OHLC
bodies are editable rectangles with separate rise/fall paint. HLC close values
use a short right-facing tick. Titles, category labels, value ticks and optional
close labels remain native text boxes. Deterministic compiler-owned IDs use a
slash-delimited namespace that cannot collide with PPJ IDs.

The lowering is not inferred in reverse. Embedded PPJ restores candlestick
semantics. Without it, ordinary PPTX import returns the truthful group.

### Reuse bounded axis state

`yAxis.min`, `max`, `majorUnit`, `title` and text styles control the price scale;
explicit bounds must contain every high/low. `xAxis.title`, text styles and
`tickLabelInterval` control ordered category labels. Auto bounds add a small
padding and choose a finite four-interval scale.

### Keep style finite

`style.candlestick` requires `up`, `down` and `wick`, with a body-width ratio,
optional gridline stroke, value labels and the existing chart text-style shape.
Ordinary ChartPart legend, stacking, markers, trendlines, error bars, data-label
and surface-fill fields fail closed.

## Risks and mitigations

- **Unreadable dense plots:** reject a frame whose category slot cannot hold a
  visible body instead of silently hiding evidence.
- **Misleading prices:** validate all OHLC inequalities and explicit axis
  bounds before output.
- **False native-chart claim:** document and inspect the result as an editable
  vector group, not a ChartPart.
- **Semantic loss after snapshot removal:** fallback projection is deliberately
  an ordinary group.
