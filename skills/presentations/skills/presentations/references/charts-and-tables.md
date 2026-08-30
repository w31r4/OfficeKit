# Charts and tables

Choose the visual from the relationship the audience must understand.

| Relationship | Prefer |
| --- | --- |
| trend or change over ordered time | line, area, or ordered columns |
| magnitude across categories | bars or columns with a common baseline |
| distribution or correlation | scatter, bubble, or distribution view |
| intensity across two categorical dimensions | heatmap |
| multivariate profile across a small shared scale | standard radar |
| contribution to change | waterfall |
| precise lookup or rule comparison | table |
| mixed measures sharing a truthful category axis | deliberate combo chart |

Avoid pies with many categories, decorative 3D, dual axes without a necessary
and clearly labelled relationship, truncated scales that exaggerate change,
and a chart that merely repeats one large number.

## Evidence integrity

- Preserve source values, units, time windows, denominators, and uncertainty.
- Use direct labels where they reduce legend lookup.
- Keep common scales for comparisons that claim comparability.
- Show missing values and estimates honestly.
- Put the exact source and as-of date on the page or in a visible source area.

Use JSON `null` for a genuinely missing Y observation. OfficeKit keeps the
logical point count and writes a native chart gap; it does not coerce the value
to zero or invent an estimate. This works for the bounded authored chart
families, including numeric scatter/bubble Y values. Numeric `xValues` and
`bubbleSizes` remain complete arrays. In a source-bound chart, edit measured
values around an existing gap, but do not add or remove gaps: missing-point
topology remains tied to the exact source cache.

Lines, markers, labels, confidence intervals, error bars, and decision
thresholds are protected foreground evidence. Bars, areas, fills, masks, and
annotations may not hide them. For truthful combo charts, first keep the real
shared plot, render the line above pale or transparent bars, reserve clearance
around markers, and re-anchor only colliding labels. Split into aligned panels
only when the units are incompatible or the truthful shared plot remains
illegible after those repairs.

Do not force separation merely to make a checker quiet. Do not change scale,
data, or category position to manufacture whitespace.

## PPJ structure

Chart elements keep a stable `id`, explicit `frame`, typed `chartType`, data,
and style. Tables declare column widths and typed rows; they are not collections
of manually aligned text boxes. Use chart/table styles to define semantic
roles, then override only when a specific datum needs emphasis.

Chart and table frames may use `rotation`, `flipH`, and `flipV`. These are real
native graphic-frame transforms and survive PPJ build, import, and a
capability-issued source edit. Prefer ordinary zero-degree evidence layouts;
rotate a chart or table only when the orientation itself communicates a real
spatial relationship, never to decorate a page.

The authored chart compiler owns these native visual controls:

- bar, column, line, area, pie, doughnut, scatter, bubble, standard radar,
  bounded semantic waterfall, bounded vector heatmap, and bounded bar-line
  combo plots;
- legend visibility and top, bottom, left, or right placement;
- ordinary, stacked, and percent-stacked grouping where the chart family
  supports it;
- bar direction, gap width, category/value-axis visibility, major gridlines,
  data-label visibility and bounded label position;
- chart-area, plot-area, and series none, solid, or bounded direct-RGB gradient
  fills, including solid and per-stop opacity;
- editable series line width, dash, opacity, cap, join, and bounded markers;
- category/value-axis titles, number formats, label interval, tick-label
  typeface, size, bold, italic and direct RGB/alpha, value bounds and major
  unit; bounded combo charts may declare the matching secondary pair;
- chart-title Latin/East Asian typeface, size, bold, italic and direct
  RGB/alpha plus canonical line-chart smoothing and direct color variation;
- structured data labels for value, category, series, percentage and native
  position;
- direct marker symbol, size, RGB/alpha fill and bounded stroke;
- exponential, linear, logarithmic, moving-average, polynomial and power
  trendlines on bar, column and line series;
- fixed-value, percentage, standard-deviation and standard-error error bars on
  bar, column and line series.

Use the analytical fields as chart semantics, not as decorative paint. A
typical evidence series can be expressed without replacing the chart with
shapes:

```json
{
  "xAxis": { "title": "Quarter", "tickLabelInterval": 1 },
  "yAxis": { "title": "Conversion rate", "numberFormat": "0.0%", "min": 0, "max": 0.4, "majorUnit": 0.1 },
  "style": {
    "dataLabels": { "showValue": true, "position": "outside-end" }
  },
  "data": {
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": [{
      "id": "conversion",
      "name": "Conversion",
      "values": [0.18, 0.22, 0.27, 0.31],
      "marker": { "symbol": "circle", "size": 7, "fill": "#FFFFFFCC", "stroke": { "color": "#16697A", "width": 1 } },
      "trendlines": [{ "type": "linear", "stroke": { "color": "#D9A21B", "width": 1.25, "dash": "dash" } }],
      "errorBars": { "valueType": "standard-error", "direction": "y", "type": "both" }
    }]
  }
}
```

Use title typography and line behavior as explicit chart state rather than
rebuilding the visual with shapes:

```json
{
  "chartType": "line",
  "title": "Retention after launch",
  "style": {
    "titleTextStyle": {
      "fontSize": 16,
      "fontFamily": "Aptos Display",
      "fontFamilyEastAsia": "Noto Sans CJK SC",
      "bold": true,
      "italic": false,
      "color": "#16324FCC"
    },
    "legend": "bottom",
    "legendTextStyle": {
      "fontSize": 9,
      "fontFamily": "Aptos",
      "color": "#475569"
    },
    "dataLabels": {
      "showValue": true,
      "position": "above",
      "textStyle": {
        "fontSize": 8,
        "bold": true,
        "color": "#16324F"
      }
    },
    "smooth": false,
    "varyColors": true
  },
  "xAxis": {
    "title": "Quarter",
    "titleTextStyle": {
      "fontSize": 10,
      "fontFamily": "Aptos Display",
      "bold": true,
      "color": "#16324F"
    },
    "textStyle": {
      "fontSize": 9,
      "fontFamily": "Aptos",
      "fontFamilyEastAsia": "Noto Sans CJK SC",
      "color": "#475569"
    }
  }
}
```

`smooth` preserves an explicit true or false native value. `varyColors: true`
authors one direct native color-variation flag; false is canonical omission.
Both line-behavior fields are line-chart-only. Chart-title, legend, data-label,
axis-title and axis tick-label typography compile through one exact DrawingML
profile on ordinary and combo charts. A projected imported chart may issue
`setChartTextStyle`; only that capability authorizes changes to these fields.
Ordinary `setChartTitle` and `setChartData` do not authorize a style mutation.
Theme transforms, shadows, effects and irregular rich-text topology remain
source-owned and fail closed instead of being flattened.

Chart paint uses the same typed fill union as shapes and table cells. This
keeps a gradient semantic rather than rebuilding it from overlaid rectangles:

```json
{
  "style": {
    "chartAreaFill": { "type": "none" },
    "plotAreaFill": {
      "type": "gradient",
      "kind": "radial",
      "stops": [
        { "offset": 0, "color": "#F8FAFC", "opacity": 0.96 },
        { "offset": 1, "color": "#DCE7F2", "opacity": 0.72 }
      ]
    }
  },
  "data": {
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": [{
      "id": "revenue",
      "name": "Revenue",
      "values": [18, 24, 31, 39],
      "fill": {
        "type": "gradient",
        "kind": "linear",
        "angle": 90,
        "stops": [
          { "offset": 0, "color": "#0EA5A8" },
          { "offset": 1, "color": "#0B3A5B", "opacity": 0.82 }
        ]
      }
    }]
  }
}
```

For a simple solid series, `"color": "#0A84FF80"` is a compact authored
alias for the equivalent solid `fill` with alpha. Reimport canonicalizes it to
the structured fill form so the opacity remains explicit and subsequent edits
do not oscillate between two spellings. Use `fill` directly when the series
needs none/gradient paint or a separately named opacity.

The bounded profile accepts two through sixteen ordered direct-RGB stops,
linear angles, and centered radial gradients. Theme transforms, pattern/image
paint, path variants, and irregular native gradient graphs remain
source-preserved. An imported chart must issue `setChartFill`; neither
`setChartData` nor `setChartTextStyle` authorizes paint changes.

The scalar marker spelling and `showDataLabels` / `dataLabelPosition` remain
valid for older PPJ. Do not combine either legacy spelling with its structured
form. On an imported source-bound chart, a `setChartData` capability owns only
series names and values; it cannot be used to smuggle axis, marker, label,
trendline, error-bar or paint changes.

Scatter and bubble charts use numeric channels rather than shared category
labels:

```json
{
  "chartType": "bubble",
  "data": {
    "categories": [],
    "series": [{
      "id": "opportunities",
      "name": "Opportunities",
      "xValues": [10, 20, 34],
      "values": [5, 12, 8],
      "bubbleSizes": [4, 9, 16]
    }]
  }
}
```

`xValues` is required for scatter and bubble; `bubbleSizes` is required only
for bubble. Each vector has exactly the same logical length as `values`, X
values are finite, bubble sizes are positive, and `categories` is empty. Only
Y `values` may contain `null`. Do not encode numeric X values as strings merely
to reuse a category chart. On imported charts, the current `setChartData`
capability does not authorize changing X values, bubble sizes, or the positions
of missing Y observations.

Use `waterfall` for a cumulative bridge whose opening/closing totals and signed
changes must remain explicit. Author one semantic series rather than exposing
the four implementation series used by the native lowering:

```json
{
  "type": "chart",
  "id": "operating-bridge",
  "chartType": "waterfall",
  "frame": { "x": 72, "y": 112, "width": 640, "height": 300 },
  "title": "Operating bridge",
  "yAxis": { "title": "Run-rate", "min": 0, "max": 180, "majorUnit": 30 },
  "style": {
    "legend": "none",
    "gapWidth": 55,
    "waterfall": {
      "increase": { "label": "Increase", "fill": { "type": "solid", "color": "#0B8F8F" } },
      "decrease": { "label": "Decrease", "fill": { "type": "solid", "color": "#C8644A" } },
      "total": { "label": "Total", "fill": { "type": "solid", "color": "#16324F" } }
    }
  },
  "data": {
    "categories": ["Opening", "Growth", "Churn", "Cost", "Closing"],
    "series": [{
      "id": "run-rate",
      "name": "Run-rate",
      "values": [120, 40, -25, -10, 125],
      "pointRoles": ["total", "delta", "delta", "delta", "total"]
    }]
  }
}
```

A `total` value is absolute. A `delta` value is added to the running total;
every later total must equal that computed value. The bounded profile keeps the
cumulative value non-negative and requires explicit increase, decrease, and
total styles. It deliberately omits legends, automatic data labels, secondary
axes, trendlines, markers, and error bars because those would expose or confuse
the hidden offset series. OfficeKit compiles the result to one editable native
stacked-column ChartPart; it does not fake the bridge with rectangles. Exact
waterfall intent recovers from the embedded PPJ. If that snapshot is removed,
ordinary PPTX import truthfully exposes the four native series instead of
guessing that an arbitrary stacked chart is a waterfall.

Use `heatmap` only when the audience must compare a genuine two-dimensional
matrix. `data.categories` are the ordered columns; each series `name` is one
row label and its `values` are that row. Do not use a heatmap as a decorative
grid or to disguise a small ranked list that bars would explain more clearly.

```json
{
  "type": "chart",
  "id": "segment-signal-matrix",
  "chartType": "heatmap",
  "frame": { "x": 72, "y": 112, "width": 640, "height": 300 },
  "title": "Observed relationship strength",
  "style": {
    "heatmap": {
      "scale": "diverging",
      "colors": ["#C8644A", "#F8F6EF", "#0B8F8F"],
      "domain": [-10, 10],
      "midpoint": 0,
      "showValues": true,
      "showColorBar": true,
      "cellGap": 2,
      "missingFill": "#E5E7EB",
      "cellStroke": { "color": "#FFFFFF", "width": 0.5 },
      "axisTextStyle": { "fontSize": 8, "color": "#52606D" },
      "valueTextStyle": { "fontSize": 8, "bold": true }
    }
  },
  "data": {
    "categories": ["Acquisition", "Retention", "Margin", "Reliability"],
    "series": [
      { "id": "enterprise", "name": "Enterprise", "values": [8, 5, 2, null] },
      { "id": "mid-market", "name": "Mid-market", "values": [4, -2, 6, 7] },
      { "id": "smb", "name": "SMB", "values": [-6, -4, 1, 3] }
    ]
  }
}
```

The bounded profile accepts 1–32 rows and 1–32 columns, unique non-empty
labels, explicit missing cells, a two-color linear scale or three-color
diverging scale, optional value labels, and one editable vertical color bar.
An explicit diverging domain must contain its midpoint; values outside the
domain clamp to its endpoint colors. Tiny frames reject instead of silently
making labels unreadable.

PowerPoint has no standard native heatmap ChartPart. OfficeKit therefore
lowers this semantic node to one editable DrawingML group of rectangles and
text, not a PNG. The embedded PPJ restores the exact matrix intent. If that
program is removed, import returns the truthful ordinary group rather than
guessing that an arbitrary shape grid was a heatmap. Whole-object animation is
valid; `chartBuild` modes are not, because there is no native ChartPart.

Use radar only when every series is measured against the same small set of
meaningful dimensions and a common scale. It is a profile comparison, not a
replacement for precise lookup or ordered ranking:

```json
{
  "type": "chart",
  "id": "risk-profile",
  "chartType": "radar",
  "frame": { "x": 72, "y": 110, "width": 420, "height": 300 },
  "title": "Operating resilience",
  "data": {
    "categories": ["Liquidity", "Growth", "Margin", "Resilience"],
    "series": [{
      "id": "current",
      "name": "Current",
      "values": [72, 81, 64, 77],
      "stroke": { "color": "#0A84FF", "width": 2 },
      "marker": { "symbol": "circle", "size": 5 }
    }]
  },
  "xAxis": {},
  "yAxis": { "min": 0, "max": 100, "majorUnit": 20 }
}
```

PPJ compiles this to editable native `standard` radar rather than drawing a
polygon from shapes. Category/value axes, series stroke, bounded markers,
legend, labels, chart surfaces and fixed-topology source continuation use the
same chart contracts as the other category families. Filled, marker-only, 3D,
extension-bearing and irregular native radar variants remain source-owned.

Chart image/pattern paint, theme-transformed gradients, irregular sparse
caches, and unrecognized waterfall/ChartEx graphs still fail closed. Existing
unsupported native chart graphs remain source-preserved; they are not
simplified during an unrelated imported edit.

The authored table compiler owns a physical column/row grid, finite rectangular
merges, one optional header row, row/column banding flags, bounded rich text,
body and paragraph layout, none/solid/gradient/image cell fills, and direct
left, top, right, and bottom borders. An image fill uses the same local hashed
asset, crop, cover/contain/stretch/tile, and opacity contract as other PPJ image
paint. Use it when the cell itself is an evidence thumbnail, product identity,
or comparison image; do not turn ordinary data tables into decorative mosaics.
Named table styles provide defaults and inline style properties override them
field by field. More than one header row fails closed. Imported table topology
and unmodeled native style graphs remain source-owned.

Use [the complete PPJ reference](ppj.md) for exact fields, value ranges, and
compiler boundaries. This page explains visual choice; it is not a shortened
substitute for the language manual.

Render high-risk pages at final dimensions. Verify axes, labels, legends,
units, source text, value precision, empty cells, merged-looking boundaries,
and whether the visual still communicates without animation.
