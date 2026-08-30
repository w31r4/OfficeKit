# Charts and tables

Choose the visual from the relationship the audience must understand.

| Relationship | Prefer |
| --- | --- |
| trend or change over ordered time | line, area, or ordered columns |
| magnitude across categories | bars or columns with a common baseline |
| distribution or correlation | scatter, bubble, or distribution view |
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

- bar, column, line, area, pie, doughnut, scatter, bubble, and bounded bar-line
  combo plots;
- legend visibility and top, bottom, left, or right placement;
- ordinary, stacked, and percent-stacked grouping where the chart family
  supports it;
- bar direction, gap width, category/value-axis visibility, major gridlines,
  data-label visibility and bounded label position;
- chart-area and plot-area none or solid fills, including opacity;
- direct solid series color plus editable line width, dash, opacity, cap, join,
  and bounded markers;
- category/value-axis titles, number formats, label interval, tick-label
  typeface, size, bold, italic and direct RGB/alpha, value bounds and major
  unit; bounded combo charts may declare the matching secondary pair;
- chart-title Latin/East Asian typeface, size, bold, italic and direct
  RGB/alpha plus canonical line-chart smoothing and direct color variation;
- structured data labels for value, category, series, percentage and native
  position;
- direct marker symbol, size, fill and stroke;
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
      "marker": { "symbol": "circle", "size": 7, "fill": "#FFFFFF", "stroke": { "color": "#16697A", "width": 1 } },
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
    "smooth": false,
    "varyColors": true
  },
  "xAxis": {
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
Both line-behavior fields are line-chart-only. Chart-title and axis tick-label
typography compile through one exact DrawingML profile. A projected imported
chart may issue `setChartTextStyle`; only that capability authorizes changes to
these fields. Ordinary `setChartTitle` and `setChartData` do not authorize a
style mutation. Theme transforms, shadows and other rich-text effects remain
source-owned, as do legend, data-label and axis-title typography.

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
for bubble. Each vector has exactly the same length as `values`, bubble sizes
are positive, and `categories` is empty. Do not encode numeric X values as
strings merely to reuse a category chart. On imported charts, the current
`setChartData` capability does not authorize changing X values or bubble sizes.

Chart-series gradients, image paint, explicit no-fill series, missing-value
caches, radar, and waterfall still fail closed. Existing unsupported native
chart graphs remain source-preserved; they are not simplified during an
unrelated imported edit.

The authored table compiler owns a physical column/row grid, finite rectangular
merges, one optional header row, row/column banding flags, bounded rich text,
body and paragraph layout, none/solid/gradient cell fills, and direct left,
top, right, and bottom borders. Named table styles provide defaults and inline
style properties override them field by field. More than one header row and
image-filled cells fail closed. Imported table topology and unmodeled native
style graphs remain source-owned.

Use [the complete PPJ reference](ppj.md) for exact fields, value ranges, and
compiler boundaries. This page explains visual choice; it is not a shortened
substitute for the language manual.

Render high-risk pages at final dimensions. Verify axes, labels, legends,
units, source text, value precision, empty cells, merged-looking boundaries,
and whether the visual still communicates without animation.
