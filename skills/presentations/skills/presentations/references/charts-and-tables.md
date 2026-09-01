# Charts and tables

Choose the visual from the relationship the audience must understand.

| Relationship | Prefer |
| --- | --- |
| trend or change over ordered time | line, area, or ordered columns |
| changing composition and total magnitude over ordered time | streamgraph |
| magnitude across categories | bars or columns with a common baseline |
| small exact counts where one repeated glyph has a declared unit | pictographic bar or column |
| distribution or correlation | scatter, bubble, or distribution view |
| intensity across two categorical dimensions | heatmap |
| open/high/low/close movement over ordered periods | candlestick |
| hierarchical part-to-whole allocation | treemap |
| hierarchical part-to-whole path across concentric levels | sunburst |
| conserved magnitude moving through a directed process | sankey |
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

Treat native data-label defaults as unsafe. Configure `showValue`,
`showCategory`, `showSeries`, and `showPercent` explicitly at the smallest
scope that communicates the claim. Keep a series name in one legend or
heading, not repeated beside every mark. For endpoint-only evidence, use two
endpoint marks or a visibly two-point comparison; never let a renderer bridge
missing categories into an invented continuous trend. Render the chart at
final size and remove any label collision before delivery, even when structural
review passes.

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
  bounded semantic waterfall, bounded vector heatmap, bounded vector
  candlestick, bounded vector streamgraph, bounded vector pictographic bars and
  columns, bounded vector treemap, bounded vector sunburst, bounded vector
  sankey, bounded column-line-area category combo plots, and bounded editable
  scatter/bubble plus line/area/column numeric overlays;
- legend visibility and top, bottom, left, or right placement;
- ordinary, stacked, and percent-stacked grouping where the chart family
  supports it;
- pie/doughnut first-slice orientation from 0 through 360 degrees and
  doughnut center-hole size from 10 through 90 percent;
- bubble scale from 0 through 300 percent and native size interpretation by
  bubble area or width;
- bar direction, gap width, category/value-axis visibility, major gridlines,
  data-label visibility and bounded label position;
- presence-aware axis reversal plus direct RGB/no-fill axis and major-grid
  lines with bounded width, dash, opacity, cap, and join;
- chart-area, plot-area, and series none, solid, or bounded direct-RGB gradient
  fills, including solid and per-stop opacity;
- editable series line width, dash, opacity, cap, join, and bounded markers;
- category/value-axis titles, number formats, label interval, tick-label
  typeface, size, bold, italic and direct RGB/alpha, value bounds and major
  unit; bounded combo charts may declare the matching secondary pair;
- chart-title Latin/East Asian typeface, size, bold, italic and direct
  RGB/alpha plus canonical line-chart smoothing and direct color variation;
- structured data labels for value, category, series, percentage, native
  position, direct number format and bounded typography at plot, series and
  sparse point scope;
- sparse point fill, outline and pie/doughnut explosion for measured exceptions;
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
  "xAxis": {
    "title": "Quarter",
    "tickLabelInterval": 1,
    "axisLine": { "color": "#16697A", "width": 1 }
  },
  "yAxis": {
    "title": "Conversion rate",
    "numberFormat": "0.0%",
    "min": 0,
    "max": 0.4,
    "majorUnit": 0.1,
    "gridLine": { "color": "#D8E1E8", "width": 0.75, "dash": "dot" }
  },
  "style": {
    "dataLabels": {
      "showValue": true,
      "position": "outside-end",
      "numberFormat": "0.0%"
    }
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

Use series defaults when one series needs a different label policy, and use
sparse point overrides only for evidence that needs an exception. Point
indices are zero-based, strictly increasing and must address an existing
non-missing point:

```json
{
  "data": {
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": [{
      "id": "conversion",
      "name": "Conversion",
      "values": [0.18, 0.22, 0.27, 0.31],
      "dataLabels": {
        "showValue": true,
        "numberFormat": "0.0%",
        "textStyle": { "fontSize": 8, "color": "#334155" },
        "points": [
          { "index": 0, "showValue": false },
          { "index": 3, "position": "top", "textStyle": { "bold": true, "color": "#D9A21B" } }
        ]
      }
    }]
  }
}
```

Do not use point labels as a substitute for a readable scale. Custom label
text, manual label layout, label shapes/effects, leader-line graphs and
source-linked number formats remain source-owned and fail closed.

Use `pointStyles` when one real datum needs visual emphasis without rebuilding
the chart as shapes. Keep the list sparse and sorted; a missing observation
cannot be styled. `explosion` is native only on pie and doughnut:

```json
{
  "data": {
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": [{
      "id": "conversion",
      "name": "Conversion",
      "values": [0.18, 0.22, 0.27, 0.31],
      "pointStyles": [{
        "index": 3,
        "fill": { "type": "solid", "color": "#D9A21B" },
        "stroke": { "color": "#16324F", "width": 1.25 }
      }]
    }]
  }
}
```

Imported point styling requires `setChartFill`. Marker overrides, picture
options, 3D state, effects, extensions and irregular point graphs stay
source-owned rather than being approximated.

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

For a circular chart, rotate the first slice only when it improves the reading
path or aligns a declared slice with a label or adjacent explanation. Size a
doughnut hole to reserve a meaningful center annotation; do not rotate slices
or enlarge the hole merely to make the chart look different:

```json
{
  "chartType": "doughnut",
  "style": {
    "legend": "right",
    "startAngle": 135,
    "holeSize": 68,
    "dataLabels": { "showPercent": true, "position": "center" }
  }
}
```

`startAngle` applies only to `pie` and `doughnut`; `holeSize` applies only to
`doughnut`. Imported editable circular charts expose `setChartPlot`. If that
capability is absent, retain the native plot instead of replacing it with
shapes or an image.

When the title needs an emphasized result, keep the title native and use the
same bounded paragraph/run vocabulary as ordinary text:

```json
{
  "title": {
    "paragraphs": [{
      "runs": [
        { "text": "Measured profile: " },
        {
          "text": "−42% incidents",
          "style": {
            "bold": true,
            "color": "#A83232",
            "fontFamilyEastAsia": "Noto Serif CJK SC"
          }
        }
      ]
    }]
  },
  "style": {
    "titleTextStyle": {
      "fontSize": 14,
      "fontFamily": "Aptos",
      "fontFamilyEastAsia": "Noto Sans CJK SC",
      "color": "#16324F"
    }
  }
}
```

`titleTextStyle` is a default for properties omitted by a structured run;
explicit run typography wins. Formula-backed titles, title hyperlinks,
WordArt, effects, and unknown title containers remain source-owned.

`smooth` preserves an explicit true or false native value. `varyColors: true`
authors one direct native color-variation flag; false is canonical omission.
Both line-behavior fields are line-chart-only. Structured chart titles compile
to bounded native DrawingML paragraphs and runs on ordinary and combo charts.
A projected imported chart may issue `setChartTitle` for that title content and
`setChartTextStyle` for uniform title defaults, legend, data labels, axis titles
and tick labels. Each capability authorizes only its declared field. Theme
transforms and other effect-bearing chart text remain source-owned and fail
closed instead of being flattened.

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
  "style": {
    "bubbleScale": 145,
    "bubbleSizeMode": "area"
  },
  "xAxis": {
    "reverse": false,
    "axisLine": { "color": "#355C7D", "width": 1 },
    "axisLineArrow": { "end": "triangle" }
  },
  "yAxis": {
    "gridLine": { "color": "#D7DEE5", "width": 0.75, "dash": "dot" }
  },
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
to reuse a category chart. `bubbleScale` changes the native chart-wide scale,
not each datum's meaning; `bubbleSizeMode` chooses whether the native size
values represent area or width.

When exact visible size is part of the evidence, declare a bounded mapping:

```json
"style": {
  "bubbleSizeScale": "log",
  "bubbleRadiusRange": [5, 24]
}
```

`bubbleSizeScale` is `sqrt`, `linear`, or `log`; the radius pair is measured in
points and must increase within 2–72. OfficeKit uses one shared size domain for
all bubble series. Declaring either field intentionally compiles the chart as
editable DrawingML ellipses and axes, because native ChartML cannot promise
exact visible radii or logarithmic sizing. Whole-object animation remains
available; native chart-build animation does not. Without the embedded PPJ,
reimport returns the honest editable group and does not infer data semantics.

On imported charts, `setChartData` owns only the bounded data vectors and does
not authorize changing X values, bubble sizes, or the positions of missing Y
observations. Bubble scale and size semantics require `setChartPlot`; label
format requires `setChartLabels`; axis direction and line styling require
`setChartAxis`. `axisLineArrow` accepts independently bounded `start` and `end`
values: `none`, `triangle`, `stealth`, `diamond`, `oval`, or `open`. Use an end
arrow only when direction is part of the analytical meaning; do not add arrows
as decoration, to grid lines, or to radar spokes. When the matching capability
is absent, preserve the native graph rather than redrawing it with shapes.
Per-point label formats, logarithmic transforms, custom arrow sizing,
theme/effect line graphs, and other irregular native formatting remain
source-owned. `bubbleSizeScale` and `bubbleRadiusRange` are authored vector
semantics, not permission to rewrite an imported native ChartPart.

Use `combo` only when two or three different plot families share one real
ordered category domain. Each series declares `chartType: "column" | "line" | "area"`
and `axis: "primary" | "secondary"`. At least two distinct plot families are
required, one family must remain primary, and all series of the same family use
the same axis pair:

```json
{
  "type": "chart",
  "id": "volume-margin-band",
  "chartType": "combo",
  "frame": { "x": 72, "y": 112, "width": 640, "height": 300 },
  "title": "Volume grew while margin stayed inside plan",
  "xAxis": { "title": "Quarter" },
  "yAxis": { "title": "Units", "min": 0 },
  "secondaryXAxis": { "visible": false },
  "secondaryYAxis": { "title": "Margin (%)", "min": 0, "max": 40 },
  "style": { "legend": "bottom", "gapWidth": 80 },
  "data": {
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": [
      {
        "id": "volume",
        "name": "Volume",
        "values": [120, 138, 151, 172],
        "chartType": "column",
        "axis": "primary",
        "fill": { "type": "solid", "color": "#CBD5E1" }
      },
      {
        "id": "plan-band",
        "name": "Plan band",
        "values": [24, 26, 29, 31],
        "chartType": "area",
        "axis": "secondary",
        "fill": { "type": "solid", "color": "#0B8F8F", "opacity": 0.18 }
      },
      {
        "id": "margin",
        "name": "Margin",
        "values": [22, 27, 28, 33],
        "chartType": "line",
        "axis": "secondary",
        "stroke": { "color": "#0B8F8F", "width": 2.5 },
        "marker": "circle"
      }
    ]
  }
}
```

OfficeKit emits separate native `c:areaChart`, `c:barChart` and `c:lineChart`
plots in one editable ChartPart. Area is written behind columns and the line is
written last so its evidence stroke remains visible. A secondary pair is
optional, but must be complete and must serve at least one whole plot family.
Horizontal bars, splitting one family across both axis pairs and a one-family
pseudo-combo are rejected.

When one combo series is `scatter` or `bubble`, the whole element deliberately
switches to the bounded numeric profile. Every series then supplies explicit,
strictly increasing `xValues`, shared categories stay empty, and line, area or
column evidence uses the same real value/value scale:

```json
{
  "type": "chart",
  "id": "adoption-response",
  "chartType": "combo",
  "frame": { "x": 72, "y": 112, "width": 640, "height": 300 },
  "title": "Observed adoption follows the fitted response",
  "xAxis": { "title": "Exposure", "numberFormat": "0.0" },
  "yAxis": { "title": "Adoption", "numberFormat": "0" },
  "style": { "legend": "right", "bubbleScale": 90 },
  "data": {
    "categories": [],
    "series": [
      {
        "id": "observed",
        "name": "Observed",
        "chartType": "bubble",
        "xValues": [1, 2, 3, 4],
        "values": [18, 31, 47, 66],
        "bubbleSizes": [8, 14, 20, 12],
        "color": "#0B8F8FCC"
      },
      {
        "id": "fitted",
        "name": "Fitted",
        "chartType": "line",
        "xValues": [1, 2, 3, 4],
        "values": [20, 32, 46, 64],
        "stroke": { "color": "#16324F", "width": 1.5 }
      }
    ]
  }
}
```

This is an editable DrawingML group, not a disguised categorical ChartPart.
Filled area and column marks sit behind lines and points. The profile accepts
2–8 series and 2–64 complete points per series, one shared axis pair, bounded
axis formatting, and `none` or `right` legend placement. Area or column
overlays require a truthful zero baseline. `bubbleScale` is 10–300. An explicit
`bubbleSizeScale` or `bubbleRadiusRange` uses the same shared domain
and exact editable ellipse mapping described above. Line and scatter series may
use bounded markers, but scatter cannot choose `none`;
bubble, area and column marker settings are rejected instead of ignored.
Secondary axes, formulas,
trendlines, error bars and `chartBuild` animation fail closed. Without the
embedded PPJ, import returns the editable group and does not guess chart data.

Use a streamgraph when the audience needs to see both changing composition and
the changing total across an ordered domain. Use an ordinary line or area chart
when exact point lookup, a common zero baseline, or independent series trends
matter more. Do not use a streamgraph merely because curved bands look richer.

```json
{
  "type": "chart",
  "id": "audience-composition",
  "chartType": "area",
  "frame": { "x": 72, "y": 112, "width": 720, "height": 300 },
  "title": "Audience composition changed without losing reach",
  "xAxis": {
    "visible": true,
    "textStyle": { "fontSize": 8, "color": "#52606D" }
  },
  "style": {
    "stacking": "stream",
    "legend": "right",
    "titleTextStyle": { "fontSize": 14, "bold": true },
    "legendTextStyle": { "fontSize": 8 }
  },
  "data": {
    "categories": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
    "series": [
      {
        "id": "new",
        "name": "New",
        "values": [22, 28, 35, 31, 38, 44],
        "fill": {
          "type": "gradient",
          "kind": "linear",
          "angle": 0,
          "stops": [
            { "offset": 0, "color": "#0B8F8F" },
            { "offset": 1, "color": "#74C7C7", "opacity": 0.82 }
          ]
        }
      },
      { "id": "returning", "name": "Returning", "values": [31, 34, 30, 39, 42, 46], "color": "#F2C14ECC" },
      { "id": "enterprise", "name": "Enterprise", "values": [12, 14, 18, 22, 29, 36], "color": "#C8644AE6" }
    ]
  }
}
```

`stacking: "stream"` is limited to area charts with 2–12 complete,
non-negative series over 3–64 unique ordered categories. Every category needs
a positive total. OfficeKit centers each category's raw total around the plot
midline, so band thickness retains the actual magnitudes; it does not silently
normalize every period to 100 percent. The compiler writes one editable cubic
DrawingML path per series plus ordinary title, category, and legend text.

The embedded PPJ restores the exact semantic chart. If that program is removed,
PPTX import returns the truthful editable group instead of guessing a
streamgraph from arbitrary paths. Whole-object animation is valid; native
`chartBuild`, secondary axes, markers, trendlines, error bars,
missing/negative values, and native point-label overrides are not. The latter
belong to real native ChartParts, not to this generated streamgraph group.

Use a pictographic bar only when a small whole count becomes easier to grasp by
repeating one meaningful symbol. State the conversion explicitly. Do not use
icons as decorative replacements for a precise common-baseline bar chart, and
do not suggest a fractional person, site, device, or event.

```json
{
  "type": "chart",
  "id": "verified-participants",
  "chartType": "bar",
  "frame": { "x": 72, "y": 112, "width": 720, "height": 280 },
  "title": "Verified participants by cohort",
  "data": {
    "categories": ["Control", "Pilot", "Follow-up"],
    "series": [
      {
        "id": "participants",
        "name": "Participants",
        "values": [30, 50, 20],
        "color": "#0B8F8F",
        "symbol": {
          "kind": "icon",
          "iconName": "fas:user",
          "unit": 10,
          "gap": 2,
          "showValue": true,
          "unitLabel": "participants"
        }
      }
    ]
  }
}
```

The bounded profile accepts one bar or column series, 2–12 unique string
categories, complete non-negative values, at most 32 symbols per category and
192 total. Every value must divide exactly by `unit`; OfficeKit rejects rather
than clips or rounds a symbol. `kind: "preset"` with an existing DrawingML
preset such as `star5` is the alternative to a pinned offline `iconName`.
Series `color`, solid/gradient `fill`, and optional `stroke` control the units.

Each unit, category label, value label, title and unit statement remains an
editable native object with a stable child ID. Embedded PPJ restores the exact
chart. Without it, import returns the honest editable group and never guesses
counts from repeated shapes. Whole-object animation is valid; native
`chartBuild`, multiple series, axes, fractional symbols, markers, trendlines,
error bars and other ChartPart-only options fail closed.

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

Use `candlestick` when the body and wick encode real ordered open/high/low/close
observations. Do not use it as a generic trend line or infer missing OHLC values
from a close-only series. `values` carries close, `highValues` and `lowValues`
are required, and `openValues` distinguishes OHLC from HLC:

```json
{
  "type": "chart",
  "id": "daily-price-range",
  "chartType": "candlestick",
  "frame": { "x": 72, "y": 112, "width": 640, "height": 300 },
  "title": "Daily OHLC",
  "xAxis": { "title": "Session", "tickLabelInterval": 1 },
  "yAxis": { "title": "USD", "numberFormat": "0.0", "min": 88, "max": 120, "majorUnit": 8 },
  "style": {
    "candlestick": {
      "up": {
        "fill": { "type": "solid", "color": "#0B8F8F" },
        "stroke": { "color": "#086E6E", "width": 0.6 }
      },
      "down": {
        "fill": { "type": "solid", "color": "#C8644A" },
        "stroke": { "color": "#8B3E2F", "width": 0.6 }
      },
      "wick": { "color": "#16324F", "width": 0.8, "cap": "round" },
      "bodyWidthRatio": 0.55,
      "showCloseValues": false,
      "gridlineStroke": { "color": "#CBD5E1", "width": 0.5 },
      "axisTextStyle": { "fontSize": 8, "color": "#52606D" }
    }
  },
  "data": {
    "categories": ["D1", "D2", "D3", "D4"],
    "series": [
      {
        "id": "price",
        "name": "Price",
        "openValues": [92, 96, 94, 101],
        "highValues": [98, 99, 103, 104],
        "lowValues": [90, 91, 92, 96],
        "values": [96, 94, 101, 99]
      },
      {
        "id": "moving-average",
        "name": "Moving average",
        "chartType": "line",
        "values": [94, 95, 97, 99],
        "stroke": { "color": "#F2C14E", "width": 1.4 },
        "marker": { "symbol": "circle", "size": 4, "fill": "#F2C14E" }
      }
    ]
  }
}
```

The profile accepts one OHLC/HLC body plus up to four aligned `line`, `area` or
`column` overlays across 1–64 unique ordered string categories. Every open and
close must lie inside its low/high interval. Omit `openValues` for HLC;
OfficeKit then draws an editable close tick instead of inventing a body. Every
overlay value is explicit and participates in the same Y domain. Filled area
and column overlays render before wicks and bodies; line overlays render after
the bodies so a moving average remains visible.
`showCloseValues` is limited to 16 observations, and the frame must leave enough
width for native marks and labels. Bounded axis number formats are `0`, `0.0`,
`0.00`, `#,##0`, `#,##0.0`, and `#,##0.00`.

PowerPoint has no portable authored candlestick ChartPart in this compiler
profile. OfficeKit lowers the semantic node to one editable DrawingML group of
native wick connectors, body shapes, axes, gridlines, and text. Embedded PPJ
restores exact OHLC/HLC intent; without it, import truthfully returns the group
and does not reverse-engineer arbitrary marks into financial evidence.
Whole-object animation remains available, while `chartBuild`, secondary axes,
legends, trendlines, error bars, numeric-X overlays and automatic data labels
fail closed. Overlay line markers use the ordinary bounded marker vocabulary;
area and column marker settings are rejected instead of ignored.

Use `treemap` for hierarchical part-to-whole evidence when area is the intended
comparison. It is not a substitute for a ranked bar chart when exact ordering
or small differences matter. `categories`, `values`, and `parents` are aligned;
each parent name references one globally unique category, and `null` marks a
root:

```json
{
  "type": "chart",
  "id": "budget-allocation",
  "chartType": "treemap",
  "frame": { "x": 72, "y": 112, "width": 640, "height": 300 },
  "title": "Budget allocation",
  "style": {
    "treemap": {
      "rootColors": ["#0B8F8F", "#C8644A", "#F2C14E"],
      "border": { "color": "#FFFFFF", "width": 0.75 },
      "gap": 2,
      "headerHeight": 17,
      "depthLighten": 0.1,
      "showValues": true,
      "labelTextStyle": { "fontSize": 9, "bold": true },
      "valueTextStyle": { "fontSize": 8 }
    }
  },
  "data": {
    "categories": ["Engineering", "Frontend", "Backend", "Sales", "Enterprise", "SMB"],
    "series": [{
      "id": "budget",
      "name": "Budget",
      "levels": 2,
      "values": [1000, 400, 600, 800, 500, 300],
      "parents": [null, "Engineering", "Engineering", null, "Sales", "Sales"]
    }]
  }
}
```

The bounded profile accepts one series, 1–128 positive nodes, 1–16 roots, and
at most eight hierarchy levels. Every named parent must exist, parent chains
must be acyclic, and each non-leaf value must equal the sum of its direct
children. Those checks prevent a visually plausible rectangle mosaic from
silently contradicting its evidence.

Set the series `levels` to show only the first 1–8 levels without deleting
deeper nodes from PPJ. A node at the last visible level receives its branch's
full rectangle and becomes the visible summary leaf. Omit `levels` to show the
complete available hierarchy.

NativeAOT uses a deterministic squarified layout and emits one editable
DrawingML group of rectangles and text. Root colors cycle in root order;
descendants retain the root hue and lighten by the declared depth step. Labels
are omitted only when a native rectangle is too small to hold readable text;
the node remains an editable named shape and the group retains accessibility
evidence. Embedded PPJ restores the exact forest; without it, import truthfully
returns the ordinary group. Whole-object animation is supported, but
`chartBuild`, axes, legends, trendlines, error bars, and arbitrary per-node
expression-driven paint are not.

Use `sunburst` when the audience must follow a hierarchy from a root through
concentric levels while retaining part-to-whole area. Prefer `treemap` for
denser label comparison and bars for precise rank. Sunburst uses the same
aligned `categories`, positive `values`, and nullable `parents` channel as
treemap:

```json
{
  "type": "chart",
  "id": "portfolio-contribution",
  "chartType": "sunburst",
  "frame": { "x": 72, "y": 96, "width": 640, "height": 340 },
  "title": "Contribution by portfolio",
  "style": {
    "sunburst": {
      "rootColors": ["#0B8F8F", "#C8644A"],
      "border": { "color": "#FFFFFF", "width": 0.6 },
      "innerRadiusRatio": 0.18,
      "ringGap": 1.5,
      "segmentGapDegrees": 1,
      "startAngle": -90,
      "clockwise": true,
      "depthLighten": 0.1,
      "showValues": true,
      "labelTextStyle": { "fontSize": 8, "bold": true },
      "valueTextStyle": { "fontSize": 7 }
    }
  },
  "data": {
    "categories": ["Company", "Product", "Operations", "Platform", "Apps", "Delivery", "Support"],
    "series": [{
      "id": "portfolio",
      "name": "Contribution",
      "levels": 2,
      "values": [100, 55, 45, 30, 25, 20, 25],
      "parents": [null, "Company", "Company", "Product", "Product", "Operations", "Operations"]
    }]
  }
}
```

The bounded profile accepts one series, 1–96 nodes, 1–16 roots, and at most six
levels. Every parent must exist, the forest must be acyclic, and each non-leaf
value must equal its direct-child sum. Root values allocate root angles in
declared order; children partition their parent angle, and depth selects the
ring. Segment and ring gaps are visual state, never hidden data changes.
Set series `levels` to 1–6 when the audience should see only the leading
hierarchy. The compiler divides the available radius among visible levels, so
the remaining rings become wider rather than leaving blank space. The full
forest and its totals remain in PPJ; omit the field to show every level.

NativeAOT emits one editable DrawingML group. Every annular sector is a named
custom-geometry shape whose circular edges use bounded cubic paths; labels are
ordinary text boxes and are omitted when the measured sector cannot hold them.
No PNG or ChartPart is introduced. Embedded PPJ restores the exact hierarchy;
without it, import exposes the ordinary custom-shape group and does not infer
sunburst semantics from arbitrary arcs. Whole-object animation is supported;
`chartBuild`, axes, legends, trendlines, error bars, and expression-driven
per-node paint fail closed.

Use `sankey` when ribbon thickness must communicate conserved magnitude moving
through a directed process. Use an ordinary process diagram when sequence or
responsibility matters but edge width has no quantitative meaning. Declare a
stable node catalog in `categories`; the one series carries aligned positive
`values`, `sources`, and `targets`:

```json
{
  "type": "chart",
  "id": "conversion-flow",
  "chartType": "sankey",
  "frame": { "x": 50, "y": 96, "width": 860, "height": 330 },
  "title": "Lead conversion flow",
  "style": {
    "sankey": {
      "nodeColors": ["#16324F", "#0B8F8F", "#F2C14E", "#C8644A"],
      "nodeStroke": { "color": "#FFFFFF", "width": 0.5 },
      "nodeWidth": 14,
      "nodeGap": 10,
      "nodeAlign": "right",
      "nodeColorMap": { "Paid": "#C1121F" },
      "flowOpacity": 0.42,
      "flowCurvature": 0.72,
      "flowColorMode": "source",
      "showValues": true,
      "labelTextStyle": { "fontSize": 8, "bold": true },
      "valueTextStyle": { "fontSize": 7, "color": "#52606D" }
    }
  },
  "data": {
    "categories": ["Leads", "Qualified", "Trial", "Nurture", "Paid", "Churn"],
    "series": [{
      "id": "accounts",
      "name": "Accounts",
      "values": [100, 60, 40, 45, 15, 25, 15],
      "sources": ["Leads", "Qualified", "Qualified", "Trial", "Trial", "Nurture", "Nurture"],
      "targets": ["Qualified", "Trial", "Nurture", "Paid", "Churn", "Paid", "Churn"]
    }]
  }
}
```

The bounded profile accepts 2–64 unique nodes and 1–256 unique directed edges.
Every endpoint must be declared, every node must participate, flows must be
positive, and the graph must be acyclic. Any node with both incoming and
outgoing edges must conserve flow. Combine repeated endpoint pairs explicitly;
do not rely on a renderer to merge them silently.

NativeAOT assigns stable topological columns, scales node/ribbon thickness from
the same values, and stacks each ribbon consistently at both endpoints. Flows
are closed cubic custom shapes behind native node rectangles and ordinary text.
No PNG or ChartPart is introduced. Embedded PPJ restores exact graph semantics;
without it, import returns the ordinary editable group and does not infer a
Sankey from arbitrary ribbons. Whole-object animation is supported;
`chartBuild`, cycles, negative flows, non-conserving internal nodes, arbitrary
graph constraints, and expression-driven paint fail closed.

`nodeAlign: "left"` keeps the earliest topological placement. `"justify"`
pushes sinks to the last column. `"right"` aligns each node by its longest
remaining route to a sink, which is useful when short branches should end near
the outcome rather than near the source. `nodeColorMap` overrides the palette
by exact declared node name; undeclared names are rejected instead of silently
creating a second category.

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
  "spokeAxis": {
    "show": true,
    "min": 0,
    "max": 100,
    "majorUnit": 20,
    "label": {
      "numberFormat": "0",
      "fontSize": 8.5,
      "color": "#475569"
    },
    "axisLine": { "color": "#CBD5E1", "width": 0.75 },
    "gridLine": { "color": "#E2E8F0", "width": 0.5, "dash": "dot" }
  }
}
```

PPJ compiles this to editable native `standard` radar rather than drawing a
polygon from shapes. Category/value axes, series stroke, bounded markers,
legend, labels, chart surfaces and fixed-topology source continuation use the
same chart contracts as the other category families. Filled, marker-only, 3D,
extension-bearing and irregular native radar variants remain source-owned.

`spokeAxis` is the preferred radar interface. It describes one coordinate
system instead of exposing the underlying category/value axis pair: `axisLine`
styles the radial spokes, while `gridLine` styles the concentric rings. Set
`label: false` to hide numeric scale labels without hiding the rings. Do not
combine `spokeAxis` with `xAxis`, `yAxis`, secondary axes, or the legacy
`style.showGridlines` fields. Canonical imported standard radar charts project
the same object and may continue through capability-issued local edits;
custom label positions, logarithmic scales, arrows, filled/3D variants and
irregular native axis graphs remain source-owned.

Chart image/pattern paint, theme-transformed gradients, irregular sparse
caches, and unrecognized waterfall/ChartEx graphs still fail closed. Existing
unsupported native chart graphs remain source-preserved; they are not
simplified during an unrelated imported edit.

The authored table compiler owns a physical column/row grid, finite rectangular
merges, zero or more bounded header rows, row/column banding flags, bounded rich
text, body and paragraph layout, none/solid/gradient/image cell fills, and
direct left, top, right, and bottom borders. An image fill uses the same local
hashed asset, crop, cover/contain/stretch/tile, and opacity contract as other
PPJ image paint. Use it when the cell itself is an evidence thumbnail, product
identity, or comparison image; do not turn ordinary data tables into decorative
mosaics.

For a two-level analytical header, declare the semantic count and shared
fallbacks once:

```json
{
  "style": {
    "headerRows": 2,
    "headerCellFill": { "type": "solid", "color": "#E8EEF3" },
    "headerTextStyle": {
      "verticalAlignment": "middle",
      "defaultText": { "font": "sans", "size": 10, "bold": true, "color": "#16324F" }
    },
    "defaultCellFill": { "type": "none" }
  }
}
```

Cell-local `fill` and `textStyle` win over the header fallback; header styling
wins over the ordinary table defaults. `headerRows` cannot exceed the physical
row count, and header-only styling without a header row is rejected. OfficeKit
writes direct editable cell formatting for every declared header row and the
ordinary native first-row flag. Exact counts above one survive through the
embedded PPJ; a third-party import without that program conservatively reports
only the native first-row fact. Imported table topology and unmodeled native
style graphs remain source-owned.

For compact repeated styling, `cellStyle` supplies the base; `bodyStyles`
cycles through rows between the first and last row; `firstRowStyle`,
`lastRowStyle`, `firstColumnStyle`, and `lastColumnStyle` add structural roles.
Each style can set `fill`, `textStyle`, and individual `borders`. Explicit cell
properties always win. When row and column roles set the same property,
`rowOverColumn` chooses the winner and defaults to `true`. This is bounded style
inheritance, not a selector language: expand unusual exceptions directly on
the affected cell.

Use [the complete PPJ reference](ppj.md) for exact fields, value ranges, and
compiler boundaries. This page explains visual choice; it is not a shortened
substitute for the language manual.

Render high-risk pages at final dimensions. Verify axes, labels, legends,
units, source text, value precision, empty cells, merged-looking boundaries,
and whether the visual still communicates without animation.
