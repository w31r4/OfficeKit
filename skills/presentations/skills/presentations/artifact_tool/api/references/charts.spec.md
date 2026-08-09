# Charts

`slide.charts` creates chart elements with slide placement and chart configuration.

## Canonical OfficeKit PPTX Boundary

The model and reference-shaped API describe many chart families, but native
OfficeKit PPTX creation/import/edit is deliberately narrower:

- `bar`, `line`, `pie`, standard `area`, and fixed 50%-hole `doughnut`
  accept literal categories and one finite value per category. Pie and doughnut
  have no axes; `dataLabels.showPercent` is available only for those circular
  families.
- Marker-only `scatter` accepts no shared categories. Every series owns aligned
  finite `xValues` and Y `values`, and the chart uses two numeric value axes.
  Markers may carry fill and border styling; a series line is rejected rather
  than silently changing the chart into a connected scatter plot.
- Bounded 2D `bubble` uses the same numeric X/Y contract plus one aligned,
  finite, positive `bubbleSize` per point. Native size semantics are area-based;
  3D, negative, and custom-scale variants are outside this profile.
- `combo` is native only for clustered bar plus standard line plots. Every
  series declares `chartType: "bar"` or `chartType: "line"`, with at least
  one primary bar and one line. Bars always use the primary category/value
  pair. Lines are either all primary, or every line declares
  `axisGroup: "secondary"` and uses the canonical secondary category/value
  pair at the top and right of the chart.
- These profiles permit title, legend, basic fill/line styling, markers only on
  line and scatter series, chart-level data labels, and bounded primary axes
  where the family has axes. Bar and line series, including combo members,
  additionally permit the bounded trendline and error-bar profiles below.
  Their plot/series/point/trendline-count/error-bar-presence topology is fixed
  after import.
- PPTX charts are literal-data ChartParts. Formula references, external or
  embedded workbooks, stacked area, non-50% doughnut geometry, connected or
  smooth scatter, bubble 3D/negative/custom-scale semantics, mixed
  primary/secondary combo line groups, secondary bars, point overrides,
  per-series data labels, trendline labels/extensions/complex line graphs,
  formula-backed custom error bars without an explicit embedded-workbook route,
  error-bar extensions/complex line graphs, and other chart families fail
  closed or remain source-bound. Do not re-create an irregular imported chart
  from its visible values and claim it was preserved.

Use `inspect` before editing an imported chart, make the smallest supported
change, export to a distinct output, import once more, and render the final
slide for visual QA. The bundled LibreOfficeDev 26.8 alpha currently overlaps
the dual value-axis tick labels for cache-only PPTX secondary-axis combos even
though its series and right-axis title render; retain the OOXML/round-trip gate
and use a Microsoft PowerPoint/native-host lane for release-grade placement QA.

The runnable
`examples/officekit-chart-families-workflow.mjs` authors area, doughnut,
scatter, and bubble charts, inventories their native ChartParts, imports and
edits one semantic field in each family, exports and imports a second time, and
writes a real Playwright PNG plus a source/output-bound audit.

`examples/officekit-chart-trendline-workflow.mjs` performs the same audited
author/import/edit/reimport/render loop for line/combo trendlines and bounded
standard-deviation/custom-literal error bars.

### Bounded bar/line trendlines

Each bar or line series accepts at most 16 trendlines. `type` is one of
`linear`, `exp`, `log`, `power`, `poly`, or `movingAvg`. Polynomial `order` is
2 through 6; moving-average `period` is 2 through
`min(255, values.length - 1)`. `forward` and `backward` are 0 through 1,000,000
in 0.5-category increments. `intercept` must be finite and within JavaScript's
safe-integer magnitude. `displayEquation`, `displayRSquared`, `name`, and the
same simple RGB `line` profile are optional.

```ts
const chart = slide.charts.add("line", {
  categories: ["Q1", "Q2", "Q3", "Q4"],
  series: [{
    name: "Pipeline",
    values: [42, 51, 63, 78],
    trendlines: [{
      type: "linear",
      name: "Pipeline projection",
      forward: 0.5,
      displayEquation: true,
      displayRSquared: true,
      line: { fill: "#7C3AED", width: 1.5, style: "dash" },
    }],
  }],
});
```

Do not add or remove trendlines after import. If a native trendline contains a
label, extension, unknown child, non-RGB/theme color, or complex line graph,
OfficeKit retains the original ChartPart but does not expose a lossy editable
projection.

### Bounded bar/line error bars

Each bar or line series, including a combo member, accepts at most one native
`c:errBars` projection. The reference-compatible shorthand is:

```ts
errorBars: {
  type: "standardError" | "percentage" | "standardDeviation" | "none";
  value?: number;
  endStyle?: "cap" | "noCap";
  line?: LineConfig;
}
```

The compatibility-superset form exposes the underlying bounded semantics:

```ts
errorBars: {
  direction?: "x" | "y";                 // default y
  type?: "both" | "minus" | "plus";     // default both
  valueType?: "fixedVal" | "percentage" | "stdDev" | "stdErr" | "cust";
  value?: number;                         // not valid for stdErr or cust
  noEndCap?: boolean;
  plusValues?: number[];
  minusValues?: number[];
  line?: LineConfig;
}
```

Custom literal arrays must be non-negative and exactly match the series point
count; provide only the sides admitted by `type`. PPTX formula-backed
`plusFormula`/`minusFormula` data is rejected by the native literal ChartPart
path because OfficeKit does not invent an embedded workbook relationship.
Imported error-bar presence is fixed, but its bounded value, cap, side-cache,
and line fields can be edited in place. Duplicate nodes, extensions, unknown
children, malformed caches, or theme/complex line graphs preserve the original
ChartPart and make it read-only rather than being partially projected.

### Canonical secondary-line combo

```ts
slide.charts.add("combo", {
  categories: ["Q1", "Q2", "Q3"],
  series: [
    { name: "Revenue", chartType: "bar", values: [42, 48, 57] },
    { name: "Margin", chartType: "line", axisGroup: "secondary", values: [12, 15, 18] },
  ],
  axes: {
    category: { title: "Quarter" },
    value: { title: "Revenue ($M)" },
    secondary: {
      category: { title: "Quarter" },
      value: { title: "Margin (%)", min: 0, max: 25 },
    },
  },
});
```

Use the secondary form only when every line is secondary and both secondary
axes are present. Otherwise omit `axisGroup` and `axes.secondary` to use the
shared-primary form.

## Resolved From Inspect

```ts
const chart = presentation.resolve("ch/b2c3d4e5");
chart.title = "Updated chart title";
chart.yAxis = { numberFormatCode: "$#,##0M" };
chart.series.getItemAt(0).values = [3.1, 3.7, 4.2, 4.8];
```

Use `presentation.inspect({ kind: "chart", search })` to find the `ch/...`
anchor id. If an imported chart resolves as an image, preserve it as an image or
rebuild it as a native chart intentionally.

## Add Chart

```ts
const chart = slide.charts.add(chartType, {
  position,
  title,
  titleTextStyle,
  categories,
  series,
  hasLegend,
  legend,
  barOptions,
  lineOptions,
  areaOptions,
  pieOptions,
  doughnutOptions,
  treemapOptions,
  mapOptions,
  funnelOptions,
  boxWhiskerOptions,
  histogramOptions,
  view3d,
  scatterOptions,
  xAxis,
  yAxis,
  dataLabels,
  dataTable,
  chartFill,
  chartLine,
  plotAreaFill,
  plotAreaLine,
});
```

Small chart option enums are listed below.

## Chart Inline Types

```ts
type ChartTypeName =
  | "line" | "pie" | "bar" | "doughnut" | "scatter" | "bubble" | "radar"
  | "treemap" | "sunburst" | "map" | "waterfall" | "line3D" | "pie3D"
  | "area3D" | "bar3D" | "funnel" | "histogram" | "boxWhisker" | "stock"
  | "surface3D" | "ofPie" | "surface" | "pareto" | "combo" | "area";

type ChartConfig = {
  position?: { left?: number; top?: number; width?: number; height?: number };
  title?: string;
  titlePlacement?: "none" | "aboveChart" | "centeredOverlay";
  titleTextStyle?: ChartTextStyleConfig;
  categories?: string[];
  series?: ChartSeriesConfig[];
  hasLegend?: boolean;
  legend?: ChartLegendConfig;
  barOptions?: { direction?: "bar" | "column"; grouping?: "clustered" | "stacked" | "percentStacked"; varyColors?: boolean; gapWidth?: number; gapDepth?: number; overlap?: number; bar3dShape?: number };
  lineOptions?: { grouping?: "standard" | "stacked" | "percentStacked"; smooth?: boolean; varyColors?: boolean };
  areaOptions?: { grouping?: "standard" | "stacked" | "percentStacked"; varyColors?: boolean };
  pieOptions?: { firstSliceAngle?: number };
  doughnutOptions?: { holeSize?: number; firstSliceAngle?: number };
  treemapOptions?: { parentLabelLayout?: "none" | "overlapping" | "banner" };
  mapOptions?: { mapArea?: "world" | "auto" | "dataOnly" | "region"; projection?: "mercator" | "auto" | "miller" | "albers"; labelLayout?: "none" | "bestFit" | "showAll"; dataLevel?: "auto" | "county" | "postalCode" | "countryOrRegion" | "stateOrProvince" | "stateCode" | "countyCode" | "countryOrRegionCode"; showUnknown?: boolean; onlyRegionsWithData?: boolean };
  funnelOptions?: { gapWidth?: number };
  boxWhiskerOptions?: { showMeanLine?: boolean; showMeanMarker?: boolean; showNonOutliers?: boolean; showOutliers?: boolean; quartileMethod?: "inclusive" | "exclusive" };
  histogramOptions?: { binWidth?: number; intervalClosed?: number; aggregated?: boolean };
  view3d?: { rotX?: number; rotY?: number; perspective?: number; rightAngleAxes?: boolean };
  scatterOptions?: { style?: "line" | "lineWithMarkers" | "marker" | "smooth" | "smoothWithMarkers"; varyColors?: boolean };
  xAxis?: ChartAxisConfig;
  yAxis?: ChartAxisConfig;
  dataLabels?: ChartDataLabelsConfig;
  dataTable?: ChartDataTableConfig;
  chartFill?: FillConfig;
  chartLine?: LineConfig;
  plotAreaFill?: FillConfig;
  plotAreaLine?: LineConfig;
  displayBlanksAs?: "zero" | "gap" | "span";
  styleIndex?: number;
};
```

## Grouped Edits

```ts
chart.xAxis = axisConfig;
chart.yAxis = axisConfig;
chart.legend = legendConfig;
chart.dataLabels = dataLabelsConfig;
chart.dataTable = dataTableConfig;
```

## Series

```ts
const chart = slide.charts.add(chartType, {
  categories,
  series: [
    {
      name: seriesName,
      categories,
      values,
      xValues,
      categoryPaths,
      fill: fillConfig,
      line: lineConfig,
      marker: markerConfig,
      points,
      dataLabelOverrides,
      trendlines,
      errorBars,
    },
  ],
});
```

## Series Inline Type

```ts
type ChartSeriesConfig = {
  name: string;
  categories?: string[];
  values?: number[];
  xValues?: number[];
  categoryPaths?: string[][];
  bubbleSizes?: number[];
  explosion?: number;
  fill?: FillConfig;
  line?: LineConfig;
  stroke?: LineConfig;
  marker?: { symbol?: "circle" | "diamond" | "dot" | "none" | "plus" | "square" | "star" | "triangle" | "x"; size?: number; fill?: FillConfig; line?: LineConfig };
  points?: Array<{ idx: number; fill?: FillConfig; line?: LineConfig; stroke?: LineConfig }>;
  dataLabelOverrides?: Array<{ idx: number; text?: string; position?: string; fill?: FillConfig; line?: LineConfig; stroke?: LineConfig; showValue?: boolean; showSeriesName?: boolean; showCategoryName?: boolean; showPercent?: boolean; textStyle?: ChartTextStyleConfig }>;
  trendlines?: Array<{
    type: "linear" | "exp" | "log" | "power" | "poly" | "movingAvg";
    name?: string;
    order?: number;
    period?: number;
    forward?: number;
    backward?: number;
    intercept?: number;
    displayEquation?: boolean;
    displayRSquared?: boolean;
    line?: LineConfig;
  }>;
  errorBars?: {
    direction?: "x" | "y";
    type?: "both" | "minus" | "plus" | "standardError" | "percentage" | "standardDeviation" | "none";
    errorBarType?: "both" | "minus" | "plus";
    valueType?: "fixedVal" | "percentage" | "stdDev" | "stdErr" | "cust";
    value?: number;
    noEndCap?: boolean;
    endStyle?: "cap" | "noCap";
    plusValues?: number[];
    minusValues?: number[];
    line?: LineConfig;
  };
  valuesFormatCode?: string;
  xValuesFormatCode?: string;
};
```

## Data Label Overrides

```ts
const override = chart.series
  .getItemAt(seriesIndex)
  .dataLabelOverrides.add(dataPointIdx);
override.text = labelText;
override.position = labelPosition;
override.textStyle.fontSize = fontSizePx;
override.textStyle.fill = textFill;
override.fill = fillConfig;
override.stroke = lineConfig;
```

## Chart Areas

```ts
const chart = slide.charts.add(chartType, {
  chartFill,
  plotAreaFill,
});
```

## Axis, Legend, Label Inline Types

```ts
type ChartTextStyleConfig = {
  fontSize?: number;
  fill?: FillConfig;
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  alignment?: "left" | "center" | "right" | "justify";
};

type ChartLegendConfig = {
  position?: "left" | "top" | "topRight" | "right" | "bottom";
  overlay?: boolean;
  fill?: FillConfig;
  line?: LineConfig;
  textStyle?: ChartTextStyleConfig;
};

type ChartDataLabelsConfig = {
  position?: "center" | "inEnd" | "outEnd";
  showValue?: boolean;
  showSeriesName?: boolean;
  showCategoryName?: boolean;
  showPercent?: boolean;
  showLeaderLines?: boolean;
  textStyle?: ChartTextStyleConfig;
  fill?: FillConfig;
  line?: LineConfig;
};

type ChartAxisConfig = {
  visible?: boolean;
  title?: string | { text?: string; textStyle?: ChartTextStyleConfig };
  numberFormatCode?: string;
  min?: number;
  max?: number;
  majorUnit?: number;
  minorUnit?: number;
  position?: "bottom" | "left" | "right" | "top";
  tickLabelPosition?: "nextTo" | "high" | "low" | "none" | string;
  textStyle?: ChartTextStyleConfig;
  line?: LineConfig;
  majorGridlines?: LineConfig | null;
  minorGridlines?: LineConfig | null;
};
```

## Cookbook

```ts
// Executive horizontal bar chart.
slide.charts.add("bar", {
  position: { left: 96, top: 160, width: 720, height: 360 },
  categories: ["Enterprise", "Mid-market", "SMB"],
  series: [{ name: "ARR", values: [42, 28, 17], fill: "#2563eb" }],
  barOptions: { direction: "bar", grouping: "clustered", gapWidth: 44 },
  hasLegend: false,
  xAxis: { visible: false, majorGridlines: null },
  yAxis: { textStyle: { fill: "#475569", fontSize: 13 }, line: { style: "solid", fill: "#e2e8f0", width: 1 } },
  dataLabels: { showValue: true, position: "outEnd", textStyle: { fill: "#0f172a", fontSize: 13, bold: true } },
});
```

```ts
// Compact trend line with muted grid.
slide.charts.add("line", {
  position: { left: 96, top: 150, width: 880, height: 280 },
  categories: ["Jan", "Feb", "Mar", "Apr"],
  series: [{ name: "Conversion", values: [31, 34, 37, 43], line: { style: "solid", fill: "#0f766e", width: 3 } }],
  legend: { position: "bottom", overlay: false },
  yAxis: { numberFormatCode: "0%", majorGridlines: { style: "solid", fill: "#e2e8f0", width: 1 } },
});
```

```ts
// Doughnut chart with labels outside.
slide.charts.add("doughnut", {
  categories: ["Product", "Sales", "Support"],
  series: [{ name: "Share", values: [52, 31, 17] }],
  dataLabels: { showPercent: true, showCategoryName: true, position: "outEnd" },
  legend: { position: "right" },
});
```
