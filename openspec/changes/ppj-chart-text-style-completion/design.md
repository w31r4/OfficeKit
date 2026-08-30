# Design

## PPJ surface

```json
{
  "type": "chart",
  "style": {
    "legend": "bottom",
    "legendTextStyle": { "fontFamily": "Aptos", "fontSize": 9 },
    "dataLabels": {
      "showValue": true,
      "textStyle": { "fontSize": 8, "bold": true, "color": "#16324F" }
    }
  },
  "xAxis": {
    "title": "Half-year",
    "titleTextStyle": { "fontFamily": "Georgia", "fontSize": 10 },
    "textStyle": { "fontSize": 8 }
  }
}
```

The existing `chartTextStyle` definition remains the only value contract.
Legend style requires a visible legend, data-label style requires a structured
`dataLabels` object, and axis-title style requires a non-empty axis title.

## Native mapping

- legend: `c:legend/c:txPr`
- plot-level data labels: `c:dLbls/c:txPr`
- axis title: the single canonical `a:r/a:rPr` inside `c:axis/c:title`

The existing canonical DrawingML text reader/writer is generalized rather than
duplicated. Source-bound editing is allowed only when the exact native graph is
recognized again from the source package.

## PPJ lowering

The existing `setChartTextStyle` capability covers the three new locations.
The source-bound compiler accepts changes only to:

- `style.legendTextStyle`
- `style.dataLabels.textStyle`
- `<axis>.titleTextStyle`

All other style and axis properties remain equal during that operation. Combo
charts require identical plot-level data-label style on both native plots.

## Verification

Extend the existing comprehensive PPJ chart sample once. It must build,
re-import, project from a source package, modify all new locations, and project
the requested values again. No per-property matrix or new test fixture is
created.
