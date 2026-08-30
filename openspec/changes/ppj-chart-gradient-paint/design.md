# Design

## PPJ surface

```json
{
  "type": "chart",
  "style": {
    "chartAreaFill": { "type": "none" },
    "plotAreaFill": {
      "type": "gradient",
      "kind": "linear",
      "angle": 90,
      "stops": [
        { "offset": 0, "color": "#F8FAFC" },
        { "offset": 1, "color": "#DCEFEA", "opacity": 0.6 }
      ]
    }
  },
  "data": {
    "categories": ["Q1", "Q2"],
    "series": [{
      "id": "revenue",
      "name": "Revenue",
      "values": [12, 18],
      "fill": {
        "type": "gradient",
        "kind": "linear",
        "angle": 0,
        "stops": [
          { "offset": 0, "color": "#0B8F8F" },
          { "offset": 1, "color": "#F2C14E" }
        ]
      }
    }]
  }
}
```

## Native mapping

The shared `SpreadsheetChartSurfaceFill` owns one direct `a:noFill`, direct
sRGB `a:solidFill`, or the existing literal `a:gradFill` profile. Series carry
an additive fill message while retaining the old RGB field as a read/write
compatibility alias.

## Source-bound lowering

Canonical imported charts issue `setChartFill`. It covers only
`style.chartAreaFill`, `style.plotAreaFill`, and `data.series[].fill`.
Categories, values, line/marker state and every other style field remain equal.
The writer re-proves the exact native fill graph before replacement.

## Verification

Extend the existing integrated PPJ chart contract once. It builds, imports,
projects, changes all three fill locations, rebuilds from the source package,
and projects the values again. No fill matrix or new fixture is added.
