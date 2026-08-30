# Design

## Public program

```json
{
  "type": "chart",
  "id": "risk-profile",
  "chartType": "radar",
  "frame": { "x": 1, "y": 1.5, "width": 5.5, "height": 4.5 },
  "data": {
    "categories": ["Liquidity", "Growth", "Margin", "Resilience"],
    "series": [{
      "id": "current",
      "name": "Current",
      "values": [72, 81, 64, 77],
      "stroke": { "color": "#0A84FF", "width": 2 },
      "marker": { "symbol": "circle", "size": 5 }
    }]
  }
}
```

Radar keeps the existing PPJ chart data and style structures. Categories are
shared around the perimeter; every series supplies one numeric value per
category. It has paired native category/value axes and may use the existing
series stroke and marker profiles.

## Native profile

The compiler emits `c:radarChart` with `c:radarStyle val="standard"`,
`c:varyColors val="0"`, ordinary `c:ser` children and two axis IDs. The shared
ChartSpace reader issues editability only when that exact profile is present.
Unknown children, transformed colors, unsupported marker graphs and alternate
radar styles keep the chart source owned and fail closed.

## Boundaries

- Supported: literal and bounded reference caches, multiple series, native
  category/value axes, legend, data labels, direct RGB stroke/marker styles,
  chart and plot surfaces, source-bound data/style continuation.
- Rejected: filled/marker-only radar styles, 3D, secondary axes, trendlines,
  error bars, null points and topology-changing edits.
- Waterfall remains declared for compatibility discovery but authored build
  continues to reject it until a native semantic profile is implemented.

## Verification

Extend one existing chart round-trip contract with radar and add one compact
PPJ build/project assertion. Do not create an effect matrix or new fixture farm.

