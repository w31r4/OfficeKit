# Current primitive audit

## Verified current facts

- Generated PPJ reference: 2,202 lines.
- Compared finite DSL reference: 1,886 lines.
- PPJ preset geometries: 176.
- Compared preset geometries: 177 after excluding the table header.
- Exact preset difference: PPJ lacks `upArrow` and `lineInv`; PPJ alone has
  `flowChartData`.
- PPJ chart families: bar, column, line, area, pie, doughnut, scatter, bubble,
  radar, waterfall, heatmap, candlestick, treemap, sunburst, Sankey and combo.

## Classification

| Observation | Classification | Current decision |
| --- | --- | --- |
| named icon | semantic convenience gap | separate pinned offline-catalog design |
| inline LaTeX | semantic gap | separate formula representation/lowering design |
| `upArrow`, `lineInv` | finite semantic gap | implement now |
| independent line syntax | already covered | use `connector` free points/anchors/routing |
| remote URL image/fill | intentional constraint | keep local SHA-256 assets |
| table first/last/body style inheritance | convenience gap | implement now |
| Sankey right alignment and named colors | bounded vector gap | implement now |
| scatter/bubble/candlestick families | already covered individually | retain current bounded profiles |
| broader Cartesian/candlestick mixing | semantic gap | later chart topology/wire slice |
| stream area and pictographic bars | semantic gap | later vector/native chart slice |
| bubble scale/range | semantic gap | later chart plot-control slice |
| arbitrary pie angle/hole | semantic gap | later native chart writer slice |
| axis reverse/arrows/line/grid styling | semantic gap | later axis wire/writer slice |
| radar spoke styling | semantic gap | later radar axis slice |
| label number formats and point overrides | partial gap | later label/point-style slice |
| treemap/sunburst display depth | semantic gap | later vector hierarchy slice |

This audit compares observable language contracts. It does not copy a third
party implementation or claim parity from a visually similar raster result.
