# Current primitive audit

## Verified current facts

- Generated PPJ reference before this batch: 2,202 lines; after schema sync:
  2,233 lines.
- Compared finite DSL reference: 1,886 lines.
- PPJ preset geometries before this batch: 176; after adding the two missing
  public presets: 178.
- Compared preset geometries: 177 after excluding the table header.
- PPJ now includes the compared `upArrow` and `lineInv` spellings while
  retaining the compatible `flowChartData` alias.
- PPJ chart families: bar, column, line, area, pie, doughnut, scatter, bubble,
  radar, waterfall, heatmap, candlestick, treemap, sunburst, Sankey and combo.

## Classification

| Observation | Classification | Current decision |
| --- | --- | --- |
| named icon | implemented, bounded | pinned offline Font Awesome Free catalog |
| inline LaTeX | implemented, bounded | finite non-executable LaTeX to editable Office Math |
| `upArrow`, `lineInv` | implemented | shared preset geometry profile registry |
| independent line syntax | already covered | use `connector` free points/anchors/routing |
| remote URL image/fill | intentional constraint | keep local SHA-256 assets |
| table first/last/body style inheritance | implemented | bounded deterministic precedence |
| Sankey right alignment and named colors | implemented | bounded vector compiler |
| scatter/bubble/candlestick families | already covered individually | retain current bounded profiles |
| broader Cartesian/candlestick mixing | implemented, bounded | numeric Cartesian and OHLC overlay profiles |
| stream area and pictographic bars | implemented, bounded | editable vector compilers |
| bubble scale/range | partial | native scale/mode implemented; pixel-radius range remains unsupported |
| arbitrary pie angle/hole | implemented, bounded | native legal angle and hole-size ranges |
| axis reverse/arrows/line/grid styling | implemented, bounded | native direction, direct line/grid style and finite axis endpoints |
| radar spoke styling | implemented, bounded | semantic native spoke/ring coordinate system |
| label number formats and point overrides | implemented, bounded | plot defaults plus native series defaults and sparse zero-based point overrides; custom label text/layout/effects and source-linked formats remain source-owned |
| treemap/sunburst display depth | implemented, bounded | finite visible-level controls retain full hierarchy |

This audit compares observable language contracts. It does not copy a third
party implementation or claim parity from a visually similar raster result.

## Implemented batch

- `upArrow` and `lineInv` now share the same schema/profile source as the other
  authored preset geometries.
- Tables now have a bounded base style, cycling body-row styles, structural
  first/last row and column styles, explicit row/column conflict order, and
  cell-local final precedence.
- Sankey charts now support reverse-depth `right` alignment and exact
  declared-node color overrides; unknown node names fail validation.
- `connector` is explicitly documented as the one ordinary line primitive.
- Named icons, native inline formulas, finite stream/pictographic charts,
  numeric and candlestick overlays, circular geometry, hierarchy display
  levels, native plot/axis formatting, and the radar spoke coordinate system
  were subsequently closed as bounded PPJ features. Remaining limitations are
  recorded by capability rather than left as an undifferentiated language gap.

The existing comprehensive authored PPJ contract exercised all new state in
one compile/native inspection/import/determinism run. No effect matrix, new
fixture farm, or full test suite was added.
