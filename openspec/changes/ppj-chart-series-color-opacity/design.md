# Design

`series.color` remains a compact authored alias. Opaque colors retain the
legacy direct series-color representation. A color with alpha lowers to
`SpreadsheetChartSurfaceFill.solid_rgb` plus its presence-aware opacity, the
same representation already used by `series.fill`.

Projection deliberately emits the canonical structured form:

```json
{
  "fill": { "type": "solid", "color": "#0A84FF", "opacity": 0.5 }
}
```

This prevents two equivalent spellings from producing unstable reimported
programs. Imported source-bound changes still require `setChartFill`; a
`setChartData` capability cannot change opacity.

