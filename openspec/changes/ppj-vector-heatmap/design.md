## Context

PowerPoint has no standard native heatmap ChartPart. Treating a heatmap as a
native chart would therefore overstate editability, while rendering it to one
image would discard the strongest OfficeKit advantage. PPJ already has a
stable semantic chart node, an editable DrawingML group writer, exact embedded
program recovery and deterministic object identities.

## Decisions

### Use the existing matrix-shaped chart data model

`data.categories` supplies ordered x labels. Each `data.series[]` supplies one
ordered y label through `name` and one row of values. Categories and row names
must be unique non-empty strings; the bounded profile accepts at most 32 by 32
cells. Missing values remain explicit `null` cells.

### Lower to a native group

The compiler emits one `p:grpSp` at the chart frame. Its children are ordinary
editable rectangles and text boxes for the title, axes, cells, optional value
labels and optional colorbar. Child IDs derive from the stable PPJ chart ID and
matrix coordinates. This avoids a private raster payload and keeps visual
editing possible in PowerPoint.

The native group is a lowering, not a second semantic format. Embedded PPJ
recovers the exact heatmap. Without that source program, projection returns the
truthful group; arbitrary imported shape grids are never inferred as heatmaps.

### Require an explicit scale style

`style.heatmap` is required. A linear scale uses exactly two colors. A
diverging scale uses exactly three colors and a midpoint (default zero). An
optional two-number domain overrides the data-derived range. Values outside
the domain clamp to the endpoint colors. Equal endpoints, a midpoint outside
the domain and non-finite values fail before output.

### Keep the first profile finite

The profile owns one title, row/column labels, a flat rectangular matrix,
optional per-cell values and one vertical colorbar. It excludes clustering,
hierarchical axes, nonlinear scales, tooltip state, arbitrary formulas and
chart-build animation. Whole-object entrance animation remains valid because
the native target is one group.

## Risks and mitigations

- **Tiny unreadable cells:** reject frames whose resolved cell width or height
  is below the documented minimum instead of silently shrinking labels.
- **Text contrast:** choose black or white value labels from computed cell
  luminance unless the user supplies a bounded value text style color.
- **Semantic loss after external cleanup:** exact semantics depend on embedded
  PPJ by design; fallback projection is explicitly a group rather than a false
  round-trip claim.
- **Overusing heatmaps:** Skill guidance requires a genuine two-dimensional
  matrix question and discourages decorative color grids.
