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
  and bounded markers.

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
