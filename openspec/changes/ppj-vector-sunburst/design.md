# Design

## Semantic hierarchy

Sunburst reuses the treemap-aligned `categories`, positive `values`, and nullable
`parents` arrays. It accepts one series, 1–96 globally named nodes, 1–16 roots,
at most six levels, an acyclic forest, and exact direct-child totals. The tighter
node/depth limits bound radial label pressure and custom-path size.

## Radial layout

Root angles follow root values in declared order. Each parent interval is
partitioned by its direct children. Depth maps to a fixed ring; optional inner
radius, ring gap, segment gap, start angle and direction are finite scalar
style. Root colors cycle by root and descendants blend toward white by depth.

Each annular sector is one native DrawingML custom geometry. Circular arcs are
approximated by deterministic cubic Bézier segments no wider than 90 degrees,
so the shape remains inside the existing PPJ custom-path vocabulary and can be
projected without a private arc command. Labels are ordinary text boxes placed
at the sector midpoint only when measured arc/ring space is sufficient.

## Recovery and animation

The authored PPTX embeds the exact PPJ. Without that program, import returns an
ordinary group of editable custom shapes and text. Whole-object animation is
allowed; ChartPart build modes are not.
