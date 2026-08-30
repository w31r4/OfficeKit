## Context

Kimi exposes stream stacking as a rendering semantic on area series. Native
PowerPoint ChartParts have ordinary and percentage stacking, but no stable
centered-stream mode. OfficeKit already compiles heatmaps, candlesticks,
treemaps, sunbursts and Sankey diagrams into editable DrawingML groups when a
native ChartPart cannot represent the requested semantics.

## Decisions

### 1. Reuse the area-chart vocabulary

An authored streamgraph is `chartType: "area"` with
`style.stacking: "stream"`. Data remains ordinary ordered categories plus two
through twelve aligned non-negative series. No second DSL or special data
table is introduced.

### 2. Compile a bounded centered silhouette

For each category, the sum of all series is centered around the plot midline.
Each series becomes one closed cubic DrawingML path between its lower and upper
cumulative boundaries. This preserves relative thickness and total-magnitude
variation rather than silently normalizing every category to 100 percent.

The compiler emits stable child IDs, direct fills, optional direct strokes,
category labels and end labels. The result remains editable native geometry;
it is not an image or opaque SVG.

### 3. Keep the truthful boundary explicit

The bounded profile requires 2..12 series, 3..64 unique ordered categories,
finite non-negative values, and at least one positive value at every category.
It rejects secondary axes, markers, trendlines, error bars, data labels and
other native-ChartPart-only options. Ordinary line or area charts remain the
right choice when precise point lookup or independent baselines matter.

### 4. Recovery follows authored PPJ truth

OfficeKit-authored PPTX embeds the PPJ program and restores `stream` exactly.
Without that snapshot, the DrawingML group imports truthfully as editable
shapes and is never guessed back into a streamgraph. Source-bound third-party
ChartParts remain unchanged and receive no stream mutation capability.

## Lean verification

Extend the existing comprehensive authored-PPJ contract with one streamgraph.
Prove bounded validation, native editable paths, deterministic build and exact
embedded-PPJ recovery. Do not create a new fixture or chart matrix.

