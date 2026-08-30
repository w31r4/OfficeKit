## Context

DrawingML represents the rotation of pie and doughnut plots with
`c:firstSliceAng` and represents a doughnut's center opening with `c:holeSize`.
Both are scalar plot properties; neither changes series topology, relationships
or cached data. OfficeKit already owns the surrounding canonical pie/doughnut
ChartSpace profile but unnecessarily accepts only zero and fifty.

## Decisions

### 1. Keep the public vocabulary semantic

`chart.style.startAngle` is an integer from 0 through 360 for `pie` and
`doughnut`. `chart.style.holeSize` is an integer from 10 through 90 for
`doughnut` only. PPJ does not expose native element names.

### 2. Preserve native absence

Additive optional wire fields distinguish an absent first-slice angle from an
explicit zero. An authored doughnut still emits the required canonical native
hole size, defaulting to 50 when PPJ omits the field. Projection records the
native value when present. Other chart types reject either field.

### 3. Treat imported edits as bounded plot edits

An editable imported pie or doughnut receives `setChartPlot` for
`chart.plot`. A PPJ diff may change only the two modeled scalars and the codec
patches the existing ChartPart. The source hash, target hash, chart type,
series topology and all unknown siblings are re-proved before writing.

### 4. Make the Agent use orientation deliberately

The chart guide recommends rotation only to align the first slice, labels or a
center annotation with the page composition. Hole size exists to reserve a
meaningful center field, not to turn every pie into a decorative ring.

## Lean verification

Extend the existing comprehensive authored PPJ contract with one doughnut
chart. Prove custom native scalars, projection, capability issuance, one
source-bound edit and reimport. Do not create a chart-option matrix, fixture
farm or separate test file.

