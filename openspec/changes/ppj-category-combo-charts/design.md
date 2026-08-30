# Design

## Language boundary

A combo chart has aligned shared categories and 2–256 series. Every series
declares `chartType: column | line | area` and `axis: primary | secondary`.
At least two distinct plot families and at least one primary family are
required. All series of one family use the same axis pair. A secondary pair is
created only when at least one complete family selects it.

Horizontal bars are excluded because their category/value orientation is not
compatible with the shared category axis used by line and area plots. Scatter
and bubble remain ordinary numeric charts until PPJ owns a separate mixed
numeric-axis contract.

## Native chart

The compiler emits at most one `c:barChart`, `c:areaChart` and `c:lineChart`.
Area is emitted behind columns, and line is emitted last so evidence strokes
and markers stay visible. Primary plots reference axis IDs 1/2; secondary plots
reference 3/4. Series retain their PPJ order through native `idx` and `order`.

Import accepts the same two-or-three-plot profile, verifies one shared category
array, identical data-label semantics and one consistent axis pair per family,
then projects stable per-series type and axis fields. Fixed-topology edits patch
the existing plots without rebuilding unrelated chart XML.

## Recovery

Embedded PPJ remains exact. Removing the embedded snapshot and importing the
PPTX still recovers the bounded native combo as typed PPJ. Unsupported numeric
or irregular combinations remain opaque rather than being coerced.

