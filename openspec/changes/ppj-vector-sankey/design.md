# Design

## Data contract

One series owns positive `values` plus aligned `sources` and `targets` arrays.
`data.categories` declares 2–64 unique node names and stable color/order identity.
The graph has 1–256 edges, is acyclic, every endpoint exists, and every internal
node conserves flow within numeric tolerance.

## Native layout

Topological longest-path depth selects a column. Optional sink justification
moves terminal nodes to the final column. Within each column, declared node
order is stable. One global scale maps flow magnitude to node and ribbon
thickness; finite node/column gaps keep the layout readable.

Flows are emitted first as closed cubic custom-geometry ribbons. Native node
rectangles and bounded labels follow, so their z-order remains legible and all
objects stay editable. Colors cycle by node and flow color follows its source
or target as declared.

## Recovery

Embedded PPJ restores graph semantics exactly. Without it, import returns an
ordinary group of custom shapes and text. Whole-object animation is supported;
ChartPart build modes are not.
