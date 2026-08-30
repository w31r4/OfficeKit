# Design

## Semantic profile

One `treemap` chart owns one series. `data.categories` are globally unique node
names, `values` are positive node values, and `parents` is an aligned
array of string or null. Parents may appear anywhere in the array. Validation
proves that every named parent exists, the graph is a forest, maximum depth is
eight, there are at most sixteen roots and 128 nodes, and each non-leaf value
equals the sum of its direct children within a numeric tolerance.

The finite forest and explicit node bound make validation and layout terminate.
No expression language or recursive component evaluation is involved.

## Style profile

`style.treemap` requires one to sixteen `rootColors`. It may set a native border,
gap, header height, descendant lightening amount, label/value text styles, and
whether values are shown. Root colors cycle in deterministic root order;
descendants blend the root color toward white by depth. This is intentionally a
smaller and more predictable profile than an arbitrary per-node paint program.

## Native lowering

NativeAOT lays out each sibling set with a deterministic squarified algorithm,
then recurses into the bounded forest. Every node becomes an editable rectangle;
intermediate nodes reserve a header strip and leaves can add native text when
their measured frame is large enough. Compiler-owned child IDs use the reserved
`<element-id>/treemap/...` namespace so the PPJ node map only binds the outer
semantic element.

OOXML has no portable authored treemap ChartPart. The group representation is
therefore the honest native surface. It supports whole-object animation but not
ChartPart build modes.

## Recovery

An authored PPTX embeds the exact PPJ and restores the hierarchy. If the private
program is removed, import exposes the ordinary native group and does not infer
semantic parent/value relationships from geometry.
