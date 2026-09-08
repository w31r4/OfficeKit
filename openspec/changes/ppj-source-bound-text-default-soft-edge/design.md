# Design: source-bound paragraph default-text soft edge

## Owner

The owner is one direct paragraph default run properties graph:

```text
p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:softEdge
```

The strict reader accepts one direct `a:softEdge` with a bounded `rad`, with
no preceding effect, one valid direct `a:outerShdw`, or one valid full-span
direct `a:reflection`. The soft edge has no modeled children. The paragraph
index is the native index; a separate leaf kind keeps this owner distinct
from an inline run's `textSoftEdgeRadiusEmu`.

## Safety boundary

Multiple effect lists, glow, inner shadow, combined effect chains,
non-full-span reflection, reflection transform variants, unknown DrawingML
children, extensions, malformed radius values, missing or duplicate direct
owners, and paragraph/run topology changes remain source-owned or fail closed.
No source-bound insertion, removal, or reordering is allowed.

## Edit proof

The leaf kind is `textDefaultSoftEdgeRadiusEmu` and its PPJ location is
`paragraph.style.defaultText.softEdge.radius`. The edit plan re-proves the
selected `a:p`, direct `a:pPr`, `a:defRPr`, strict effect list, and direct
`a:softEdge` before token-splicing only `rad` in the owning SlidePart. No
protobuf or wire-version change is needed. The focused fixture removes the
embedded PPJ, changes both the paragraph default and inline run radii,
verifies SlidePart-only mutation and package validity, and reprojects both
values.
