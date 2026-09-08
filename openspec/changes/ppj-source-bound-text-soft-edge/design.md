# Design: source-bound text soft edge

## Owner

The owner is a direct DrawingML rich-text run:

```text
p:sp/a:txBody/a:p/a:r/a:rPr/a:effectLst/a:softEdge
```

The strict reader accepts one direct `a:softEdge` with a bounded `rad`, with
no preceding effect, one valid direct `a:outerShdw`, or one valid full-span
direct `a:reflection`. The soft edge has no modeled children. A preceding
multi-effect chain is deliberately outside this owner.

Only ordinary text runs issue `textSoftEdgeRadiusEmu`. Paragraph `defaultText`
is projected through the existing text-style owner but does not issue an
independent source-bound leaf in this slice.

## Safety boundary

Multiple effect lists, glow, inner shadow, combined glow/inner/outer/reflection
chains, non-full-span reflection, reflection transform variants, unknown
DrawingML children, extensions, malformed radius values, fields, breaks, and
other inline topologies remain source-owned or fail closed. No source-bound
insertion, removal, reordering, or effect conversion is allowed.

## Edit proof

The leaf carries the existing source/native hashes and exact text-leaf index.
The edit plan re-proves the selected `a:r`, direct `a:rPr`, strict effect
topology, and direct `a:softEdge` before token-splicing only `rad` in the
owning SlidePart. The focused fixture removes the embedded PPJ snapshot,
changes the radius, verifies SlidePart-only mutation and package validity, and
reprojects the value.
