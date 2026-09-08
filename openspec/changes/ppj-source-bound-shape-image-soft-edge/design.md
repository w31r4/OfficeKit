# Design: source-bound shape and image soft edge

## Owners

The shape owner is:

```text
p:sp/p:spPr/a:effectLst/[a:outerShdw|a:reflection]a:softEdge
```

The picture owner is the corresponding direct `p:pic/p:spPr` effect list.
The strict reader accepts one direct `a:softEdge` with bounded `rad`, no child
elements, and no preceding effect, one valid direct outer shadow, or one valid
full-span direct reflection.

The PPJ projection uses `style.softEdge` for ordinary shapes and lines, and
`softEdge` for pictures. Each source-bound owner issues one
`shapeSoftEdgeRadiusEmu` or `imageSoftEdgeRadiusEmu` native leaf.

## Safety boundary

Multiple effect lists, glow, inner shadow, outer-shadow plus reflection chains,
non-full-span reflection, reflection transform variants, unknown children,
extensions, malformed or out-of-range radii, and topology changes remain
source-owned or fail closed. No source-bound insertion, removal, reordering, or
effect conversion is allowed.

## Edit proof

The edit plan re-proves the selected `p:sp` or `p:pic`, direct effect-list
topology, and `a:softEdge/@rad` precondition before token-splicing only the
radius in the owning SlidePart. The focused fixture removes the embedded PPJ,
changes both shape and picture radii, verifies SlidePart-only mutation,
retained preceding effects, package validity, and second-projection recovery.
