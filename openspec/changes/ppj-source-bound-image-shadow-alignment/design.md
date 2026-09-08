# Design: source-bound picture shadow alignment

## Owner

The owner is one picture shape-properties graph:

```text
p:pic/a:spPr/a:effectLst/a:outerShdw/@algn
```

The existing strict shadow reader preserves the optional alignment as the
canonical `tl`, `t`, `tr`, `l`, `ctr`, `r`, `bl`, `b`, or `br` token. This slice
exposes it only when the attribute is present; the picture payload, mask,
border, shadow color, opacity, geometry and other effect topology remain
source-owned.

## Safety boundary

Missing `algn`, multiple effect lists, multiple effects, non-direct or
extension children, malformed topology, and relationship changes remain
source-owned or fail closed. No source-bound insertion, removal, or reordering
is allowed.

## Edit proof

The leaf kind is `imageShadowAlignment` and its PPJ location is
`image.shadow.alignment`. The edit plan re-proves the selected picture, direct
picture properties, bounded effect list, direct `a:outerShdw`, and explicit
`algn` attribute before token-splicing only that attribute in the owning
`SlidePart`. No protobuf or wire-version change is needed. The focused
source-bound picture fixture keeps rotation, blur, distance, opacity and the
image payload as unrelated state, changes the alignment leaf, verifies
SlidePart-only mutation and package validity, and reprojects the new token.
