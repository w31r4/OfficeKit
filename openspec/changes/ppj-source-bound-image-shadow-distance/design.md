# Design: source-bound picture shadow distance

## Owner

The owner is one picture shape-properties graph:

```text
p:pic/a:spPr/a:effectLst/a:outerShdw/@dist
```

The existing strict shadow reader preserves the optional distance as an EMU
integer. This slice exposes it only when the attribute is present and uses the
canonical non-negative integer range; the picture payload, mask, border,
shadow color, opacity, blur, rotation flag and other effect topology remain
source-owned.

## Safety boundary

Missing `dist`, multiple effect lists, multiple effects, non-direct or
extension children, malformed topology, and relationship changes remain
source-owned or fail closed. No source-bound insertion, removal, or reordering
is allowed.

## Edit proof

The leaf kind is `imageShadowDistanceEmu` and its PPJ location is
`image.shadow.distance`. The edit plan re-proves the selected picture, direct
picture properties, bounded effect list, direct `a:outerShdw`, and explicit
`dist` attribute before token-splicing only that attribute in the owning
`SlidePart`. No protobuf or wire-version change is needed. The focused
source-bound picture fixture keeps rotation, blur, opacity and the image
payload as unrelated state, changes the distance leaf, verifies SlidePart-only
mutation and package validity, and reprojects the new distance.
