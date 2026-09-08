# Design: source-bound picture shadow direction

## Owner

The owner is one picture shape-properties graph:

```text
p:pic/a:spPr/a:effectLst/a:outerShdw/@dir
```

The existing strict shadow reader preserves the optional direction as a
`1/60000`-degree integer. PPJ exposes the native leaf as degrees and normalizes
it back to the exact integer token before proof; the picture payload, mask,
border, shadow color, opacity, blur, distance, rotation flag and other effect
topology remain source-owned.

## Safety boundary

Missing `dir`, multiple effect lists, multiple effects, non-direct or
extension children, malformed topology, and relationship changes remain
source-owned or fail closed. No source-bound insertion, removal, or reordering
is allowed.

## Edit proof

The leaf kind is `imageShadowDirectionDegrees` and its PPJ location is
`image.shadow.angle`. The edit plan re-proves the selected picture, direct
picture properties, bounded effect list, direct `a:outerShdw`, and explicit
`dir` attribute before token-splicing only that attribute in the owning
`SlidePart`. No protobuf or wire-version change is needed. The focused
source-bound picture fixture keeps rotation, blur, distance, opacity and the
image payload as unrelated state, changes the direction leaf, verifies
SlidePart-only mutation and package validity, and reprojects the new angle.
