# Design: source-bound picture shadow blur

## Owner

The owner is one picture shape-properties graph:

```text
p:pic/a:spPr/a:effectLst/a:outerShdw/@blurRad
```

The existing strict shadow reader preserves the optional blur radius as an EMU
integer. This slice exposes it only when the attribute is present and uses the
canonical non-negative integer range; the picture payload, mask, border,
shadow color, opacity, distance, rotation flag and other effect topology remain
source-owned.

## Safety boundary

Missing `blurRad`, multiple effect lists, multiple effects, non-direct or
extension children, malformed topology, and relationship changes remain
source-owned or fail closed. No source-bound insertion, removal, or reordering
is allowed.

## Edit proof

The leaf kind is `imageShadowBlurRadiusEmu` and its PPJ location is
`image.shadow.blur`. The edit plan re-proves the selected picture, direct
picture properties, bounded effect list, direct `a:outerShdw`, and explicit
`blurRad` attribute before token-splicing only that attribute in the owning
`SlidePart`. No protobuf or wire-version change is needed. The focused
source-bound picture fixture keeps rotation, distance, opacity and the image
payload as unrelated state, changes the blur leaf, verifies SlidePart-only
mutation and package validity, and reprojects the new blur.
