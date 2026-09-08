# Design: source-bound picture shadow theme color

## Owner

The owner is one picture shape-properties graph:

```text
p:pic/a:spPr/a:effectLst/a:outerShdw/a:schemeClr/@val
```

The existing strict shadow reader preserves the canonical theme token. This
slice exposes it only when the direct color is a supported bare scheme token;
the picture payload, mask, border, shadow geometry, rotation flag, alpha and
other effect topology remain source-owned.

## Safety boundary

Missing color, RGB color, color transforms, multiple effect lists, multiple
effects, non-direct or extension children, malformed topology, and
relationship changes remain source-owned or fail closed. No source-bound
insertion, removal, or reordering is allowed.

## Edit proof

The leaf kind is `imageShadowColorScheme` and its PPJ location is
`image.shadow.color`. The edit plan re-proves the selected picture, direct
picture properties, bounded effect list, direct `a:outerShdw`, and explicit
`a:schemeClr` before token-splicing only its `val` attribute in the owning
`SlidePart`. No protobuf or wire-version change is needed. The focused
source-bound picture fixture keeps rotation, blur, distance and alpha as
unrelated shadow state, changes the theme-color leaf, verifies SlidePart-only
mutation and package validity, and reprojects the new token.
