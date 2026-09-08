# Design: source-bound picture shadow opacity

## Owner

The owner is one picture shape-properties graph:

```text
p:pic/a:spPr/a:effectLst/a:outerShdw/(a:srgbClr|a:schemeClr)/a:alpha/@val
```

The existing strict shadow reader preserves the alpha as a thousandth-percent
integer. This slice exposes it only when the direct color has one explicit
alpha token in the canonical `0` through `100000` range; the picture payload,
mask, border, shadow geometry, rotation flag and other effect topology remain
source-owned.

## Safety boundary

Missing alpha, multiple effect lists, multiple effects, missing or ambiguous
direct RGB/theme colors, non-direct or extension children, malformed topology,
and relationship changes remain source-owned or fail closed. No source-bound
insertion, removal, or reordering is allowed.

## Edit proof

The leaf kind is `imageShadowOpacityThousandthPercent` and its PPJ location is
`image.shadow.opacity`. The edit plan re-proves the selected picture, direct
picture properties, bounded effect list, direct `a:outerShdw`, direct color and
explicit alpha before token-splicing only that alpha attribute in the owning
`SlidePart`. No protobuf or wire-version change is needed. The focused
source-bound picture fixture adds the explicit rotation flag as an unrelated
shadow attribute, changes the alpha leaf, verifies SlidePart-only mutation and
package validity, and reprojects both opacity and the preserved rotation flag.
