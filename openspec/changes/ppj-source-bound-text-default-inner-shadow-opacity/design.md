# Design: source-bound paragraph default-text inner-shadow alpha

## Owner

The owner is one direct paragraph default run properties graph:

```text
p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:innerShdw/(a:srgbClr|a:schemeClr)/a:alpha/@val
```

The existing strict inner-shadow reader already validates one direct color
and at most one direct alpha. This slice exposes the alpha only when that
alpha is present and bounded; the paragraph index is the native index and
the color topology remains source-owned.

## Safety boundary

Missing or duplicate alpha, multiple effect lists, glow, reflection, soft
edge, non-direct or extension children, malformed topology, and paragraph or
run topology changes remain source-owned or fail closed. No source-bound
insertion, removal, or reordering is allowed.

## Edit proof

The leaf kind is
`textDefaultInnerShadowOpacityThousandthPercent` and its PPJ location is
`paragraph.style.defaultText.innerShadow.opacity`. The edit plan re-proves
the selected `a:p`, direct `a:pPr`, `a:defRPr`, bounded effect list, direct
`a:innerShdw`, direct color, and exactly one direct `a:alpha` before
token-splicing only `alpha/@val` in the owning `SlidePart`. No protobuf or
wire-version change is needed. The focused fixture uses an explicit
default-text inner-shadow alpha, changes the token, verifies SlidePart-only
mutation and package validity, and reprojects the value.
