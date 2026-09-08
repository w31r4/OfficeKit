# Design: source-bound paragraph default-text outer-shadow rotation

## Owner

The owner is one direct paragraph default run properties graph:

```text
p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:outerShdw/@rotWithShape
```

The existing strict outer-shadow reader already preserves the optional
`rotateWithShape` value. This slice exposes it only when the attribute is
present and encoded as the bounded `0`/`1` boolean form; the paragraph index is
the native index and all other effect/color/geometry topology remains
source-owned.

## Safety boundary

Missing `rotWithShape`, multiple effect lists, glow, inner shadow, reflection,
soft edge, non-direct or extension children, malformed topology, and paragraph
or run topology changes remain source-owned or fail closed. No source-bound
insertion, removal, or reordering is allowed.

## Edit proof

The leaf kind is `textDefaultShadowRotateWithShape` and its PPJ location is
`paragraph.style.defaultText.shadow.rotateWithShape`. The edit plan re-proves
the selected `a:p`, direct `a:pPr`, `a:defRPr`, bounded effect list, direct
`a:outerShdw`, and explicit `rotWithShape` attribute before token-splicing only
that attribute in the owning `SlidePart`. No protobuf or wire-version change
is needed. The focused fixture adds the explicit authored boolean, changes the
leaf, verifies SlidePart-only mutation and package validity, and reprojects the
value.
