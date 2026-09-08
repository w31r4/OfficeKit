# Design: source-bound paragraph default-text outer-shadow distance

## Owner

The owner is one direct paragraph default run properties graph:

```text
p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:outerShdw/@dist
```

The existing strict outer-shadow reader validates one direct outer shadow and
its direct color/geometry. This slice exposes only the existing distance; the
paragraph index is the native index and all other shadow attributes remain
source-owned.

## Safety boundary

Missing or malformed `dist`, multiple effect lists, glow, inner shadow,
reflection, soft edge, non-direct or extension children, and paragraph or run
topology changes remain source-owned or fail closed. No source-bound
insertion, removal, or reordering is allowed.

## Edit proof

The leaf kind is `textDefaultShadowDistanceEmu` and its PPJ location is
`paragraph.style.defaultText.shadow.distance`. The edit plan re-proves the
selected `a:p`, direct `a:pPr`, `a:defRPr`, one direct `a:effectLst`, and one
direct `a:outerShdw` before token-splicing only `dist` in the owning
`SlidePart`. No protobuf or wire-version change is needed. The focused fixture
uses an explicit default-text outer-shadow distance, changes the token,
verifies SlidePart-only mutation and package validity, and reprojects the
value.
