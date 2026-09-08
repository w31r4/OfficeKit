# Design: source-bound paragraph default-text inner-shadow blur

## Owner

The owner is one direct paragraph default run properties graph:

```text
p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:innerShdw
```

The strict reader accepts one direct `a:innerShdw` with a bounded `blurRad`,
optionally followed by one valid direct `a:outerShdw`. The inner shadow keeps
its direct color, distance, direction, and alpha source-owned. The paragraph
index is the native index; a separate leaf kind keeps this owner distinct from
an inline run's `textInnerShadowBlurRadiusEmu`.

## Safety boundary

Multiple effect lists, glow, reflection, soft edge, non-direct or
extension children, malformed geometry, missing or duplicate direct owners,
and paragraph/run topology changes remain source-owned or fail closed. No
source-bound insertion, removal, or reordering is allowed.

## Edit proof

The leaf kind is `textDefaultInnerShadowBlurRadiusEmu` and its PPJ location is
`paragraph.style.defaultText.innerShadow.blur`. The edit plan re-proves the
selected `a:p`, direct `a:pPr`, `a:defRPr`, strict effect list, and direct
`a:innerShdw` before token-splicing only `blurRad` in the owning SlidePart. No
protobuf or wire-version change is needed. The focused fixture removes the
embedded PPJ, changes the paragraph default blur alongside the existing
default-text and inline-run effect leaves, verifies SlidePart-only mutation
and package validity, and reprojects all values.
