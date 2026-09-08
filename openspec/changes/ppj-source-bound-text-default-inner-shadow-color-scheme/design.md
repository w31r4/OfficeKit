# Design: source-bound paragraph default-text inner-shadow theme color

## Owner

The owner is one direct paragraph default run properties graph:

```text
p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:innerShdw/a:schemeClr/@val
```

The strict reader accepts one direct `a:innerShdw` with one direct theme color,
optionally followed by one valid direct `a:outerShdw`. The inner shadow keeps
its blur, distance, direction, RGB color, alpha, and other topology
source-owned. The paragraph index is the native index; a separate leaf kind
keeps this owner distinct from an inline run's
`textInnerShadowColorScheme`.

## Safety boundary

Multiple effect lists, glow, reflection, soft edge, non-direct or extension
children, malformed topology, missing or duplicate direct colors, and
paragraph/run topology changes remain source-owned or fail closed. No
source-bound insertion, removal, or reordering is allowed.

## Edit proof

The leaf kind is `textDefaultInnerShadowColorScheme` and its PPJ location is
`paragraph.style.defaultText.innerShadow.color`. The edit plan re-proves the
selected `a:p`, direct `a:pPr`, `a:defRPr`, bounded effect list, direct
`a:innerShdw`, and direct `a:schemeClr` before token-splicing only `val` in the
owning SlidePart. No protobuf or wire-version change is needed. The focused
fixture uses a direct theme-colored default inner shadow, changes the theme
token alongside the existing default-text and inline-run effect leaves,
verifies SlidePart-only mutation and package validity, and reprojects the
value.
