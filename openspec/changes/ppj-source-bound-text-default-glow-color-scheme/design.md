# Design: source-bound paragraph default-text glow theme color

## Owner

The owner is one direct paragraph default run properties graph:

```text
p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:glow/a:schemeClr/@val
```

The strict reader accepts one direct `a:glow` with one direct theme color,
optionally followed by one valid direct `a:outerShdw`. The glow radius, RGB
color, and alpha remain source-owned in this slice. The paragraph index is
the native index; a separate leaf kind keeps this owner distinct from an
inline run's `textGlowColorScheme`.

## Safety boundary

Multiple effect lists, inner shadow, reflection, soft edge, non-direct or
extension children, malformed topology, duplicate colors, and RGB colors
remain source-owned or fail closed. No source-bound insertion, removal, or
reordering is allowed.

## Edit proof

The leaf kind is `textDefaultGlowColorScheme` and its PPJ location is
`paragraph.style.defaultText.glow.color`. The edit plan re-proves the selected
`a:p`, direct `a:pPr`, `a:defRPr`, bounded effect list, direct `a:glow`, and
direct `a:schemeClr` before token-splicing only `val` in the owning
`SlidePart`. No protobuf or wire-version change is needed. The focused fixture
uses a direct theme-colored default glow, changes the theme token, verifies
SlidePart-only mutation and package validity, and reprojects the value.
