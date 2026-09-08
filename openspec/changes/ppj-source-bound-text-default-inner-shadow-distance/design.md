# Design: source-bound paragraph default-text inner-shadow distance

## Owner

The owner is one direct paragraph default run properties graph:

```text
p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:innerShdw
```

The strict reader accepts one direct `a:innerShdw` with a bounded `dist`,
optionally followed by one valid direct `a:outerShdw`. The inner shadow keeps
its blur, direction, color and alpha source-owned. The paragraph index is the
native index; a separate leaf kind keeps this owner distinct from an inline
run's `textInnerShadowDistanceEmu`.

## Safety boundary

Multiple effect lists, glow, reflection, soft edge, non-direct or extension
children, malformed geometry, missing or duplicate direct owners, and
paragraph/run topology changes remain source-owned or fail closed. No
source-bound insertion, removal, or reordering is allowed.

## Edit proof

The leaf kind is `textDefaultInnerShadowDistanceEmu` and its PPJ location is
`paragraph.style.defaultText.innerShadow.distance`. The edit plan re-proves
the selected `a:p`, direct `a:pPr`, `a:defRPr`, bounded effect list, and direct
`a:innerShdw` before token-splicing only `dist` in the owning SlidePart. No
protobuf or wire-version change is needed. The focused fixture removes the
embedded PPJ, changes the paragraph default distance alongside the existing
default-text and inline-run effect leaves, verifies SlidePart-only mutation and
package validity, and reprojects the value.
