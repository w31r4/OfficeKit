# Design: source-bound paragraph default-text reflection blur

## Owner

The owner is one direct paragraph default run properties graph:

```text
p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:reflection
```

The strict reader accepts one full-span direct `a:reflection` with a bounded
`blurRad`, optionally preceded by one valid direct `a:outerShdw`. The reflection
keeps its start/end opacity, distance, direction and full-span positions
source-owned. The paragraph index is the native index; a separate leaf kind
keeps this owner distinct from an inline run's
`textReflectionBlurRadiusEmu`.

## Safety boundary

Multiple effect lists, glow, inner shadow, soft edge, non-direct or extension
children, malformed geometry, missing or duplicate direct owners, and paragraph
/run topology changes remain source-owned or fail closed. No source-bound
insertion, removal, or reordering is allowed.

## Edit proof

The leaf kind is `textDefaultReflectionBlurRadiusEmu` and its PPJ location is
`paragraph.style.defaultText.reflection.blur`. The edit plan re-proves the
selected `a:p`, direct `a:pPr`, `a:defRPr`, strict full-span effect list, and
direct `a:reflection` before token-splicing only `blurRad` in the owning
SlidePart. No protobuf or wire-version change is needed. The focused fixture
removes the embedded PPJ, changes the paragraph default blur alongside the
existing default-text and inline-run effect leaves, verifies SlidePart-only
mutation and package validity, and reprojects all values.
