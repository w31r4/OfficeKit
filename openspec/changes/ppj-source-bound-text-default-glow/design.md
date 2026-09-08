# Design: source-bound paragraph default-text glow radius

## Owner

The owner is one direct paragraph default run properties graph:

```text
p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:glow
```

The strict reader accepts one direct `a:glow` with a bounded `rad`, followed by
no effect or one valid direct `a:outerShdw`. The glow has no modeled children.
The paragraph index is the native index; a separate leaf kind keeps this owner
distinct from an inline run's `textGlowRadiusEmu`.

## Safety boundary

Multiple effect lists, inner shadow, reflection, soft edge, non-direct or
extension children, malformed radius values, missing or duplicate direct
owners, and paragraph/run topology changes remain source-owned or fail closed.
No source-bound insertion, removal, or reordering is allowed. Color and alpha
are preserved as source-owned values and do not receive leaves in this slice.

## Edit proof

The leaf kind is `textDefaultGlowRadiusEmu` and its PPJ location is
`paragraph.style.defaultText.glow.radius`. The edit plan re-proves the
selected `a:p`, direct `a:pPr`, `a:defRPr`, strict effect list, and direct
`a:glow` before token-splicing only `rad` in the owning SlidePart. No protobuf
or wire-version change is needed. The focused fixture removes the embedded
PPJ, changes the paragraph default and inline run effect owners independently,
verifies SlidePart-only mutation and package validity, and reprojects all
values.
