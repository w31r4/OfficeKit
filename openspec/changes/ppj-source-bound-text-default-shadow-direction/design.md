# Design: source-bound paragraph default-text outer-shadow direction

## Owner

The owner is one direct paragraph default run properties graph:

```text
p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:outerShdw/@dir
```

The existing strict outer-shadow reader validates one direct outer shadow and
its direct color/geometry. This slice exposes only the existing direction in
60000ths of a degree; the paragraph index is the native index and all other
shadow attributes remain source-owned.

## Safety boundary

Missing or malformed `dir`, multiple effect lists, glow, inner shadow,
reflection, soft edge, non-direct or extension children, and paragraph or run
topology changes remain source-owned or fail closed. No source-bound
insertion, removal, or reordering is allowed.

## Edit proof

The leaf kind is `textDefaultShadowDirectionDegrees` and its PPJ location is
`paragraph.style.defaultText.shadow.angle`. The edit plan re-proves the
selected `a:p`, direct `a:pPr`, `defRPr`, one direct `a:effectLst`, and one
direct `a:outerShdw` before token-splicing only `dir` in the owning
`SlidePart`. No protobuf or wire-version change is needed. The focused fixture
uses an explicit default-text outer-shadow direction, changes the token,
verifies SlidePart-only mutation and package validity, and reprojects the
value.
