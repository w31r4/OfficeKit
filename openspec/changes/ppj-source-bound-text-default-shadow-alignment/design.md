# Design: source-bound paragraph default-text outer-shadow alignment

## Owner

The owner is one direct paragraph default run properties graph:

```text
p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:outerShdw/@algn
```

The existing strict outer-shadow reader validates one direct outer shadow and
its direct color/geometry. This slice exposes only the existing alignment;
the paragraph index is the native index and all other shadow attributes remain
source-owned.

## Safety boundary

Missing or malformed `algn`, multiple effect lists, glow, inner shadow,
reflection, soft edge, non-direct or extension children, and paragraph or run
topology changes remain source-owned or fail closed. No source-bound
insertion, removal, or reordering is allowed.

## Edit proof

The leaf kind is `textDefaultShadowAlignment` and its PPJ location is
`paragraph.style.defaultText.shadow.alignment`. The edit plan re-proves the
selected `a:p`, direct `a:pPr`, `defRPr`, one direct `a:effectLst`, and one
direct `a:outerShdw` before token-splicing only `algn` in the owning
`SlidePart`. No protobuf or wire-version change is needed. The focused fixture
uses an explicit default-text outer-shadow alignment, changes the token,
verifies SlidePart-only mutation and package validity, and reprojects the
value.
