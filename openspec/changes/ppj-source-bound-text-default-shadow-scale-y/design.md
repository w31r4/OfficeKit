## Context

The paragraph default-run codec already reads and writes the full bounded
outer-shadow transform set, including `sx`, `sy`, `kx`, and `ky`. The native PPJ
projection now exposes the source-bound `sx` token, but the matching `sy` token
still cannot be addressed independently after import.

## Goals / Non-Goals

**Goals:**

- issue one source-bound leaf for an existing paragraph default `sy` token;
- reuse the signed scale validation and effect-owner proof used by `scaleX`;
- splice one XML attribute and prove SlidePart-only change plus re-projection.

**Non-Goals:**

- adding or deleting an outer shadow or its `sy` attribute;
- exposing `scaleX`, `skewX`, `skewY`, WordArt, 3-D, or arbitrary effect graphs;
- changing authored shadow syntax, protobuf fields, or host rendering claims.

## Decisions

Use `textDefaultShadowScaleY` as a distinct native leaf kind, with PPJ location
`paragraph.style.defaultText.shadow.scaleY`. The leaf uses the paragraph index
as its native index and stores the DrawingML `sy` integer as edit-plan
authority while projecting the ratio in 1/100000 units. This follows the
existing `textDefaultShadowScaleX` pattern and keeps paragraph defaults
separate from direct run or shape shadows.

The edit proof re-reads the same shape, paragraph, `pPr`, `defRPr`, `effectLst`,
and single `outerShdw`, checks that its `sy` token equals the expected integer,
and replaces only that token. A missing `sy` receives no inferred default;
malformed and complex graphs remain source-owned.

## Risks / Trade-offs

- [Native ratio precision] → convert through the existing 1/100000 scaled leaf
  grammar and compare the raw integer token before patching.
- [Effect graph widening] → keep the default-shadow proof strict and leave
  other transforms source-owned.
- [Generated reference drift] → regenerate the presentation Skill/manual and
  run the focused source-bound regression plus OpenSpec and portability gates.
