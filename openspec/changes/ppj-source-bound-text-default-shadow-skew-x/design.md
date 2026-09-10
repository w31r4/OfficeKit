## Context

The paragraph default-run codec already reads and writes the full bounded
outer-shadow transform set, including `sx`, `sy`, `kx`, and `ky`. The native
PPJ projection exposes the source-bound `sx` and `sy` tokens, but the matching
`kx` token still cannot be addressed independently after import.

## Goals / Non-Goals

**Goals:**

- issue one source-bound leaf for an existing paragraph default `kx` token;
- use the native 1/60000-degree signed integer and strict skew bounds;
- splice one XML attribute and prove SlidePart-only change plus re-projection.

**Non-Goals:**

- adding or deleting an outer shadow or its `kx` attribute;
- exposing `skewY`, WordArt, 3-D, or arbitrary effect graphs;
- changing authored shadow syntax, protobuf fields, or host rendering claims.

## Decisions

Use `textDefaultShadowSkewX` as a distinct native leaf kind, with PPJ location
`paragraph.style.defaultText.shadow.skewX`. The leaf uses the paragraph index
as its native index and stores the DrawingML `kx` integer as edit-plan
authority while projecting degrees in 1/60000 units. This follows the existing
default-shadow scale leaves and keeps paragraph defaults separate from direct
run or shape shadows.

The edit proof re-reads the same shape, paragraph, `pPr`, `defRPr`, `effectLst`,
and single `outerShdw`, checks that its `kx` token equals the expected integer,
and replaces only that token. A missing `kx` receives no inferred default;
malformed, bound, and complex graphs remain source-owned. The validator accepts
only canonical signed Int32 tokens strictly between -5400000 and 5400000.

## Risks / Trade-offs

- [Angle precision] → convert through the existing 1/60000 scaled leaf grammar
  and compare the raw integer token before patching.
- [Effect graph widening] → keep the default-shadow proof strict and leave
  `skewY` and other transforms source-owned.
- [Generated reference drift] → regenerate the presentation Skill/manual and
  run the focused source-bound regression plus OpenSpec and portability gates.
