## Context

The paragraph default-run codec already reads and writes the full bounded
outer-shadow transform set, including `sx`, `sy`, `kx`, and `ky`. The native PPJ
projection currently advertises only the default-shadow color, geometry,
opacity, and `rotWithShape` leaves. Ordinary authored default-text shadows can
therefore carry `scaleX`, but source-bound callers cannot address that token.

## Goals / Non-Goals

**Goals:**

- issue one source-bound leaf for an existing paragraph default `sx` token;
- reuse the existing signed scale validation and effect-owner proof;
- splice one XML attribute and prove SlidePart-only change plus re-projection.

**Non-Goals:**

- adding or deleting an outer shadow or its `sx` attribute;
- exposing `scaleY`, `skewX`, `skewY`, WordArt, 3-D, or arbitrary effect graphs;
- changing authored shadow syntax, protobuf fields, or host rendering claims.

## Decisions

Use `textDefaultShadowScaleX` as a distinct native leaf kind, with PPJ location
`paragraph.style.defaultText.shadow.scaleX`. The leaf uses the paragraph index
as its native index and stores the DrawingML `sx` integer as its edit-plan
authority while projecting the ratio in 1/100000 units. This follows the
existing default-shadow leaf pattern and avoids conflating paragraph defaults
with direct run or shape shadows.

The edit proof re-reads the same shape, paragraph, `pPr`, `defRPr`,
`effectLst`, and single `outerShdw`, checks that its `sx` token equals the
expected integer, and replaces only that token. A missing `sx` does not receive
a default or an inferred value; malformed and complex graphs remain
source-owned.

## Risks / Trade-offs

- [Native ratio precision] → convert through the existing 1/100000 scaled leaf
  grammar and compare the raw integer token before patching.
- [Effect graph widening] → keep the default-shadow proof strict and leave
  other transforms source-owned.
- [Generated reference drift] → regenerate the presentation Skill/manual and
  run the focused source-bound regression plus OpenSpec and portability gates.
