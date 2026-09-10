## Context

The paragraph default-run codec already reads and writes `fadeDir` through
`PresentationReflection.FadeDirectionAngle60000`. The PPJ projector also emits
the `fadeAngle` presentation value for this bounded reflection profile, but the
native leaf registry, proof, and XML patch path do not address the token.

## Goals / Non-Goals

**Goals:**

- issue one source-bound leaf for an existing paragraph default reflection
  `fadeDir` token;
- use the native unsigned 1/60000-degree integer with the existing strict
  circular-angle bound;
- splice one XML attribute and prove SlidePart-only change plus re-projection.

**Non-Goals:**

- adding or deleting a reflection or its `fadeDir` attribute;
- exposing reflection scale, skew, alignment, position, or rotate transforms in
  this field;
- broadening ordinary run reflection support, changing protobuf fields, or
  making host-rendering claims.

## Decisions

Use `textDefaultReflectionFadeAngleDegrees` as a distinct native leaf kind,
with PPJ location `paragraph.style.defaultText.reflection.fadeAngle`. The leaf
uses the paragraph index as its native index and stores the DrawingML `fadeDir`
integer as edit-plan authority while projecting degrees in 1/60000 units. This
matches the existing default-reflection scalar leaves while keeping the
source-bound owner explicit.

The edit proof re-reads the same shape, paragraph, `pPr`, `defRPr`, `effectLst`,
and single `reflection`, checks that its `fadeDir` token equals the expected
integer, and replaces only that token. A missing `fadeDir` receives no inferred
default; malformed, bounded, or complex graphs remain source-owned. The
validator accepts only canonical integer tokens in `0 <= token < 21600000`.

## Risks / Trade-offs

- [Angle precision] → convert through the existing 1/60000 scaled leaf grammar
  and compare the raw integer token before patching.
- [Effect graph widening] → keep the default-reflection proof strict and leave
  all other transforms source-owned.
- [Generated reference drift] → regenerate the presentation Skill/manual and
  run the focused source-bound regression plus OpenSpec and portability gates.
