## Context

See proposal.md. Shared chart typography already transports glow, outer shadow and soft edge. PptxInnerShadowCodec accepts optional native geometry but currently belongs to Presentation; chart parsers must not depend on that assembly.

## Goals / Non-Goals

Goals: preserve inner-shadow presence and independent sibling effects throughout all shared chart text owners and vector defaults. Non-goals: reflection, unknown effects and complete F-07 or host verification.

## Decisions

- Introduce chartTextInnerShadow rather than weakening ordinary text innerShadow: only color is required; blur 0..1000 points, distance 0..100000 points, angle -360..360 degrees, opacity 0..1 or opacity token. Reuse chart color/theme/opacity resolution and ties-even native rounding.
- Append PresentationInnerShadow at chart text wire field 22. Move the existing value codec to Shared and exclude it from Presentation compilation; keep ordinary text behavior intact.
- Replace prefix peeling with strict ordered per-effect parsing of glow, innerShdw, outerShdw, softEdge. Reject repeats, reordered/unknown nodes, extra attributes or nested alpha extensions. Enforce chart geometry upper bounds as well as native validation.
- Project only present attributes; deletion removes only the selected field/effect. Vector default styles clone only into unset run effects; explicit run overrides remain authoritative.

## Risks / Trade-offs

- Parser composition could widen unrelated profiles: retain negative source-owned graph tests and schema validation of the full supported order.
- Optional alpha/theme identity could be materialized: test explicit zero, absence, direct theme and declared grammar tokens through fresh projection.
- Shared working tree has unrelated preview changes: implement and publish exact scoped blobs from an isolated checkout.
