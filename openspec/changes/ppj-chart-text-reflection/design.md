## Context

See proposal.md. PresentationReflection already carries five optional native scalars. PptxReflectionCodec reads optional values on an explicit full-span native owner but currently requires and materializes every value when writing. Chart character effects already have a strict rank-based parser for four independent effects.

## Goals / Non-Goals

Goals: close chart reflection through every shared text style, including presence and independent lifecycle. Non-goals: partial spans, extra native transform attributes, production SVG adoption or host appearance.

## Decisions

- Add chartTextReflection with the existing five property constraints and no required scalars. Empty object means an explicit full-span effect. Ordinary PPJ reflection retains its five required properties, avoiding an authored language change there.
- Reuse PresentationReflection at chart style field 23. Move PptxReflectionCodec into Shared; allow optional native scalars in its validator and write only present attributes. Keep its explicit stPos=0/endPos=100000 owner proof and ordinary imported graph policy.
- Insert reflection between outerShdw and softEdge in the chart effect parser/writer. Parse only childless, known-attribute native reflections via the shared direct reader. Empty/partial native properties retain wire presence.
- Reuse chart range/opacity token conversion and ties-even EMU/angle rounding. Project only present keys. Vector defaults copy only into runs without an explicit reflection; ordinary explicit rich-run reflection remains fully authored under its existing schema.

## Risks / Trade-offs

- Shared native validator changes could regress ordinary text/shapes: run existing reflection/effect tests and verify partial-value read/write explicitly; ordinary authored schema stays intact.
- Source-bound empty object could be confused with deletion: lifecycle checks both states from fresh native projections, including exact no-op and untouched ZIP entries.
- Unsupported spans/transforms remain outside the wire representation: keep them source-owned and test no-op/edit rejection instead of normalizing them.

## Discovered reflection regression

The existing ordinary-text reflection source-bound test failed identically on the prior inner-shadow assembly and this implementation. `textReflectionDirectionDegrees` was registered in both IntegerKinds and ScaledNumberKinds, so the earlier integer branch normalized 45 degrees to 45 instead of the native expected 2700000. Remove this one conflicting integer entry and extend the existing end-to-end regression to fractional initial/edited angles. Keep other leaf contracts unchanged; this repairs the same reflection state lifecycle without weakening source proof.
