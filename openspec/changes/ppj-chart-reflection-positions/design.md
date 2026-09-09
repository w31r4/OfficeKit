## Context

See proposal.md. The chart reflection value already carries five optional scalars. Its parser currently shares the strict ordinary full-span reader and its writer always adds start/end positions. Every shared chart text style and vector default uses this value.

## Goals / Non-Goals

Goals: direct chart alpha-ramp positions, independent native presence and source-bound edits. Non-goals: ordinary PPJ position syntax, wider ordinary imported topology, transforms, new native leaves or visual acceptance.

## Decisions

- Add optional uint32 start_position_thousandth_percent/end_position_thousandth_percent at PresentationReflection fields 6/7; chart PPJ numbers use 0..1 and nearest native unit with ties to even. Values are independent; do not impose an undocumented start-before-end constraint.
- Preserve positions in the value reader/writer. Add an explicit chart opt-in to the direct reader for variable/missing positions; ordinary callers retain their existing explicit full-span proof. Validate bounds before conversion.
- Shared writer emits only present positions. Ordinary authored and source-bound reflection builders explicitly set 0/100000, preserving their prior canonical output. Chart builders do not materialize missing positions; chart projection emits only present keys. Vector default propagation already clones the complete value.
- Extend the existing reflection lifecycle to absent, partial, explicit 0/1, reversed and equal positions, attribute removal and whole-effect removal/recreation. Check untouched sibling effects/package bytes. Keep malformed cases malformed by replacing newly supported span fixtures with out-of-bound/unknown-attribute fixtures.

## Risks / Trade-offs

- Ordinary imported guards could accidentally expand: use an opt-in reader parameter and retain all ordinary non-full-span tests.
- Default serialization changes could break ordinary output: set its old constants at both ordinary builders and run the effect suite.
- Terminology could imply clipping: these are positions along the alpha gradient ramp, not slide coordinates or crop geometry.

## Sources

Microsoft's [Reflection reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.reflection?view=openxml-3.0.1) describes start/end positions as positive fixed percentages along the alpha ramp. The upstream [docx4j schema fragment and generated getters](https://raw.githubusercontent.com/plutext/docx4j/master/docx4j-openxml-objects/src/main/java/org/docx4j/dml/CTReflectionEffect.java) give omitted stPos/endPos defaults of 0/100000. Host appearance remains unverified.
