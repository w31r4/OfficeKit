## Why

F-03 still lacks an independently editable paragraph default reflection. The current required five-value authoring profile loses optional source state and omits reflection positions and transforms, although the codec already models those values for chart text.

## What Changes

- Expose the full existing reflection value model under ordinary text/shape paragraph defaultText: optional blur, distance, angle, start/end opacity, start/end position, fadeAngle, scaleX/Y, skewX/Y, alignment and rotateWithShape.
- Support assignment, independent scalar removal/restoration, whole-effect and single-effect-wrapper deletion with exact setTextParagraphStyle authority; an empty object retains the effect with native defaults.
- Preserve neighboring defaults, direct runs, mixed effects and source XML/ZIP. Unmodeled source effect graphs remain source-owned and reject replacement.
- Synchronize Help, schema, focused guidance, generated metadata, preview diagnostics and backlog evidence with a small set of original-source regressions.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-reflection-lifecycle`: Full optional reflection values and independent source-bound lifecycle for direct paragraph defaults.

### Modified Capabilities

None.

## Impact

PPJ paragraph-default schema, authored/source compiler, projector, semantic authority and default-run codec; shared lifecycle fixtures, preview input assessment and Presentation Skill metadata. Reuses the existing wire fields and effect writer. Ordinary run/shape/image profiles, placeholder inheritance and host rendering retain their current boundaries.
