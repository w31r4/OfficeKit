## Why

F-03 still lacks an independent lifecycle for paragraph default outer shadows. Required geometry and the ordinary source reader omit optional state and shadow transforms already modeled by chart text.

## What Changes

- Expose all 11 direct shadow fields under ordinary text/shape paragraph defaultText: required color and optional opacity, blur, distance, angle, alignment, rotateWithShape, scaleX/Y and skewX/Y.
- Support assignment, independent optional-field removal/restoration, whole-effect and shadow-only wrapper deletion under exact setTextParagraphStyle authority.
- Preserve RGB/RGBA, source theme identity, grammar color precedence and transformed color resolution. Keep mixed effect XML, neighboring paragraphs and direct runs intact.
- Synchronize schema, Help, focused guidance, generated metadata, preview diagnostics and backlog with original-source regressions.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-shadow-lifecycle`: Full optional direct paragraph outer-shadow values with source-preserving assignment, deletion and restoration.

### Modified Capabilities

None.

## Impact

PPJ paragraph schema, authored/source compiler, projector, field authority and default-run effect patching; shared lifecycle fixtures and Presentation Skill metadata. Existing wire fields and chart shadow values are reused. Ordinary run/shape/image syntax, placeholder inheritance and host rendering retain their current boundaries.
