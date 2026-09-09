## Why

F-07 still cannot express direct chart reflections with scale, skew, fade direction or alignment. Such attributes currently make otherwise editable chart text source-owned. Complete the direct reflection object without adding an effect-graph escape hatch.

## What Changes

- Add optional `fadeAngle`, `scaleX`, `scaleY`, `skewX`, `skewY`, `alignment` and `rotateWithShape` to `chartTextStyle.reflection`.
- Preserve independent presence, signed scales, explicit zero/false, and native precision through authored compilation, imported projection, edits, style precedence and vector defaults.
- Retain strict graph/order and ordinary imported reflection guards; extend the existing lifecycle experiment and wire preservation regression.

## Capabilities

### New Capabilities

- `ppj-chart-reflection-transforms`: Complete direct chart reflection transforms and their independent lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema, additive PresentationReflection wire fields and generated bindings; shared reflection codec, chart compiler/projector, existing tests, capability registry, focused reference, coverage and backlog. No production preview or host acceptance claim.
