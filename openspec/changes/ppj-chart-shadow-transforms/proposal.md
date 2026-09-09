## Why

F-07 chart typography still loses editability for outer shadows with direct scale or skew attributes. Complete the existing `chartTextStyle.shadow` object so those shadows can be authored, projected and edited independently.

## What Changes

- Add optional signed `scaleX/scaleY` ratios and `skewX/skewY` degree values to chart text shadows.
- Retain native presence, zero, negative scales and precision across chart owners, style precedence and vector defaults.
- Extend the existing shadow lifecycle experiment and wire regression; retain ordinary imported shadow and unknown-graph boundaries.

## Capabilities

### New Capabilities

- `ppj-chart-shadow-transforms`: Direct chart text outer-shadow scaling and skew with independent editing.

### Modified Capabilities

None.

## Impact

PPJ schema, additive PresentationShadow wire fields and generated bindings; native shadow codec, shared chart compiler/projector, existing tests, registry and focused documentation. No runtime dependencies or production preview changes.
