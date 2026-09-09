## Why

F-07 chart typography cannot express direct text outer shadows even though ordinary PPJ text already has a shadow model. Imported chart character effects consequently remain source-owned.

## What Changes

- Add `chartTextStyle.shadow` with color and optional blur, distance, angle, opacity, alignment and rotateWithShape; preserve omitted attributes separately from explicit zero/false.
- Carry the existing PresentationShadow wire message through chart typography, native character effects, projection, source-bound editing, style precedence and generated vector text.
- Recognize one direct outer shadow only; preserve unknown effect graphs without granting edits.
- Add a small lifecycle regression and update field documentation/coverage.

## Capabilities

### New Capabilities
- `ppj-chart-text-shadow`: Direct chart text outer-shadow authoring and source-bound lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema, additive protobuf field, shared DrawingML shadow/style codecs, authored/source-bound/projector/vector mapping, JS preservation test and generated references. No root runtime dependency, Office wire version change or SVG/host acceptance claim.
