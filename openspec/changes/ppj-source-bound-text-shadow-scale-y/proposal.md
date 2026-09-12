## Why

Imported rich-text runs can preserve and edit direct outer-shadow geometry and horizontal scale, but a valid DrawingML `outerShdw/@sy` transform is still treated as opaque. That leaves the vertical scale value missing from the P0 text-effect profile even though the native shadow model already carries a bounded signed representation.

## What Changes

- Expose a direct run `style.shadow.scaleY` value through the opaque native leaf `textShadowScaleY`.
- Accept the strict direct-run owner only when `outerShdw/@sy` is an existing canonical signed Int32 token, geometry and one RGB/theme color are valid, and other shadow transforms remain absent.
- Compile a source-bound edit by replacing only the proven `outerShdw/@sy` token in the owning SlidePart, preserving the remaining effect XML and ZIP parts.
- Add schema, capability registry, generated reference, backlog/coverage notes, and a focused authored → source-bound → second-projection regression.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-shadow-scale-y`: Direct imported rich-text run outer-shadow vertical scale with a bounded source-bound token splice.

### Modified Capabilities

- None.

## Impact

- Native presentation projection and source-bound edit planning in `PpjNativeLeafProjection` and `PptxEditPlanCodec`.
- PPJ schema/registry and presentation Skill references.
- One focused Office Open XML codec test; no protocol version change and no host rendering claim.
