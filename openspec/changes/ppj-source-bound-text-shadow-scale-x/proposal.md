## Why

Imported rich-text runs can already preserve and edit most direct outer-shadow scalars, but a run carrying the valid DrawingML `outerShdw/@sx` transform is currently treated as opaque. That leaves one directly addressable PPJ value missing from the P0 text-effect profile even though the codec already has a bounded signed scale representation.

## What Changes

- Expose a direct run `style.shadow.scaleX` value through the opaque native leaf `textShadowScaleX`.
- Accept the strict direct-run owner only when `outerShdw/@sx` is an existing canonical signed Int32 token, geometry and one RGB/theme color are valid, and other shadow transforms remain absent.
- Compile a source-bound edit by replacing only the proven `outerShdw/@sx` token in the owning SlidePart, preserving the remaining effect XML and ZIP parts.
- Add schema, capability registry, generated reference, backlog/coverage notes, and a focused authored → source-bound → second-projection regression.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-shadow-scale-x`: Direct imported rich-text run outer-shadow horizontal scale with a bounded source-bound token splice.

### Modified Capabilities

- None.

## Impact

- Native presentation projection and source-bound edit planning in `PpjNativeLeafProjection` and `PptxEditPlanCodec`.
- PPJ schema/registry and presentation Skill references.
- One focused Office Open XML codec test; no protocol version change and no host rendering claim.
