## Why

Imported rich-text runs already preserve direct outer-shadow geometry and scale transforms, but a valid DrawingML `outerShdw/@kx` token still makes the run opaque. This leaves the horizontal shadow skew gap open even though the native shadow model and paragraph-default path already carry bounded skew values.

## What Changes

- Expose direct run `style.shadow.skewX` through the opaque native leaf `textShadowSkewX`.
- Accept only a strict direct-run owner with one existing canonical signed `kx` token, required geometry, one RGB/theme color, and no other shadow transform or sibling effect.
- Compile a source-bound edit by replacing only the proven `outerShdw/@kx` token in the owning SlidePart, preserving the remaining effect XML and ZIP parts.
- Add schema, capability registry, generated reference, backlog/coverage notes, and a focused authored → source-bound → second-projection regression.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-shadow-skew-x`: Direct imported rich-text run outer-shadow horizontal skew with a bounded source-bound token splice.

### Modified Capabilities

- None.

## Impact

- Native presentation projection and source-bound edit planning in `PptxShadowCodec`, `PptxTextCodec`, `PpjNativeLeafProjection`, `PpjPresentationProjector`, and `PptxEditPlanCodec`.
- PPJ schema/registry, generated presentation capability matrix, and presentation Skill references.
- One focused Office Open XML codec test; no protocol version change and no host-rendering claim.
