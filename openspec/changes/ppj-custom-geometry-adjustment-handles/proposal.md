## Why

F-04 still cannot express custom geometry adjustment handles in PPJ even though the native codec models XY and polar handles. Expose both bounded records so a shape's adjustment controls survive projection and source edits.

## What Changes

- Add ordered geometry.adjustmentHandles with kind xy/polar, controlled adjustment names, optional paired ranges and required position.
- Preserve numeric local points/degrees and native reference strings through authored export and fresh projection.
- Allow source position/range edits while preserving native handle order, kind and controlled adjustment identity; keep reference-backed path geometry and mask/clip owners outside this shape profile.

## Capabilities

### New Capabilities
- `ppj-custom-geometry-adjustment-handles`: Typed XY and polar custom adjustment controls in PPJ.

### Modified Capabilities
None.

## Impact

PPJ schema, shared shape lowering, projection/source authority, focused native lifecycle regression, preview diagnostics and Help/registry/Skill documentation. Existing wire suffices.
