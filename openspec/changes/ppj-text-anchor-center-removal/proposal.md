## Why

F-03/F-06 anchorCenter currently preserves explicit booleans but cannot remove the direct source attribute. Deletion must remain distinct from false and independent of verticalAlignment.

## What Changes

- Add an additive native deletion marker for anchor_center while retaining its existing field number and optional-boolean API.
- Complete PPJ source removal/restoration and guarded simple-style removal across supported text owners and tables.
- Extend shared lifecycle tests, protocol evidence and public field descriptions.

## Capabilities

### New Capabilities
- `ppj-text-anchor-center-removal`: optional anchor-center presence lifecycle.

### Modified Capabilities

None.

## Impact

Proto and generated JS binding, native body validation/application/normalization, PPJ compiler, shared tests, schema/Help/registry and text coverage. Existing serialized true/false values remain unchanged; deletion requires an updated codec.
