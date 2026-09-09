## Why

F-03/F-06 fromWordArt preserves booleans but cannot remove its direct source marker. Deletion must remain distinct from false and must preserve text warp state.

## What Changes

- Add native deletion intent without changing existing setter encoding.
- Complete PPJ source removal/restoration and guarded simple-style deletion for supported text owners and tables.
- Reuse boolean/wire experiments with retained text warp and update public semantics.

## Capabilities

### New Capabilities
- `ppj-text-wordart-marker-removal`: WordArt marker presence lifecycle.

### Modified Capabilities

None.

## Impact

Proto/generated JS, native validation/application/normalization, PPJ compiler, shared tests and field documentation. New deletion requires updated codec; full WordArt rendering remains open.
