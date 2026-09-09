## Why

F-03/F-06 forceAntiAlias preserves booleans but cannot remove a direct source hint. False and absence must remain distinct.

## What Changes

- Add a native deletion marker without changing existing boolean encoding.
- Complete PPJ source removal/restoration and guarded simple-style removal across supported owners.
- Reuse boolean/wire experiments and update field semantics and preview limits.

## Capabilities

### New Capabilities
- `ppj-text-force-antialias-removal`: optional anti-alias hint presence lifecycle.

### Modified Capabilities

None.

## Impact

Proto/generated JS, native validation/application/normalization, PPJ compiler, lifecycle tests and public field documentation. New deletion requires updated codec; no host rendering claim.
