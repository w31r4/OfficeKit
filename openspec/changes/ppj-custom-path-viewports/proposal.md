## Why

F-04 still rejects recognized custom shapes whose paths use different or default native coordinate extents. PPJ needs one per-path viewport override so those paths retain their own coordinate spaces.

## What Changes

- Add optional geometry.paths[].viewport {width,height}; omitted inherits geometry.viewBox, zero uses the native default independently per axis.
- Preserve heterogeneous/default extents through authored export, projection and source path edits.
- Keep numeric command units unchanged and mask/clip viewport overrides outside this shape profile.

## Capabilities

### New Capabilities
- `ppj-custom-path-viewports`: Per-path custom-shape coordinate extents and native default axes.

### Modified Capabilities
None.

## Impact

Schema, custom path lowering, projection/source geometry gates, one focused native lifecycle regression and preview diagnostics; synchronized Help/registry/Skill documentation. Native wire already represents positive/zero extents.
