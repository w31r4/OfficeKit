## Why

F-04 now expresses custom geometry connectionSites, but connectors cannot name an index: native site bindings project as frame-anchor auto. Connect from/to directly to custom sites while retaining native binding identity and geometry dependencies.

## What Changes

- Add from/to {element, connectionSite} for a zero-based custom-shape site index, distinct from coordinates and frame anchors.
- Resolve declared site values/references through native validation and existing frame/group transforms; emit native target/index bindings.
- Preserve custom site indexes on projection and issue endpoint authority for fully represented custom targets; preset/unsupported site bindings retain their separate source boundary.
- Support source index/target changes, detachment and semantic target geometry/frame dependency updates. Reject dependent geometry native-leaf edits until those can participate in coordinate recomputation.

## Capabilities

### New Capabilities
- `ppj-connector-custom-site-endpoints`: Explicit custom geometry site bindings in PPJ connector endpoints.

### Modified Capabilities
None.

## Impact

Endpoint schema/model, resolver, authored/source compiler, projector/capabilities, focused native regression, preview diagnostics and Help/registry/Skill docs. Existing connector wire target/index fields suffice.
