## Why

F-04 cannot yet express native custom-shape connection sites in PPJ. Shapes with these sites become opaque despite the codec already supporting their ordered angle/coordinate values and formula references.

## What Changes

- Add geometry.connectionSites as ordered angle/x/y objects for custom shapes; numeric coordinates use local points and angles degrees, with native reference strings also accepted.
- Preserve authored/fresh projection and allow source value edits at existing indexes under setGeometry.
- Retain native fixed source list length because indexes identify connector bindings; keep handles and referenced paths source-owned.

## Capabilities

### New Capabilities
- `ppj-custom-geometry-connection-sites`: Ordered custom connection-site values and references in PPJ.

### Modified Capabilities
None.

## Impact

Schema, shared custom shape lowering, projector, source field authority, one native lifecycle regression, preview diagnostics and synchronized Help/registry/Skill coverage. Existing wire fields suffice.
