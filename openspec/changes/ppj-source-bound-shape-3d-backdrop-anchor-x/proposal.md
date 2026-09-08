# Why

Imported ordinary shapes can carry a complete 3-D backdrop plane, but the
backdrop remains entirely opaque even when its anchor is a direct literal
coordinate. This leaves one small, independently token-spliceable part of the
scene unavailable to PPJ.

# What Changes

- Expose one strict source-bound `shape3dSceneBackdropAnchorXEmu` native leaf
  for `a:scene3d/a:backdrop/a:anchor/@x`.
- Permit a source-bound edit to replace only that owning SlidePart coordinate
  token.
- Add one minimal projection/edit/reprojection regression.
- Keep the backdrop normal/up vectors, the other anchor coordinates, and wider
  scene graphs source-owned.

# Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-backdrop-anchor-x`: Project and token-splice one
  strict backdrop anchor X coordinate.

### Modified Capabilities

None.

# Impact

- Extends the PPJ protobuf/native leaf model, JSON schema, capability registry,
  presentation Skill reference, and native PPTX projection/edit-plan codec.
- Adds a focused Open XML validity and byte-preservation regression.
- No Office wire-version change, dependency change, relationship change, or
  source-free authoring behavior.
