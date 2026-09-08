# Why

The strict source-bound backdrop profile exposes all anchor coordinates, but
the first coordinate of the backdrop normal vector is still opaque despite
sharing the same complete literal vector owner and EMU unit.

# What Changes

- Expose one strict source-bound `shape3dSceneBackdropNormalDxEmu` native leaf
  for `a:scene3d/a:backdrop/a:norm/@dx`.
- Permit a source-bound edit to replace only that owning SlidePart coordinate
  token.
- Add one minimal projection/edit/reprojection regression.
- Keep the backdrop anchor, remaining normal coordinates, up vector, and wider
  scene graphs source-owned.

# Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-backdrop-normal-dx`: Project and token-splice one
  strict backdrop normal X coordinate.

### Modified Capabilities

None.

# Impact

- Extends the PPJ protobuf/native leaf model, JSON schema, capability registry,
  presentation Skill reference, and native PPTX projection/edit-plan codec.
- Adds a focused Open XML validity and byte-preservation regression.
- No Office wire-version change, dependency change, relationship change, or
  source-free authoring behavior.
