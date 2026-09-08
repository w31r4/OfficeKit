# Why

The strict source-bound backdrop profile now exposes anchor X and Y, but the
matching Z coordinate is still opaque despite sharing the same complete
literal anchor owner and EMU unit.

# What Changes

- Expose one strict source-bound `shape3dSceneBackdropAnchorZEmu` native leaf
  for `a:scene3d/a:backdrop/a:anchor/@z`.
- Permit a source-bound edit to replace only that owning SlidePart coordinate
  token.
- Add one minimal projection/edit/reprojection regression.
- Keep anchor X/Y, the backdrop normal/up vectors, and wider scene graphs
  source-owned.

# Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-backdrop-anchor-z`: Project and token-splice one
  strict backdrop anchor Z coordinate.

### Modified Capabilities

None.

# Impact

- Extends the PPJ protobuf/native leaf model, JSON schema, capability registry,
  presentation Skill reference, and native PPTX projection/edit-plan codec.
- Adds a focused Open XML validity and byte-preservation regression.
- No Office wire-version change, dependency change, relationship change, or
  source-free authoring behavior.
