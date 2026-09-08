# Why

The strict source-bound backdrop profile now exposes the anchor X coordinate,
but the matching Y coordinate is still opaque even though it has the same
independent literal owner and unit.

# What Changes

- Expose one strict source-bound `shape3dSceneBackdropAnchorYEmu` native leaf
  for `a:scene3d/a:backdrop/a:anchor/@y`.
- Permit a source-bound edit to replace only that owning SlidePart coordinate
  token.
- Add one minimal projection/edit/reprojection regression.
- Keep anchor X/Z, the backdrop normal/up vectors, and wider scene graphs
  source-owned.

# Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-backdrop-anchor-y`: Project and token-splice one
  strict backdrop anchor Y coordinate.

### Modified Capabilities

None.

# Impact

- Extends the PPJ protobuf/native leaf model, JSON schema, capability registry,
  presentation Skill reference, and native PPTX projection/edit-plan codec.
- Adds a focused Open XML validity and byte-preservation regression.
- No Office wire-version change, dependency change, relationship change, or
  source-free authoring behavior.
