# Why

Imported ordinary shapes can carry a complete light-rig rotation override in
a direct 3-D scene while the rest of the scene remains unsupported. PPJ now
projects the bare light-rig preset and direction, but a light-rig with its
`a:rot` child is still opaque.

# What Changes

- Expose one strict source-bound `shape3dSceneLightRigRotationLatitude60000`
  native leaf for the direct `a:rot/@lat` value.
- Permit a source-bound edit to replace only the owning SlidePart's
  `a:lightRig/a:rot/@lat` token.
- Add one minimal projection/edit/reprojection regression.
- Keep camera overrides, light longitude/revolution, backdrop editing, and
  complex scene graphs out of scope.

# Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-light-rig-rotation-latitude`: Project and
  token-splice one strict light-rig rotation latitude.

### Modified Capabilities

None.

# Impact

- Extends the PPJ protobuf/native leaf model, JSON schema, capability registry,
  presentation Skill reference, and native PPTX projection/edit-plan codec.
- Adds a focused Open XML validity and byte-preservation regression.
- No Office wire-version change, dependency change, relationship change, or
  source-free authoring behavior.
