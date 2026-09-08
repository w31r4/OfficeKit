# Why

Imported ordinary shapes can carry a complete light-rig rotation override in
a direct 3-D scene while the rest of the scene remains unsupported. PPJ now
projects the strict light-rig rotation latitude, but the matching longitude
token is still opaque.

# What Changes

- Expose one strict source-bound `shape3dSceneLightRigRotationLongitude60000`
  native leaf for the direct `a:rot/@lon` value.
- Permit a source-bound edit to replace only the owning SlidePart's
  `a:lightRig/a:rot/@lon` token.
- Add one minimal projection/edit/reprojection regression.
- Keep camera overrides, light latitude/revolution, backdrop editing, and
  complex scene graphs out of scope.

# Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-light-rig-rotation-longitude`: Project and
  token-splice one strict light-rig rotation longitude.

### Modified Capabilities

None.

# Impact

- Extends the PPJ protobuf/native leaf model, JSON schema, capability registry,
  presentation Skill reference, and native PPTX projection/edit-plan codec.
- Adds a focused Open XML validity and byte-preservation regression.
- No Office wire-version change, dependency change, relationship change, or
  source-free authoring behavior.
