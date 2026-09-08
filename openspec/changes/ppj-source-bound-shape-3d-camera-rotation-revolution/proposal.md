# Why

Imported ordinary shapes can carry a camera rotation override in a direct 3-D
scene while the rest of the scene remains unsupported. PPJ now projects the
strict camera rotation latitude and longitude, but the matching revolution
token is still opaque.

# What Changes

- Expose one strict source-bound `shape3dSceneCameraRotationRevolution60000`
  native leaf for the direct `a:camera/a:rot/@rev` value.
- Permit a source-bound edit to replace only the owning SlidePart's
  `a:camera/a:rot/@rev` token.
- Add one minimal projection/edit/reprojection regression.
- Keep camera latitude/longitude, light-rig rotation, backdrop editing, and
  complex scene graphs out of scope.

# Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-camera-rotation-revolution`: Project and
  token-splice one strict camera rotation revolution.

### Modified Capabilities

None.

# Impact

- Extends the PPJ protobuf/native leaf model, JSON schema, capability registry,
  presentation Skill reference, and native PPTX projection/edit-plan codec.
- Adds a focused Open XML validity and byte-preservation regression.
- No Office wire-version change, dependency change, relationship change, or
  source-free authoring behavior.
