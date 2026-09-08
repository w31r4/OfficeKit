# Why

Imported ordinary shapes can carry an explicit camera field-of-view override
in a direct 3-D scene while the rest of the scene remains bounded. PPJ now
keeps the camera preset and zoom override, but still loses the independent
`a:camera/@fov` token.

# What Changes

- Expose one strict source-bound `shape3dSceneCameraFov60000` native leaf.
- Permit a source-bound edit to replace only the owning SlidePart's
  `a:camera/@fov` token.
- Add one minimal projection/edit/reprojection regression.
- Keep source-free 3-D scene authoring, camera rotation, light-rig rotation,
  backdrop editing, and complex scene graphs out of scope.

# Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-camera-fov`: Project and token-splice one strict
  3-D scene camera field-of-view override.

### Modified Capabilities

None.

# Impact

- Extends the PPJ protobuf/native leaf model, JSON schema, capability registry,
  presentation Skill reference, and native PPTX projection/edit-plan codec.
- Adds a focused Open XML validity and byte-preservation regression.
- No Office wire-version change, dependency change, relationship change, or
  source-free authoring behavior.
