# Why

Imported ordinary shapes can carry a bounded preset camera inside a direct
3-D scene while the rest of the scene remains unsupported. PPJ currently has
no field for this proven token, leaving a small but identifiable part of the
PowerPoint 3-D gap opaque.

# What Changes

- Expose one strict source-bound `shape3dSceneCameraPreset` native leaf.
- Permit a source-bound edit to replace only the owning SlidePart's
  `a:camera/@prst` token.
- Add one minimal projection/edit/reprojection regression.
- Keep source-free 3-D scene authoring, camera rotation/zoom/FOV, light-rig
  editing, backdrop editing, and complex scene graphs out of scope.

# Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-camera-preset`: Project and token-splice one
  strict 3-D scene camera preset.

### Modified Capabilities

None.

# Impact

- Extends the PPJ protobuf/native leaf model, JSON schema, capability registry,
  presentation Skill reference, and native PPTX projection/edit-plan codec.
- Adds a focused Open XML validity and byte-preservation regression.
- No Office wire-version change, dependency change, relationship change, or
  source-free authoring behavior.
