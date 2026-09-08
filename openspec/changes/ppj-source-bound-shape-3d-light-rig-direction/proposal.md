# Why

Imported ordinary shapes can carry a bounded light-rig direction inside a
direct 3-D scene while the rest of the scene remains unsupported. PPJ already
projects the paired light-rig preset, but still loses this independent
`a:lightRig/@dir` token.

# What Changes

- Expose one strict source-bound `shape3dSceneLightRigDirection` native leaf.
- Permit a source-bound edit to replace only the owning SlidePart's
  `a:lightRig/@dir` token.
- Add one minimal projection/edit/reprojection regression.
- Keep source-free 3-D scene authoring, light rotation, camera editing,
  backdrop editing, and complex scene graphs out of scope.

# Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-light-rig-direction`: Project and token-splice one
  strict 3-D scene light-rig direction.

### Modified Capabilities

None.

# Impact

- Extends the PPJ protobuf/native leaf model, JSON schema, capability registry,
  presentation Skill reference, and native PPTX projection/edit-plan codec.
- Adds a focused Open XML validity and byte-preservation regression.
- No Office wire-version change, dependency change, relationship change, or
  source-free authoring behavior.
