## Why

Imported ordinary shapes can carry a direct `a:sp3d/a:bevelB/@prst` bevel
preset while the rest of the 3-D scene remains source-owned. PPJ currently has
no bounded field for this proven token, leaving the bottom-bevel preset opaque.

## What Changes

- Expose one strict source-bound `shape3dBevelBottomPreset` native leaf.
- Permit a source-bound edit to replace only the owning SlidePart's
  `bevelB/@prst` token.
- Add one minimal projection/edit/reprojection regression.
- Keep source-free 3-D authoring, scene reconstruction, and unrelated bevel or
  effect editing out of scope.

## Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-bevel-bottom-preset`: Project and token-splice one
  strict bottom-bevel preset owner.

### Modified Capabilities

None.

## Impact

- Extends the PPJ protobuf/native leaf model, JSON schema, capability registry,
  presentation Skill reference, and native PPTX projection/edit-plan codec.
- Adds a focused Open XML validity and byte-preservation regression.
- No Office wire-version change, dependency change, relationship change, or
  source-free authoring behavior.
