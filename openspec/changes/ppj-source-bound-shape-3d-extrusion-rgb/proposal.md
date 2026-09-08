## Why

Imported ordinary shapes can carry a direct 3-D extrusion color while the
rest of the scene remains unsupported. PPJ currently has no bounded field for
this proven token, leaving another small part of the PowerPoint 3-D gap opaque.

## What Changes

- Expose one strict source-bound `shape3dExtrusionRgb` native leaf.
- Permit a source-bound edit to replace only the owning SlidePart's
  `a:extrusionClr/a:srgbClr/@val` token.
- Add one minimal projection/edit/reprojection regression.
- Keep source-free 3-D authoring, scene reconstruction, bevel editing, and
  transformed or effect-bearing color graphs out of scope.

## Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-extrusion-rgb`: Project and token-splice one strict
  3-D extrusion RGB owner.

### Modified Capabilities

None.

## Impact

- Extends the PPJ protobuf/native leaf model, JSON schema, capability registry,
  presentation Skill reference, and native PPTX projection/edit-plan codec.
- Adds a focused Open XML validity and byte-preservation regression.
- No Office wire-version change, dependency change, relationship change, or
  source-free authoring behavior.
