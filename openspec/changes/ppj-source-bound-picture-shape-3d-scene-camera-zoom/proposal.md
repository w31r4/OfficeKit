## Why

The PPJ native leaf `shape3dSceneCameraZoomThousandthPercent` already covers a
strict imported shape scene, but an equivalent picture scene camera zoom stays
opaque. This leaves a small, well-bounded gap in the picture-side 3-D profile
and prevents a minimal source-preserving edit of `a:camera/@zoom`.

## What Changes

- Add an additive `PresentationImage` source-projection field for the picture
  camera zoom token.
- Extend the existing native leaf and capability registry to cover both direct
  shape and direct picture owners.
- Project and source-bound-edit only a strict picture
  `p:pic/p:spPr/a:scene3d/a:camera/@zoom` token.
- Preserve the picture relationship, crop, mask, effects, other 3-D scene
  state, and all non-target package parts through token-splice editing.
- Add one focused picture regression and refresh the presentation Skill
  reference/backlog evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-picture-shape-3d-scene-camera-zoom`: Bounded projection and
  source-bound editing of a picture's 3-D scene camera zoom token.

### Modified Capabilities

None.

## Impact

The protobuf source model, PPJ capability registry/reference, presentation
projection and edit-plan codecs, and focused Open XML codec tests are affected.
No new dependency, source-free authoring path, or protocol-version change is
needed.
