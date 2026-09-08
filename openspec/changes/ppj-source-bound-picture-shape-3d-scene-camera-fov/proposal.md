## Why

The PPJ native leaf `shape3dSceneCameraFov60000` already covers a strict
imported shape scene, while the equivalent picture FOV remains opaque. The
picture-side gap is a single bounded camera attribute that can be preserved
and edited without owning the rest of the scene graph.

## What Changes

- Add an additive `PresentationImage` source-projection field for picture
  camera FOV.
- Extend the existing native leaf and registry to cover direct picture owners.
- Project and source-bound-edit only a strict
  `p:pic/p:spPr/a:scene3d/a:camera/@fov` token.
- Preserve the picture relationship, crop/mask/effects, other camera and
  scene state, and all non-target package parts.
- Add one focused picture regression and refresh the Skill reference/backlog.

## Capabilities

### New Capabilities

- `ppj-source-bound-picture-shape-3d-scene-camera-fov`: Bounded projection and
  source-bound editing of a picture's 3-D scene camera FOV token.

### Modified Capabilities

None.

## Impact

The protobuf source model, PPJ capability registry/reference, presentation
projection and edit-plan codecs, and focused Open XML codec tests are affected.
No dependency or protocol-version change is required.
