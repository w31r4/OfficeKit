## Why

The PPJ native leaf `shape3dSceneCameraRotationLatitude60000` already covers
strict imported shapes, but the same complete camera rotation on a picture is
still opaque. Latitude is one bounded token that can be edited while preserving
the required longitude/revolution context and the rest of the scene.

## What Changes

- Add an additive `PresentationImage` source-projection field for camera
  rotation latitude.
- Extend the existing native leaf and capability registry to direct pictures.
- Project and source-bound-edit only a strict
  `p:pic/p:spPr/a:scene3d/a:camera/a:rot/@lat` token.
- Require the complete canonical camera rotation and preserve all other
  picture, scene, and package state.
- Add one focused picture regression and refresh the Skill reference/backlog.

## Capabilities

### New Capabilities

- `ppj-source-bound-picture-shape-3d-scene-camera-rotation-latitude`: Bounded
  projection and source-bound editing of picture camera rotation latitude.

### Modified Capabilities

None.

## Impact

The protobuf source model, PPJ capability registry/reference, presentation
projection and edit-plan codecs, and focused Open XML codec tests are affected.
No dependency or protocol-version change is required.
