## Why

The PPJ native leaf `shape3dSceneLightRigPreset` already covers strict
imported shapes, but the same complete scene on a picture is still opaque.
The light-rig preset is one bounded token that can be edited while preserving
the camera and the rest of the scene.

## What Changes

- Add an additive `PresentationImage` source-projection field for the light-rig
  preset.
- Extend the existing native leaf and capability registry to direct pictures.
- Project and source-bound-edit only a strict
  `p:pic/p:spPr/a:scene3d/a:lightRig/@rig` token.
- Require the complete canonical camera/light-rig pair and preserve all other
  picture, scene, and package state.
- Add one focused picture regression and refresh the Skill reference/backlog.

## Capabilities

### New Capabilities

- `ppj-source-bound-picture-shape-3d-scene-light-rig-preset`: Bounded
  projection and source-bound editing of a picture light-rig preset.

### Modified Capabilities

None.

## Impact

The protobuf source model, PPJ capability registry/reference, presentation
projection and edit-plan codecs, and focused Open XML codec tests are affected.
No dependency or protocol-version change is required.
