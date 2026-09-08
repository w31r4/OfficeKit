# Why

The native leaf `shape3dSceneBackdropUpDzEmu` already covers strict imported
shapes, but the same complete backdrop owner on a picture is still opaque. Its
up-vector Z coordinate is an independently bounded literal that can be edited
without reconstructing the scene.

# What Changes

- Add an additive `PresentationImage` source-projection field for backdrop
  up-vector Z.
- Extend the existing native leaf and capability registry to direct pictures.
- Project and source-bound-edit only a strict
  `p:pic/p:spPr/a:scene3d/a:backdrop/a:up/@dz` token.
- Require the complete canonical camera/light-rig/backdrop owner and preserve
  all other picture, scene, and package state.
- Add one focused picture regression and refresh the Skill reference/backlog.

# Capabilities

### New Capabilities

- `ppj-source-bound-picture-shape-3d-scene-backdrop-up-dz`:
  Bounded projection and source-bound editing of a picture backdrop up-vector Z.

### Modified Capabilities

None.

# Impact

The protobuf source model, PPJ capability registry/reference, presentation
projection and edit-plan codecs, and focused Open XML codec tests are affected.
No dependency or protocol-version change is required.
