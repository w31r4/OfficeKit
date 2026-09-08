# Why

The native leaf `shape3dSceneBackdropAnchorXEmu` already covers strict
imported shapes, but the same complete backdrop owner on a picture is still
opaque. Its anchor X coordinate is one bounded token that can be edited while
preserving the rest of the scene.

# What Changes

- Add an additive `PresentationImage` source-projection field for backdrop
  anchor X.
- Extend the existing native leaf and capability registry to direct pictures.
- Project and source-bound-edit only a strict
  `p:pic/p:spPr/a:scene3d/a:backdrop/a:anchor/@x` token.
- Require the complete canonical camera/light-rig/backdrop owner and preserve
  all other picture, scene, and package state.
- Add one focused picture regression and refresh the Skill reference/backlog.

# Capabilities

### New Capabilities

- `ppj-source-bound-picture-shape-3d-scene-backdrop-anchor-x`:
  Bounded projection and source-bound editing of a picture backdrop anchor X.

### Modified Capabilities

None.

# Impact

The protobuf source model, PPJ capability registry/reference, presentation
projection and edit-plan codecs, and focused Open XML codec tests are affected.
No dependency or protocol-version change is required.
