# Why

The bounded native leaf `shape3dExtrusionRgb` covers direct RGB extrusion
paint, but a direct DrawingML theme color under `a:sp3d/a:extrusionClr` is
still opaque. A bare `a:schemeClr/@val` is independently bounded and can be
edited without reconstructing the surrounding 3-D graph.

# What Changes

- Add additive source-projection fields for a 3-D extrusion theme token on
  `PresentationShape` and `PresentationImage`.
- Extend one native leaf, `shape3dExtrusionColorScheme`, to strict shape and
  picture owners.
- Project and source-bound-edit only a bare
  `p:sp|p:pic/p:spPr/a:sp3d/a:extrusionClr/a:schemeClr/@val` token.
- Preserve the remaining 3-D, picture, relationship, and package state.
- Add focused shape and picture regressions and refresh the PPJ references.

# Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-extrusion-color-scheme`:
  Bounded projection and source-bound editing of a direct 3-D extrusion theme
  color on ordinary shapes and pictures.

### Modified Capabilities

None.

# Impact

The protobuf source model, generated wire bindings, PPJ capability registry and
reference, presentation projection and edit-plan codecs, and focused Open XML
codec tests are affected. No dependency or protocol-version change is
required.
