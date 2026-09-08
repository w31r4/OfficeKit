# Proposal: expose one shape 3-D extrusion height

F-05 still treats imported shape/image 3-D effects as source-owned. Add one
narrow native leaf for a direct canonical `a:sp3d/@extrusionH` on an ordinary
shape so an agent can change the extrusion height without rebuilding the 3-D
scene or flattening the shape.

## What Changes

- Expose a child-free direct `a:sp3d/@extrusionH` as
  `shape3dExtrusionHeightEmu` on the shape's `nativeRef.leaves[]`.
- Allow a source-bound PPJ edit to replace only that `extrusionH` token in the
  owning SlidePart.
- Preserve depth, contour, material, all surrounding shape properties, and any
  complex 3-D graph as source-owned.
- Add a focused source-bound/reprojection regression and update F-05 coverage.

## Capabilities

### New Capabilities

- `ppj-source-bound-shape-3d-extrusion-height`: a bounded native leaf for one
  direct shape 3-D extrusion-height coordinate.

## Impact

- Native shape projection, edit-plan validation/proof, and SlidePart XML token
  splicing.
- The source model protobuf, capability registry, Skill reference, and PPJ
  gap/coverage documentation.
- Native codec focused tests; no Office wire version change is required.
