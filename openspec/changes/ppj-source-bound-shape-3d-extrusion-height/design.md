# Design

## Owner and boundary

- Owner: `p:sp/p:spPr/a:sp3d/@extrusionH`.
- PPJ surface: `shape.nativeRef.leaves[]`.
- Native leaf: `shape3dExtrusionHeightEmu`.
- The owner must be the single direct `a:sp3d` child, with no child elements,
  only the standard `z`, `extrusionH`, `contourW`, and `prstMaterial`
  attributes, and a canonical non-negative bounded `extrusionH` token.

## Edit proof

The codec rechecks the direct `a:sp3d` owner, child-free topology, allowed
attribute set, canonical current coordinate, and changed canonical requested
coordinate against the exact source SlidePart. The writer then replaces only
the `extrusionH` attribute value token.

## Non-goals

- No public 3-D scene, camera, light-rig, bevel, extrusion color, contour color,
  material, or arbitrary effect graph model.
- No picture 3-D owner, shape insertion/deletion, or topology changes.
- No synthesized missing `sp3d`, no coordinate normalization, and no edits to
  depth, contour, material, or other shape-property children.

## Evidence

The focused regression adds a direct child-free `a:sp3d` owner to an imported
ordinary shape, removes embedded PPJ, projects the extrusion-height leaf,
changes it, checks that only the owning SlidePart changed and the package stays
valid, then projects again to recover the new height.
