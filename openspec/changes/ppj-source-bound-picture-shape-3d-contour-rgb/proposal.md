# Proposal: extend the shape 3-D contour color leaf to picture owners

Imported pictures use `p:pic/p:spPr`, which can carry the same direct
DrawingML `a:sp3d/a:contourClr/a:srgbClr/@val` token as an ordinary shape.
The existing `shape3dContourRgb` source-bound leaf only covers `p:sp`,
leaving a picture contour color opaque even when the owner is a single
bounded child.

This change will:

- preserve a proven direct picture contour RGB token in the
  `PresentationImage` source model;
- expose it through the existing `shape3dContourRgb` native leaf on the
  picture;
- allow a source-bound edit to replace only that canonical color token in
  the owning SlidePart;
- add one minimal picture source-bound/reprojection regression.

This does not author or reconstruct picture 3-D markup, extrusion colors,
bevels, scenes, effects, masks, image relationships, or any other picture
properties.
