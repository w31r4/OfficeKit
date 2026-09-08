# Proposal: extend the shape 3-D bottom-bevel preset leaf to picture owners

Imported pictures use `p:pic/p:spPr`, which can carry the same direct
DrawingML `a:sp3d/a:bevelB/@prst` token as an ordinary shape. The existing
`shape3dBevelBottomPreset` source-bound leaf only covers `p:sp`, leaving a
picture bottom-bevel preset opaque even when the owner is a single bounded
child.

This change will:

- preserve a proven direct picture bottom-bevel preset in the
  `PresentationImage` source model;
- expose it through the existing `shape3dBevelBottomPreset` native leaf on
  the picture;
- allow a source-bound edit to replace only that canonical bevel preset token
  in the owning SlidePart;
- add one minimal picture source-bound/reprojection regression.

This does not author or reconstruct picture 3-D markup, top bevels,
bottom-bevel dimensions, bevel colors, scenes, effects, masks, image
relationships, or any other picture properties.
