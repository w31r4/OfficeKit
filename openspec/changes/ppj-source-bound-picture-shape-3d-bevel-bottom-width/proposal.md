# Proposal: extend the shape 3-D bottom-bevel width leaf to picture owners

Imported pictures use `p:pic/p:spPr`, which can carry the same direct
DrawingML `a:sp3d/a:bevelB/@w` scalar as an ordinary shape. The existing
`shape3dBevelBottomWidthEmu` source-bound leaf only covers `p:sp`, leaving a
picture bottom-bevel width opaque even when the owner is a single bounded
child.

This change will:

- preserve a proven direct picture bottom-bevel width in the
  `PresentationImage` source model;
- expose it through the existing `shape3dBevelBottomWidthEmu` native leaf on
  the picture;
- allow a source-bound edit to replace only that canonical non-negative
  `bevelB/@w` token in the owning SlidePart;
- add one minimal picture source-bound/reprojection regression.

This does not author or reconstruct picture 3-D markup, top bevels, bottom-
bevel height or preset, bevel colors, scenes, effects, masks, image
relationships, or any other picture properties.
