# Proposal: extend the shape 3-D top-bevel width leaf to picture owners

Imported pictures use `p:pic/p:spPr`, which can carry the same direct
DrawingML `a:sp3d/a:bevelT/@w` scalar as an ordinary shape. The existing
`shape3dBevelTopWidthEmu` source-bound leaf only covers `p:sp`, leaving a
picture top-bevel width opaque even when the owner is a single bounded child.

This change will:

- preserve a proven direct picture top-bevel width in the
  `PresentationImage` source model;
- expose it through the existing `shape3dBevelTopWidthEmu` native leaf on the
  picture;
- allow a source-bound edit to replace only that canonical non-negative
  `bevelT/@w` token in the owning SlidePart;
- add one minimal picture source-bound/reprojection regression.

This does not author or reconstruct picture 3-D markup, bottom bevels, bevel
colors, scenes, effects, masks, image relationships, or any other picture
properties.
