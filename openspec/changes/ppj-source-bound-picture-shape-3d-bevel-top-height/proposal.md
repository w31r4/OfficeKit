# Proposal: extend the shape 3-D top-bevel height leaf to picture owners

Imported pictures use `p:pic/p:spPr`, which can carry the same direct
DrawingML `a:sp3d/a:bevelT/@h` scalar as an ordinary shape. The existing
`shape3dBevelTopHeightEmu` source-bound leaf only covers `p:sp`, leaving a
picture top-bevel height opaque even when the owner is a single bounded child.

This change will:

- preserve a proven direct picture top-bevel height in the
  `PresentationImage` source model;
- expose it through the existing `shape3dBevelTopHeightEmu` native leaf on the
  picture;
- allow a source-bound edit to replace only that canonical non-negative
  `bevelT/@h` token in the owning SlidePart;
- add one minimal picture source-bound/reprojection regression.

This does not author or reconstruct picture 3-D markup, top-bevel width or
preset, bottom bevels, bevel colors, scenes, effects, masks, image
relationships, or any other picture properties.
