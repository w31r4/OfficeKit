# Proposal: extend the shape 3-D preset-material leaf to picture owners

Imported pictures use `p:pic/p:spPr`, which can carry the same direct
DrawingML `a:sp3d/@prstMaterial` token as an ordinary shape. The existing
`shape3dPresetMaterial` source-bound leaf only covers `p:sp`, leaving the
picture 3-D material portion of the F-05 image/effect gap opaque even when
the owner is a single child-free token from the bounded vocabulary.

This change will:

- preserve a proven direct picture `a:sp3d/@prstMaterial` in the
  `PresentationImage` source model;
- expose it through the existing `shape3dPresetMaterial` native leaf on the
  picture;
- allow a source-bound edit to replace only that canonical material token in
  the owning SlidePart;
- add one minimal picture source-bound/reprojection regression.

This does not author or reconstruct picture 3-D markup, coordinates, bevels,
scenes, effects, masks, image relationships, or any other picture properties.
