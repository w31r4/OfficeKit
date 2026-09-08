# Proposal: extend the shape 3-D depth leaf to picture owners

Imported pictures use `p:pic/p:spPr`, which can carry the same direct
DrawingML `a:sp3d/@z` depth scalar as an ordinary shape. The existing
`shape3dDepthEmu` source-bound leaf only covers `p:sp`, leaving the picture
3-D depth portion of the F-05 image/effect gap opaque even when the owner is a
single child-free scalar.

This change will:

- preserve a proven direct picture `a:sp3d/@z` in the `PresentationImage`
  source model;
- expose it through the existing `shape3dDepthEmu` native leaf on the picture;
- allow a source-bound edit to replace only that signed canonical `z` token in
  the owning SlidePart;
- add one minimal picture source-bound/reprojection regression.

This does not author or reconstruct picture 3-D markup, bevels, scenes,
effects, masks, image relationships, or any other picture properties.
