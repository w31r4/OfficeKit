# Proposal: extend the shape 3-D scene camera preset leaf to picture owners

Imported pictures use `p:pic/p:spPr`, which can carry the same direct
DrawingML `a:scene3d/a:camera/@prst` token as an ordinary shape. The existing
`shape3dSceneCameraPreset` source-bound leaf only covers `p:sp`, leaving a
picture camera preset opaque even when the scene is a strict two-child owner.

This change will:

- preserve a proven direct picture camera-preset token in the
  `PresentationImage` source model;
- expose it through the existing `shape3dSceneCameraPreset` native leaf on the
  picture;
- allow a source-bound edit to replace only that canonical camera preset
  token in the owning SlidePart;
- add one minimal picture source-bound/reprojection regression.

This does not author or reconstruct picture 3-D markup, camera transforms,
light rigs, backdrops, effects, masks, image relationships, or any other
picture properties.
