# Proposal: add a source-bound shape 3-D contour-width leaf

Imported ordinary shapes can carry a direct `a:sp3d/@contourW` contour width
while the rest of the 3-D scene remains source-owned. PPJ currently has no
field for that proven scalar.

This change will expose the child-free direct coordinate as
`shape3dContourWidthEmu`, preserve all other 3-D state, and allow a
source-bound edit to token-splice only the owning SlidePart's `contourW`
attribute. A focused regression will cover projection, changed-part scope,
Open XML validity, and reprojection.

This does not author or reconstruct `scene3d`, bevels, colors, materials, or
arbitrary 3-D graphs, and it does not change the Office wire version.
