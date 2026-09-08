# Proposal: add a source-bound shape 3-D top-bevel preset leaf

Imported ordinary shapes can carry a direct `a:sp3d/a:bevelT/@prst` preset
while the rest of the 3-D scene remains source-owned. PPJ currently has no
field for this finite native bevel choice.

This change exposes a strict top-bevel owner as `shape3dBevelTopPreset`,
allows a source-bound edit to replace only the owning SlidePart's
`bevelT/@prst` token, and adds one minimal projection/edit/reprojection
regression.

This does not author or reconstruct scene3d, bevel dimensions, bottom bevels,
colors, or other 3-D graphs, and it does not change the Office wire version.
