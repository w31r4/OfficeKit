# Proposal: add a source-bound shape 3-D preset-material leaf

Imported ordinary shapes can carry a direct `a:sp3d/@prstMaterial` token while
the surrounding 3-D scene remains source-owned. PPJ currently has no field for
that proven finite token.

This change will expose recognized child-free direct material tokens as
`shape3dPresetMaterial`, allow a source-bound edit to replace only the owning
SlidePart's `prstMaterial` attribute, and add one minimal projection/edit/
reprojection regression.

This does not author or reconstruct `scene3d`, bevels, colors, arbitrary
materials, or other 3-D graphs, and it does not change the Office wire version.
