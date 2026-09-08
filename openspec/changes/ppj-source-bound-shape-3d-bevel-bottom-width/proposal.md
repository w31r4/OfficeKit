# Proposal: add a source-bound shape 3-D bottom-bevel width leaf

Imported ordinary shapes can carry a direct `a:sp3d/a:bevelB/@w` width while
the rest of the 3-D scene remains source-owned. PPJ currently has no field for
this proven numeric bevel parameter.

This change exposes a strict bottom-bevel owner as
`shape3dBevelBottomWidthEmu`, allows a source-bound edit to replace only the
owning SlidePart's `bevelB/@w` token, and adds one minimal
projection/edit/reprojection regression.

This does not author or reconstruct scene3d, top bevels, bevel presets, colors,
or other 3-D graphs, and it does not change the Office wire version.
