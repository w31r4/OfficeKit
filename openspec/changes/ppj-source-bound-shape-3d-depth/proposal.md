# Proposal: add a source-bound shape 3-D depth leaf

Imported ordinary shapes can carry a direct `a:sp3d/@z` depth coordinate even
when the rest of the 3-D scene is intentionally opaque. PPJ currently has no
field for that proven scalar.

This change will:

- expose a child-free direct `a:sp3d/@z` as
  `shape3dDepthEmu` on the shape's `nativeRef.leaves[]`;
- allow a source-bound PPJ edit to replace only that canonical signed
  32-bit coordinate token in the owning SlidePart;
- preserve extrusion height, contour, material, child markup, and all other
  source-owned 3-D state;
- add one minimal negative-coordinate projection/edit/reprojection regression.

This does not author or reconstruct `scene3d`, bevels, colors, materials, or
arbitrary 3-D graphs, and it does not change the Office wire version.
