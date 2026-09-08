# Design

- Owner: `p:sp/p:spPr/a:scene3d/a:backdrop/a:norm/@dx`.
- Native leaf: `shape3dSceneBackdropNormalDxEmu`.
- The owner must be the single direct `a:scene3d` child of `a:spPr`. The scene
  must have exactly three direct children in schema order: a bare recognized
  camera, a bare recognized light rig, and a backdrop. The backdrop must have
  exactly one bare `a:anchor`, one bare `a:norm`, and one bare `a:up`, with all
  three coordinates present as canonical `ST_Coordinate` literals. The normal
  X value is exposed; all other coordinates are validation context only.
- The coordinate uses DrawingML EMUs and the standard bounded signed
  `ST_Coordinate` range. Projection stores only normal `dx` in the native
  shape model. The edit plan rechecks the complete backdrop owner and
  token-splices only `norm/@dx`, preserving the anchor, normal `dy/dz`, up
  vector, and all other package bytes.
- Missing coordinates, guide/reference values, malformed or out-of-range
  values, extensions, camera/light transforms, duplicate children, and other
  scene state remain source-owned.

No source-free authoring, anchor editing, normal `dy/dz` editing, up-vector
editing, camera/light-rig editing, relationship change, or full 3-D scene
reconstruction is included.
