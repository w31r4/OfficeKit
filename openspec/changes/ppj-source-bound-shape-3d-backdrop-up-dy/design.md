# Design

- Owner: `p:sp/p:spPr/a:scene3d/a:backdrop/a:up/@dy`.
- Native leaf: `shape3dSceneBackdropUpDyEmu`.
- The owner must be the single direct `a:scene3d` child of `a:spPr`. The scene
  must have exactly three direct children in schema order: a bare recognized
  camera, a bare recognized light rig, and a backdrop. The backdrop must have
  exactly one bare `a:anchor`, one bare `a:norm`, and one bare `a:up`, with all
  three coordinates present as canonical `ST_Coordinate` literals. The up Y
  value is exposed; all other coordinates are validation context only.
- The coordinate uses DrawingML EMUs and the standard bounded signed
  `ST_Coordinate` range. Projection stores only up `dy` in the native shape
  model. The edit plan rechecks the complete backdrop owner and token-splices
  only `up/@dy`, preserving the anchor, normal vector, up `dx/dz`, and all
  other package bytes.
- Missing coordinates, guide/reference values, malformed or out-of-range
  values, extensions, camera/light transforms, duplicate children, and other
  scene state remain source-owned.

No source-free authoring, anchor editing, normal editing, up `dx/dz` editing,
camera/light-rig editing, relationship change, or full 3-D scene
reconstruction is included.
