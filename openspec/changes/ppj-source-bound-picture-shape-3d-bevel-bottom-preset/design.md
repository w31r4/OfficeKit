# Design

- Owner: `p:pic/p:spPr/a:sp3d/a:bevelB/@prst`.
- Native leaf: existing `shape3dBevelBottomPreset`.
- Source model: add an additive presence-aware
  `PresentationImage.shape_3d_bevel_bottom_preset` field. It is
  source-projection-only and is not a source-free image authoring control.
- The picture owner must have exactly one direct `a:sp3d` with exactly one
  direct `a:bevelB` child. The root allows only the established direct 3-D
  scalar attributes and `bevelB` allows only `w`, `h`, and `prst`; the selected
  `prst` must be a recognized bounded DrawingML bevel-preset token. Unknown,
  malformed, duplicate, extension-bearing, extra-child, and other complex
  3-D state remains source-owned.
- Picture projection stores the bevel preset, issues the existing native
  leaf, and keeps the image relationship, crop, mask, effects, and remaining
  3-D XML untouched. The edit plan re-proves the picture owner and
  token-splices only `bevelB/@prst`.

No source-free authoring, geometry normalization, relationship changes,
effect editing, scene reconstruction, or descendant transform is included.
