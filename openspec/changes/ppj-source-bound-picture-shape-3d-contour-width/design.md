# Design

- Owner: `p:pic/p:spPr/a:sp3d/@contourW`.
- Native leaf: existing `shape3dContourWidthEmu`.
- Source model: add an additive `PresentationImage.shape_3d_contour_width_emu`
  presence-aware field. It is source-projection-only and is not a source-free
  image authoring control.
- The picture owner must have exactly one direct child-free `a:sp3d` with a
  canonical non-negative signed 32-bit `contourW` token. Other direct scalar
  3-D attributes may remain source-owned; bevel/scene/color children, unknown
  attributes, malformed values, and duplicate owners reject the bounded
  picture profile.
- Picture projection stores the contour width, issues the existing native
  leaf, and keeps the image relationship, crop, mask, effects, and remaining
  3-D XML untouched. The edit plan re-proves the picture owner and token-
  splices only `sp3d/@contourW`.

No source-free authoring, geometry normalization, relationship changes,
effect editing, scene reconstruction, or descendant transform is included.
