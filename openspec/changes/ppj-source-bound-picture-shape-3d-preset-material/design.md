# Design

- Owner: `p:pic/p:spPr/a:sp3d/@prstMaterial`.
- Native leaf: existing `shape3dPresetMaterial`.
- Source model: add an additive `PresentationImage.shape_3d_preset_material`
  string field. Empty means the source leaf is absent; a non-empty value is
  source-projection-only and is not a source-free image authoring control.
- The picture owner must have exactly one direct child-free `a:sp3d` with a
  canonical token from the existing finite preset-material vocabulary. Other
  direct scalar 3-D attributes may remain source-owned; bevel/scene/color
  children, unknown attributes, malformed values, and duplicate owners reject
  the bounded picture profile.
- Picture projection stores the material token, issues the existing native
  leaf, and keeps the image relationship, crop, mask, effects, and remaining
  3-D XML untouched. The edit plan re-proves the picture owner and token-
  splices only `sp3d/@prstMaterial`.

No source-free authoring, geometry normalization, relationship changes,
effect editing, scene reconstruction, or descendant transform is included.
