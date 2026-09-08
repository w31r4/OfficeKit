# Design

- Owner: `p:pic/p:spPr/a:sp3d/a:contourClr/a:srgbClr/@val`.
- Native leaf: existing `shape3dContourRgb`.
- Source model: add an additive `PresentationImage.shape_3d_contour_rgb`
  field. It is source-projection-only and is not a source-free image
  authoring control.
- The picture owner must have exactly one direct `a:sp3d` with exactly one
  direct `a:contourClr` child and exactly one direct `a:srgbClr` child. The
  root allows only the established direct 3-D scalar attributes; the color
  owner has no attributes or extra children; the selected `val` is a bare
  six-digit RGB token. Unknown, malformed, duplicate, extension-bearing,
  extra-child, and other complex 3-D state remains source-owned.
- Picture projection stores the color token, issues the existing native leaf,
  and keeps the image relationship, crop, mask, effects, and remaining 3-D
  XML untouched. The edit plan re-proves the picture owner and token-splices
  only `a:srgbClr/@val`.

No source-free authoring, geometry normalization, relationship changes,
effect editing, scene reconstruction, or descendant transform is included.
