# Design

- Owner: `p:sp/p:spPr/a:sp3d/a:contourClr/a:srgbClr/@val`.
- Native leaf: `shape3dContourRgb`.
- The owner must be the single direct `a:sp3d` child of `a:spPr`, with no
  other 3-D child, extension, or unknown child. Direct root attributes may be
  `z`, `extrusionH`, `contourW`, and `prstMaterial`; `a:contourClr` has no
  attributes and exactly one child, a bare `a:srgbClr` with only `val`.
  The value is a six-digit hexadecimal RGB token.
- Projection stores only the normalized RGB token in the native shape model
  and emits one leaf. The edit plan rechecks this topology and token-splices
  only the nested `srgbClr/@val` value, preserving the owner and all other
  package bytes.
- Missing, malformed, duplicate, extension-bearing, transformed, alternate
  color-model, or otherwise ambiguous contour color remains source-owned.

No source-free authoring, `a:scene3d`/camera/light editing, bevel
reconstruction, relationship change, or color-transform lowering is included.
