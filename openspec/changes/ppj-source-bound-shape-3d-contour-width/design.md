# Design

- Owner: `p:sp/p:spPr/a:sp3d/@contourW`.
- Native leaf: `shape3dContourWidthEmu`.
- The owner must be the single direct `a:sp3d` child, with no child
  elements, only the standard `z`, `extrusionH`, `contourW`, and
  `prstMaterial` attributes, and a canonical non-negative signed 32-bit
  `contourW` token.
- Projection stores the coordinate in the native shape model and emits one
  leaf. The edit plan rechecks the same owner, requires a changed canonical
  value, and token-splices only `contourW`.
- Missing, malformed, extension-bearing, child-bearing, duplicate, or
  out-of-range state remains source-owned and does not receive this leaf.

No source-free authoring, coordinate normalization, scene reconstruction,
effect editing, relationship changes, or descendant transform is included.
