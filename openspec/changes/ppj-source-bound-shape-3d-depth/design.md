# Design

- Owner: `p:sp/p:spPr/a:sp3d/@z`.
- Native leaf: `shape3dDepthEmu`.
- The owner must be the single direct `a:sp3d` child, with no child
  elements, only the standard `z`, `extrusionH`, `contourW`, and
  `prstMaterial` attributes, and a canonical signed 32-bit `z` token.
- Projection stores the coordinate in the native shape model and emits one
  leaf. The edit plan rechecks the same owner, requires a changed canonical
  value, and token-splices only the `z` attribute.
- Missing, malformed, extension-bearing, child-bearing, duplicate, or
  out-of-range state remains source-owned and does not receive this leaf.

No source-free authoring, coordinate normalization, scene reconstruction,
effect editing, relationship changes, or descendant transform is included.
