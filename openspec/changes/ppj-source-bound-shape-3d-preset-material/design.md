# Design

- Owner: `p:sp/p:spPr/a:sp3d/@prstMaterial`.
- Native leaf: `shape3dPresetMaterial`.
- The owner must be the single direct `a:sp3d` child, with no child
  elements, only the standard `z`, `extrusionH`, `contourW`, and
  `prstMaterial` attributes, and one exact token from the canonical finite
  preset-material vocabulary.
- Projection stores the token in the native shape model and emits one leaf.
  The edit plan rechecks the same owner, requires a different recognized
  token, and token-splices only `prstMaterial`.
- Missing, malformed, extension-bearing, child-bearing, duplicate, unknown,
  or otherwise ambiguous state remains source-owned and does not receive this
  leaf.

No source-free authoring, material reconstruction, effect editing,
relationship changes, or descendant transform is included.
