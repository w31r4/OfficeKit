# Design

- Owner: `p:sp/p:spPr/a:sp3d/a:bevelB/@prst`.
- Native leaf: `shape3dBevelBottomPreset`.
- The owner must be the single direct `a:sp3d` child of `a:spPr`, with no
  extension or unrecognized children. Its direct root attributes may be the
  standard `z`, `extrusionH`, `contourW`, and `prstMaterial`; its single
  `a:bevelB` child may carry only the standard `w`, `h`, and `prst` attributes.
  The `prst` value must be in the bounded DrawingML bevel-preset vocabulary.
- Projection stores only the proven preset token in the native shape model and
  emits one leaf. The edit plan rechecks the same child/attribute topology,
  requires a different bounded token, and token-splices only `bevelB/@prst`.
- Missing, malformed, duplicate, extension-bearing, unsupported, or otherwise
  ambiguous bevel state remains source-owned and does not receive this leaf.

No source-free authoring, top-bevel reconstruction, scene/light editing,
relationship changes, or descendant transform is included.
