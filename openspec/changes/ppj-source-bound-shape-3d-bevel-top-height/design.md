# Design

- Owner: `p:sp/p:spPr/a:sp3d/a:bevelT/@h`.
- Native leaf: `shape3dBevelTopHeightEmu`.
- The owner must be the single direct `a:sp3d` child of `a:spPr`, with no
  extension or unrecognized children. Its direct root attributes may be the
  standard `z`, `extrusionH`, `contourW`, and `prstMaterial`; its single
  `a:bevelT` child may carry only the standard `w`, `h`, and `prst` attributes.
  The `h` value must be a canonical non-negative 32-bit DrawingML coordinate.
- Projection stores only the proven height in the native shape model and emits
  one leaf. The edit plan rechecks the same child/attribute topology, requires
  a different canonical height, and token-splices only `bevelT/@h`.
- Missing, malformed, duplicate, extension-bearing, unsupported, or otherwise
  ambiguous bevel state remains source-owned and does not receive this leaf.

No source-free authoring, bevel preset reconstruction, scene/light editing,
relationship changes, or descendant transform is included.
