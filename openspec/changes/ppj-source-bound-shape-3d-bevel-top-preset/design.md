# Design

- Owner: `p:sp/p:spPr/a:sp3d/a:bevelT/@prst`.
- Native leaf: `shape3dBevelTopPreset`.
- The owner must be the single direct `a:sp3d` child of `a:spPr`, with no
  extension or unrecognized children. Its direct root attributes may be the
  standard `z`, `extrusionH`, `contourW`, and `prstMaterial`; its single
  `a:bevelT` child may carry only the standard `w`, `h`, and `prst` attributes.
  The `prst` value must be one of the finite DrawingML bevel-preset tokens:
  `relaxedInset`, `circle`, `slope`, `cross`, `angle`, `softRound`, `convex`,
  `coolSlant`, `divot`, `riblet`, `hardEdge`, or `artDeco`.
- Projection stores only the proven preset in the native shape model and emits
  one leaf. The edit plan rechecks the same child/attribute topology, requires
  a different finite preset token, and token-splices only `bevelT/@prst`.
- Missing, malformed, duplicate, extension-bearing, unsupported, or otherwise
  ambiguous bevel state remains source-owned and does not receive this leaf.

No source-free authoring, bevel dimension reconstruction, scene/light editing,
relationship changes, or descendant transform is included.
