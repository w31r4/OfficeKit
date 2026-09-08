# Design

## Owner and boundary

- Owner: `p:sp/p:txBody/a:bodyPr/a:flatTx/@z`.
- PPJ semantic location: `textBoxStyle.flatTextZ`.
- Native leaf: `textBodyFlatTextZ`.
- The recognized child is exactly one direct `a:flatTx` with exactly one
  no-namespace `z` attribute, no children, a canonical signed decimal value,
  and a value in the bounded signed 32-bit coordinate range. Duplicate or
  malformed children do not receive a leaf.
- The wire keeps an `int64` so the semantic model does not narrow while the
  bounded PPJ profile remains safe for JSON/JavaScript numeric transport.
- A source-bound edit changes only the digits/sign token inside `z`; it does
  not create/delete/reorder `flatTx` or touch sibling 3D children.

## Evidence

The focused test authors a text shape with `flatTextZ`, checks the direct
`a:flatTx z` XML and Open XML validity, removes embedded PPJ, projects the
semantic field and native leaf, edits the leaf, and verifies a SlidePart-only
change plus byte-identical other ZIP parts. A second projection recovers the
new coordinate. A malformed/extra-attribute source does not issue the leaf.
