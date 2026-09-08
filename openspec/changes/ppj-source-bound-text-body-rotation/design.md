# Design

## Owner and boundary

- Owner: `p:sp/p:txBody/a:bodyPr/@rot`.
- PPJ semantic location: `textBoxStyle.rotation`.
- Native leaf: `textBodyRotationDegrees`.
- Leaf values use the exact signed DrawingML integer in 1/60000 degrees;
  projected JSON values remain degrees for the PPJ surface.
- Only one existing direct `rot` attribute in the bounded signed range
  `-21600000..21600000` is exposed. Missing, duplicated, malformed, or
  unsupported body-property topology remains source-owned.
- Editing replaces only the attribute value. It does not create or remove a
  body property, rewrite paragraphs/runs, or edit the shape frame rotation.

## Evidence

The focused test authors `textBoxStyle.rotation`, compiles a PPTX, validates
the native `rot` token, removes the embedded PPJ, and projects one
`textBodyRotationDegrees` leaf. A source-bound leaf edit changes only
`ppt/slides/slide1.xml`, keeps text content and every other package part byte
identical, passes the Open XML validator, and recovers the new degree value on
second projection.
