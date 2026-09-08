# Design

## Owner and boundary

- Owner: `p:sp/p:txBody/a:bodyPr/@anchorCtr`.
- PPJ semantic location: `textBoxStyle.anchorCenter`.
- Native leaf: `textBodyAnchorCenter`.
- The PPJ value is boolean; the source-bound token is canonical `0` or `1`.
- The optional `PresentationTextBodyProperties.anchor_center = 32` wire field
  is additive to protocol 2; no protocol version change is required.
- Only one existing direct body-property attribute is exposed. Missing,
  duplicated, malformed, inherited, WordArt, or extended body-property forms
  stay source-owned.
- Editing replaces only `anchorCtr/@value`; it does not change text,
  paragraphs, runs, vertical anchor, or other body-property attributes.

## Evidence

The focused test authors a text body with anchor centering, removes the
embedded PPJ, and projects one native leaf. A source-bound boolean edit
changes only `ppt/slides/slide1.xml`, keeps text and all other package parts
byte-identical, passes the Open XML validator, and recovers the new value on
second projection. A source without a direct `anchorCtr` attribute does not
issue the leaf.
