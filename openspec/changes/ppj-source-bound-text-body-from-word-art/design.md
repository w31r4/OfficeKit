# Design

## Owner and boundary

- Owner: `p:sp/p:txBody/a:bodyPr/@fromWordArt`.
- PPJ semantic location: `textBoxStyle.fromWordArt`.
- Native leaf: `textBodyFromWordArt`.
- The PPJ value is boolean; the source-bound token is canonical `0` or `1`.
- The optional `PresentationTextBodyProperties.from_word_art = 36` wire
  field is additive to protocol 2; no protocol version change is required.
- Only one existing direct body-property attribute is exposed. Missing,
  duplicated, malformed, inherited, or extended WordArt/body-property forms
  stay source-owned.
- Editing replaces only `fromWordArt/@value`; it does not change text,
  paragraphs, runs, warp geometry, effects, or other body-property attributes.
  The field records the explicit DrawingML marker and does not promise a
  complete or host-equivalent WordArt result.

## Evidence

The focused test authors a text body with the WordArt marker, removes the
embedded PPJ, and projects one native leaf. A source-bound boolean edit
changes only `ppt/slides/slide1.xml`, keeps text and all other package parts
byte-identical, passes the Open XML validator, and recovers the new value on
second projection. A source without a direct `fromWordArt` attribute does not
issue the leaf.
