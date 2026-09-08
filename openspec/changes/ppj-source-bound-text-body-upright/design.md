# Design

## Owner and boundary

- Owner: `p:sp/p:txBody/a:bodyPr/@upright`.
- PPJ semantic location: `textBoxStyle.upright`.
- Native leaf: `textBodyUpright`.
- The PPJ value is boolean; the source-bound token is canonical `0` or `1`.
- Only one existing direct attribute is exposed. Missing, duplicated,
  malformed, or unsupported body-property topology remains source-owned.
- Editing replaces only `upright/@value`; it does not change text, paragraphs,
  runs, AutoFit, overflow, or the shape frame.

## Evidence

The focused test authors `textBoxStyle.upright`, compiles a PPTX, removes the
embedded PPJ, and projects one native leaf. A source-bound boolean edit
changes only `ppt/slides/slide1.xml`, keeps text and all other package parts
byte-identical, passes the Open XML validator, and recovers the new value on
second projection.
