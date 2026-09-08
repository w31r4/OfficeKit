# Design

## Owner and boundary

- Owner: `p:sp/p:txBody/a:bodyPr/@vertOverflow`.
- PPJ semantic location: `textBoxStyle.verticalOverflow`.
- Native leaf: `textBodyVerticalOverflow`.
- Values are the exact bounded tokens `overflow`, `ellipsis`, or `clip`.
- Only one existing direct attribute is exposed. Missing, duplicated,
  malformed, or unsupported body-property topology remains source-owned.
- Editing replaces only `vertOverflow/@value`; it does not change text,
  paragraphs, runs, AutoFit, or the shape frame.

## Evidence

The focused test authors `textBoxStyle.verticalOverflow`, compiles a PPTX,
removes the embedded PPJ, and projects one native leaf. A source-bound leaf
edit changes only `ppt/slides/slide1.xml`, keeps text and all other package
parts byte-identical, passes the Open XML validator, and recovers the new enum
on second projection.
