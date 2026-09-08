# Design

## Owner and boundary

- Owner: `p:graphicFrame/a:graphicData/a:tbl/a:tblPr/@lastRow`.
- PPJ semantic location: `table.style.lastRow`.
- Native leaf: `tableLastRow`.
- The PPJ value is boolean; the source-bound token is canonical `0` or `1`.
- The optional `PresentationTable.last_row = 20` wire field is additive to
  protocol 2; no protocol version change is required.
- Only one existing direct table-property attribute is exposed. Missing,
  duplicated, malformed, unsupported, or non-rectangular table topology stays
  source-owned.
- Editing replaces only `lastRow/@value`; it does not change cell text,
  geometry, relationships, or the other table style flags.

## Evidence

The focused test authors a rectangular table with last-row emphasis, removes
the embedded PPJ, and projects one native leaf. A source-bound boolean edit
changes only `ppt/slides/slide1.xml`, keeps cell content and all other package
parts byte-identical, passes the Open XML validator, and recovers the new value
on second projection. A source with no direct `lastRow` attribute does not
issue the leaf.
