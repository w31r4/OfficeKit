# Design

## Owner and boundary

- Owner: `p:graphicFrame/a:graphicData/a:tbl/a:tblPr/@firstRow`.
- PPJ semantic location: `table.style.headerRows`.
- Native leaf: `tableHeaderRows`.
- The PPJ value is the integer `0` or `1`; the source-bound token is the same
  canonical `0` or `1`.
- Only one existing direct table-property attribute is exposed. Missing,
  duplicated, malformed, or unsupported table topology remains source-owned.
- Editing replaces only `firstRow/@value`; it does not change cell text,
  geometry, relationships, or other table style flags.

## Evidence

The focused test authors a rectangular table with one header row, removes the
embedded PPJ, and projects one native leaf. A source-bound `1→0` edit changes
only `ppt/slides/slide1.xml`, keeps cell content and all other package parts
byte-identical, passes the Open XML validator, and recovers the new value on
second projection.
