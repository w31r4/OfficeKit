# Design

## Owner and boundary

- Owner: `p:graphicFrame/a:graphicData/a:tbl/a:tblPr/@bandRow`.
- PPJ semantic location: `table.style.bandedRows`.
- Native leaf: `tableBandedRows`.
- The PPJ value is boolean; the source-bound token is canonical `0` or `1`.
- Only one existing direct table-property attribute is exposed. Missing,
  duplicated, malformed, or unsupported table topology remains source-owned.
- Editing replaces only `bandRow/@value`; it does not change cell text,
  geometry, relationships, or other table style flags.

## Evidence

The focused test authors a rectangular table with `bandedRows`, removes the
embedded PPJ, and projects one native leaf. A source-bound boolean edit
changes only `ppt/slides/slide1.xml`, keeps cell content and all other package
parts byte-identical, passes the Open XML validator, and recovers the new
value on second projection.
