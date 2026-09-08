# Proposal: expose source-bound table banded rows

Add `tableBandedRows` as an independent PPJ native leaf for the existing
direct `a:tblPr/@bandRow` token on a recognized rectangular table.

The existing `table.style.bandedRows` semantic field remains the authoring and
projection path. The native leaf exposes only an existing canonical boolean
attribute, allowing a source-bound token edit without rebuilding the table or
changing its cell topology.
