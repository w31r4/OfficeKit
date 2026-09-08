# Proposal: expose source-bound table header rows

Add `tableHeaderRows` as an independent PPJ native leaf for the existing
direct `a:tblPr/@firstRow` token on a recognized rectangular table.

The existing `table.style.headerRows` semantic field remains the authoring and
projection path. The native leaf exposes the bounded `0|1` header-row value,
allowing a source-bound token edit without rebuilding the table or changing
its cell topology.
