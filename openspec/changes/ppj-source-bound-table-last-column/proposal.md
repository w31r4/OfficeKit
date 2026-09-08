# Proposal: expose source-bound table last-column emphasis

Add `tableLastColumnEmphasis` as an independent PPJ native leaf for the
existing direct `a:tblPr/@lastCol` token on a recognized rectangular table.

The existing `table.style.lastColumnEmphasis` semantic field remains the
authoring and projection path. The native leaf exposes only an existing
canonical boolean attribute, allowing a source-bound token edit without
rebuilding the table or changing its cell topology.
