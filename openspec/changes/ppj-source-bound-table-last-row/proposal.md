# Proposal: expose source-bound table last-row emphasis

Add `tableLastRow` as an independent PPJ native leaf for the existing direct
`a:tblPr/@lastRow` token on a recognized rectangular table.

The semantic `table.style.lastRow` field is also carried in the additive
`PresentationTable` wire model so authored and projected PPJ preserve the
flag. The native leaf edits only an existing canonical boolean token and does
not rebuild the table or change its cell topology.
