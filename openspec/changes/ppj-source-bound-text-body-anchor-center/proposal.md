# Proposal: expose source-bound text-body anchor centering

Add `textBodyAnchorCenter` as an independent PPJ native leaf for the direct
`a:bodyPr/@anchorCtr` token on a recognized text body.

The semantic `textBoxStyle.anchorCenter` field is carried in the additive
`PresentationTextBodyProperties` wire model so authored and projected PPJ
preserve the flag. The native leaf edits only an existing canonical boolean
token and does not alter text topology or other body properties.
