# Proposal: expose the direct WordArt body hint

Add `textBodyFromWordArt` as an independent PPJ native leaf for the direct
`a:bodyPr/@fromWordArt` token on a recognized text body.

The semantic `textBoxStyle.fromWordArt` field is carried in the additive
`PresentationTextBodyProperties` wire model so authored and projected PPJ
preserve the explicit WordArt marker. This change does not claim support for
WordArt transforms, warp geometry, effects, or host rendering behavior.
