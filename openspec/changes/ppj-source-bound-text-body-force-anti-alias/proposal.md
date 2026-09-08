# Proposal: expose source-bound text-body force anti-aliasing

Add `textBodyForceAntiAlias` as an independent PPJ native leaf for the direct
`a:bodyPr/@forceAA` token on a recognized text body.

The semantic `textBoxStyle.forceAntiAlias` field is carried in the additive
`PresentationTextBodyProperties` wire model so authored and projected PPJ
preserve the explicit rendering hint. The native leaf edits only an existing
canonical boolean token and does not claim a host-specific visual result.
