# Proposal: expose source-bound compatible line spacing

Add `textBodyCompatibleLineSpacing` as an independent PPJ native leaf for the
direct `a:bodyPr/@compatLnSpc` token on a recognized text body.

The semantic `textBoxStyle.compatibleLineSpacing` field is carried in the
additive `PresentationTextBodyProperties` wire model so authored and
projected PPJ preserve the explicit compatibility hint. The native leaf edits
only an existing canonical boolean token and does not infer host line metrics
or reflow text.
