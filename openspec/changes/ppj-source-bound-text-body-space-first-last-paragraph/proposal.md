# Proposal: expose source-bound first/last paragraph spacing

Add `textBodySpaceFirstLastParagraph` as an independent PPJ native leaf for
the direct `a:bodyPr/@spcFirstLastPara` token on a recognized text body.

The semantic `textBoxStyle.spaceFirstLastParagraph` field is carried in the
additive `PresentationTextBodyProperties` wire model so authored and
projected PPJ preserve the explicit paragraph-spacing hint. The native leaf
edits only an existing canonical boolean token and does not infer paragraph
metrics or reflow text.
