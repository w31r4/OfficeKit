# Proposal: expose source-bound text-body horizontal overflow

Add `textBodyHorizontalOverflow` as an independent PPJ native leaf for the
existing direct `a:bodyPr/@horzOverflow` token on a recognized text shape.

The existing `textBoxStyle.horizontalOverflow` semantic field remains the
authoring and projection path. The native leaf exposes only the finite
DrawingML enum already represented by that field, allowing a source-bound
single-token edit without rebuilding the text body.
