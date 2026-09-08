# Proposal: expose source-bound text-body vertical overflow

Add `textBodyVerticalOverflow` as an independent PPJ native leaf for the
existing direct `a:bodyPr/@vertOverflow` token on a recognized text shape.

The existing `textBoxStyle.verticalOverflow` semantic field remains the
authoring and projection path. The native leaf exposes only the finite
DrawingML enum already represented by that field, so a source-bound edit can
replace one token without rebuilding the text body.
