# Proposal: expose source-bound text-body upright

Add `textBodyUpright` as an independent PPJ native leaf for the existing
direct `a:bodyPr/@upright` token on a recognized text shape.

The existing `textBoxStyle.upright` semantic field remains the authoring and
projection path. The native leaf exposes only an existing canonical boolean
attribute, allowing a source-bound token edit without rebuilding text or
changing the shape frame.
