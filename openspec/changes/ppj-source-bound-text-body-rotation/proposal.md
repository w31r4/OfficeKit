# Proposal: expose source-bound text-body rotation

Add `textBodyRotationDegrees` as an independent PPJ native leaf for the
existing direct `a:bodyPr/@rot` token on a recognized text shape.

The field closes one direct imported text-container owner without changing the
wire model: the existing `textBoxStyle.rotation` semantic field remains the
authoring and projection path, while the native leaf carries the exact
DrawingML 60000ths-of-a-degree token needed for a source-bound edit.

The experiment will author a rotated text body, remove the embedded PPJ,
project the native leaf, edit only that leaf, and verify the target slide XML,
text topology, Open XML validity, and second projection.
