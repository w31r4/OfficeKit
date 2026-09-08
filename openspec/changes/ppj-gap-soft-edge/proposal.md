# Proposal: add authored soft-edge effects to PPJ

Add a bounded `softEdge` field for authored shape styles and image styles so
PPJ can express the direct DrawingML `a:softEdge` owner without pretending to
edit arbitrary imported effect graphs.

This is the smallest remaining F-05 effect slice: one radius value in PPJ
points, native EMU lowering, and a focused shape/picture XML experiment.
Reflection, inner shadow, glow/source-bound effects, 3-D, and complex imported
effect lists remain outside this change.
