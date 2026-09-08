# Proposal: add authored inner-shadow effects to PPJ

Add a bounded `innerShadow` field for authored shape styles and image styles so
PPJ can express the direct DrawingML `a:innerShdw` owner.

The profile reuses the existing shadow vocabulary for color, blur, distance,
angle, and opacity, and proves one shape/picture XML round trip. Imported
effect graphs, reflection, 3-D, and other unsupported effects remain opaque.
