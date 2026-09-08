# Proposal: add source-bound picture shadow theme-color ownership to PPJ

Expose a bounded `image.shadow.color` owner for an existing imported picture
outer shadow when its direct color is a DrawingML theme token. The owner
issues one native leaf so a caller can change the existing direct
`a:schemeClr/@val` token without rewriting the picture payload, mask, border,
geometry, or other shadow fields.

This slice is limited to a bare canonical theme token. Explicit RGB is
covered by the neighboring `imageShadowColorRgb` slice; missing colors,
transforms, additional effects, unknown children, extensions, malformed
topology, and other owners remain source-owned or fail closed.
