# Proposal: add source-bound picture shadow RGB color ownership to PPJ

Expose a bounded `image.shadow.color` owner for an existing imported picture
outer shadow. The owner issues one native leaf so a caller can change the
existing direct `a:srgbClr/@val` token without rewriting the picture payload,
mask, border, geometry, or other shadow fields.

This slice is deliberately limited to explicit RGB. Only a strict direct
`a:effectLst/a:outerShdw` with one direct RGB color token is editable; theme
colors, missing colors, additional effects, unknown children, extensions,
malformed topology, and other owners remain source-owned or fail closed.
