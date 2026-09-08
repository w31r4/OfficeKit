# Proposal: add source-bound picture shadow opacity ownership to PPJ

Expose a bounded `image.shadow.opacity` owner for an existing imported picture
outer shadow. The owner issues one native leaf so a caller can change the
existing direct `a:alpha/@val` token without rewriting the picture payload,
mask, border, geometry, or other shadow fields.

This slice does not infer a missing alpha or broaden picture effects. Only a
strict direct `a:effectLst/a:outerShdw` with one direct RGB/theme color and one
explicit alpha value from `0` through `100000` is editable; missing values,
additional effects, unknown children, extensions, malformed topology, and
other owners remain source-owned or fail closed.
