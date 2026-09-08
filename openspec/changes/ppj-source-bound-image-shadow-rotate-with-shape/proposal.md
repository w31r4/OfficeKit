# Proposal: add source-bound picture shadow rotation ownership to PPJ

Expose a bounded `image.shadow.rotateWithShape` owner for an existing imported
picture outer shadow. The owner issues one native leaf so a caller can change
the existing direct `a:outerShdw/@rotWithShape` token without rewriting the
picture payload, mask, border, or other shadow fields.

This slice does not infer an omitted attribute or broaden picture effects.
Only a strict direct `a:effectLst/a:outerShdw` with one explicit boolean
`rotWithShape` value is editable; missing attributes, additional effects,
unknown children, extensions, malformed topology, and other owners remain
source-owned or fail closed.
