# Proposal: add source-bound picture shadow distance ownership to PPJ

Expose a bounded `image.shadow.distance` owner for an existing imported picture
outer shadow. The owner issues one native leaf so a caller can change the
existing direct `a:outerShdw/@dist` token without rewriting the picture
payload, mask, border, or other shadow fields.

This slice does not infer a missing geometry attribute or broaden picture
effects. Only a strict direct `a:effectLst/a:outerShdw` with one explicit
non-negative canonical integer `dist` value is editable; missing values,
additional effects, unknown children, extensions, malformed topology, and
other owners remain source-owned or fail closed.
