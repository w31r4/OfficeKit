# Proposal: add source-bound ordinary-shape shadow rotation ownership to PPJ

Expose a bounded `style.shadow.rotateWithShape` owner for an existing imported
ordinary shape or line outer shadow. The owner must preserve the direct
DrawingML effect graph and issue one native leaf so a caller can change the
existing direct `a:outerShdw/@rotWithShape` token without rewriting the rest of
the shape.

This slice does not infer an omitted attribute or broaden effect topology.
Only a strict direct `a:effectLst/a:outerShdw` with one explicit boolean
`rotWithShape` value is editable; missing attributes, additional effects,
unknown children, extensions, malformed topology, and other owners remain
source-owned or fail closed.
