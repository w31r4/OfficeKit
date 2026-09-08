# Proposal: add source-bound paragraph default-text outer-shadow rotation ownership to PPJ

Expose a bounded `paragraph.style.defaultText.shadow.rotateWithShape` owner
for an existing imported paragraph default run. The owner must preserve the
direct DrawingML `a:defRPr` effect graph and issue one native leaf so a caller
can change the existing direct `a:outerShdw/@rotWithShape` token without
rewriting the surrounding slide.

This slice does not infer an omitted attribute or arbitrary paragraph effect
graphs. Only a strict direct `a:outerShdw` with one explicit boolean
`rotWithShape` value is editable; missing attributes, unknown children,
extensions, malformed topology, and other effect graphs remain source-owned or
fail closed.
