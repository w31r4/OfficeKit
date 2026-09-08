# Proposal: add source-bound paragraph default-text soft-edge ownership to PPJ

Expose a bounded `paragraph.style.defaultText.softEdge.radius` owner for an
existing imported paragraph default run. The owner must preserve the direct
DrawingML `a:defRPr` effect graph and issue one native leaf so a caller can
change that radius without rewriting the surrounding slide.

This slice does not change ordinary run ownership or infer arbitrary paragraph
effect graphs. Only a direct `a:softEdge`, or one preceded by a separately
verified outer shadow or full-span reflection, is editable; unknown children,
extensions, malformed topology, and other effects remain source-owned or fail
closed.
