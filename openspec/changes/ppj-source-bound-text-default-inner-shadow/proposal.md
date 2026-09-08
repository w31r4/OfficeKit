# Proposal: add source-bound paragraph default-text inner-shadow blur ownership to PPJ

Expose a bounded `paragraph.style.defaultText.innerShadow.blur` owner for an
existing imported paragraph default run. The owner must preserve the direct
DrawingML `a:defRPr` effect graph and issue one native leaf so a caller can
change that blur radius without rewriting the surrounding slide.

This slice does not change ordinary run ownership or infer arbitrary paragraph
effect graphs. Only a direct `a:innerShdw`, or one followed by a separately
verified outer shadow, is editable; unknown children, extensions, malformed
topology, and other inner-shadow scalars remain source-owned or fail closed.
