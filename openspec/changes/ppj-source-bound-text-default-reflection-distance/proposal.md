# Proposal: add source-bound paragraph default-text reflection distance ownership to PPJ

Expose a bounded `paragraph.style.defaultText.reflection.distance` owner for
an existing imported paragraph default run. The owner must preserve the direct
DrawingML `a:defRPr` effect graph and issue one native leaf so a caller can
change that distance without rewriting the surrounding slide.

This slice does not change ordinary run ownership or infer arbitrary paragraph
effect graphs. Only a strict full-span direct `a:reflection`, optionally
preceded by a separately verified outer shadow, is editable; unknown children,
extensions, malformed topology, and the other reflection scalars remain
source-owned or fail closed.
