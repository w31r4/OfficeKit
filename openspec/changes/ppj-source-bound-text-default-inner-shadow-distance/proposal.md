# Proposal: add source-bound paragraph default-text inner-shadow distance ownership to PPJ

Expose a bounded `paragraph.style.defaultText.innerShadow.distance` owner
for an existing imported paragraph default run. The owner must preserve the
direct DrawingML `a:defRPr` effect graph and issue one native leaf so a caller
can change that distance without rewriting the surrounding slide.

This slice does not change ordinary run ownership or infer arbitrary paragraph
effect graphs. Only a strict direct `a:innerShdw`, optionally followed by a
separately verified outer shadow, is editable; unknown children, extensions,
malformed topology, and the other inner-shadow scalars remain source-owned or
fail closed.
