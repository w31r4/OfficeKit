# Proposal: add source-bound paragraph default-text glow-radius ownership to PPJ

Expose a bounded `paragraph.style.defaultText.glow.radius` owner for an
existing imported paragraph default run. The owner must preserve the direct
DrawingML `a:defRPr` effect graph and issue one native leaf so a caller can
change that radius without rewriting the surrounding slide.

This slice does not change ordinary run ownership or infer arbitrary paragraph
effect graphs. Only a direct `a:glow`, or one followed by a separately verified
outer shadow, is editable; unknown children, extensions, malformed topology,
and other effects remain source-owned or fail closed. Glow color and alpha,
insertion, removal, and reordering remain outside this leaf.
