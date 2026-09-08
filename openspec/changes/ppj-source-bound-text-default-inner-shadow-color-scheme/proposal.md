# Proposal: add source-bound paragraph default-text inner-shadow theme ownership to PPJ

Expose a bounded `paragraph.style.defaultText.innerShadow.color` theme-token
owner for an existing imported paragraph default run. The owner must preserve
the direct DrawingML `a:defRPr` effect graph and issue one native leaf so a
caller can change an existing direct `a:schemeClr/@val` without rewriting the
surrounding slide.

This slice does not change RGB-color ownership or infer arbitrary paragraph
effect graphs. Only a strict direct `a:innerShdw` with one direct theme color,
optionally followed by a separately verified outer shadow, is editable;
unknown children, extensions, malformed topology, and other color/alpha
representations remain source-owned or fail closed.
