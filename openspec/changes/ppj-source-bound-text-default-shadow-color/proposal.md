# Proposal: add source-bound paragraph default-text outer-shadow RGB ownership to PPJ

Expose a bounded `paragraph.style.defaultText.shadow.color` RGB owner for an
existing imported paragraph default run. The owner must preserve the direct
DrawingML `a:defRPr` effect graph and issue one native leaf so a caller can
change an existing direct `a:outerShdw/a:srgbClr/@val` without rewriting the
surrounding slide.

This slice does not change theme-color ownership or infer arbitrary paragraph
effect graphs. Only a strict direct `a:outerShdw` with one direct RGB color is
editable; unknown children, extensions, malformed topology, and other
color/alpha representations remain source-owned or fail closed.
