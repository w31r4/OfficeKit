# Proposal: add source-bound paragraph default-text glow alpha to PPJ

Expose a bounded `paragraph.style.defaultText.glow.opacity` owner for an
existing imported paragraph default run. The owner must preserve the direct
DrawingML `a:defRPr` effect graph and issue one native leaf so a caller can
change an existing direct color `a:alpha/@val` without rewriting the
surrounding slide.

This slice does not infer implicit opacity or create a missing alpha child.
Only a strict direct `a:glow` whose direct RGB or theme color contains one
explicit bounded alpha, optionally followed by a separately verified outer
shadow, is editable; missing alpha, duplicate alpha, unknown children,
malformed topology, and other effect graphs remain source-owned or fail
closed.
