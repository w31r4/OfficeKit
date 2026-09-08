# Proposal: add source-bound paragraph default-text outer-shadow blur to PPJ

Expose a bounded `paragraph.style.defaultText.shadow.blur` owner for an
existing imported paragraph default run. The owner must preserve the direct
DrawingML `a:defRPr` effect graph and issue one native leaf so a caller can
change an existing direct `a:outerShdw/@blurRad` without rewriting the
surrounding slide.

This slice does not infer omitted geometry or rebuild an effect list. Only a
strict direct `a:outerShdw` with an existing bounded `blurRad` is editable;
additional effects, missing attributes, unknown children, malformed topology,
and other effect graphs remain source-owned or fail closed.
