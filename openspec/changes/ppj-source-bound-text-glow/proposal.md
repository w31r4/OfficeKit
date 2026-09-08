# Proposal: add a bounded source-bound text glow owner to PPJ

Expose an existing direct DrawingML `a:glow` on an imported text run as
`textStyle.glow`, and issue scalar native leaves for its radius, color, and
opacity. Token-splice edits stay inside the owning SlidePart.

Only one strict glow owner, optionally followed by one proven outer shadow, is
in scope. Other effect-list graphs remain source-owned or fail closed.
