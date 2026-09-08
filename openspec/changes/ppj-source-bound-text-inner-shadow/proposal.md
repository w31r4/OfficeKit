# Proposal: add source-bound text inner-shadow ownership to PPJ

Expose a bounded `textStyle.innerShadow` owner for an existing imported
rich-text run. The owner must preserve the direct DrawingML topology and issue
native leaves for the source scalars that are actually present, so a caller
can change one inner-shadow field without rewriting the surrounding slide.

This slice does not infer paragraph `defaultText` ownership or arbitrary text
effect graphs. Lists containing glow, reflection, soft edge, unknown children,
extensions, or malformed inner-shadow colors remain opaque or fail closed.
