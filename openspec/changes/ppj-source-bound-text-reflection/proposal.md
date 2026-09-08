# Proposal: add source-bound text reflection ownership to PPJ

Expose a bounded `textStyle.reflection` owner for an existing imported
rich-text run. The owner must preserve the direct DrawingML reflection
attributes and issue native leaves for the source scalars that are actually
present, so a caller can change one reflection field without rewriting the
surrounding slide.

This slice does not infer paragraph `defaultText` ownership or arbitrary text
effect graphs. Non-full-span reflections, transform variants, soft edge,
glow/inner-shadow combinations, unknown children, and malformed attributes
remain source-owned or fail closed.
