# Proposal: add source-bound text soft-edge ownership to PPJ

Expose a bounded `textStyle.softEdge` owner for an existing imported
rich-text run. The owner must preserve the direct DrawingML soft-edge node and
issue one native leaf for its existing radius, so a caller can change that
field without rewriting the surrounding slide.

This slice does not infer paragraph `defaultText` ownership or arbitrary text
effect graphs. Only a direct `a:softEdge`, or one preceded by a separately
verified outer shadow or full-span reflection, is editable; other effects,
unknown children, extensions, and malformed topology remain source-owned or
fail closed.
