# Proposal: add source-bound shape and image reflection ownership to PPJ

Expose a bounded source-bound `reflection` field for an imported ordinary
shape, line, or picture. The owner must preserve a direct full-span DrawingML
reflection node and issue native leaves for its geometry and opacity values,
so a caller can change those scalars without rewriting unrelated package
parts.

This slice accepts only one direct full-span `a:reflection`, optionally
preceded by one separately verified outer shadow. Other effects, reflection
transforms, unknown children, extensions, and malformed topology remain
source-owned or fail closed. Reflection insertion, removal, and effect-list
reordering are outside the source-bound profile.
