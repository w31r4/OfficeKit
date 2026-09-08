# Proposal: add source-bound shape and image glow ownership to PPJ

Expose a bounded source-bound `glow` field for an imported ordinary shape,
line, or picture. The owner must preserve the direct DrawingML glow node and
issue native leaves for its radius, color, and existing explicit alpha, so a
caller can change those scalars without rewriting unrelated package parts.

This slice accepts only one direct `a:glow`, optionally followed by one
separately verified outer shadow. Other effects, unknown children, extensions,
and malformed topology remain source-owned or fail closed. Glow insertion,
removal, and effect-list reordering are outside the source-bound profile.
