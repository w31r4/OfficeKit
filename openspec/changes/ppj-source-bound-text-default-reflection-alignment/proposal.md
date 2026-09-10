# Proposal: add source-bound paragraph default-text reflection alignment

Expose the existing imported paragraph default-text reflection alignment as a
complete PPJ field and a source-bound native leaf. The field maps directly to
the DrawingML `a:reflection/@algn` token, so a caller can change alignment
without rewriting the surrounding effect list or slide part.

This is a narrow experiment. It keeps the existing strict full-span reflection
owner, permits the already modeled reflection transforms, and leaves
`rotWithShape`, unknown attributes/children, and complex effect graphs
source-owned or fail closed.
