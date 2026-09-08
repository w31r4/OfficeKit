# Proposal: add authored text glow to PPJ

Add a bounded `glow` field to the PPJ text-style profile so authored text
runs and paragraph default runs can own the direct DrawingML
`a:effectLst/a:glow` effect.

The profile reuses the existing glow color, radius, and opacity vocabulary.
Imported/source-bound text effect graphs, reflection, inner shadow, WordArt,
and arbitrary effect lists remain opaque or fail closed.
