# Proposal: add authored text reflections to PPJ

Add a bounded `reflection` field to the PPJ text-style profile so authored
text runs and paragraph default runs can own the direct DrawingML
`a:effectLst/a:reflection` effect.

The profile reuses the existing reflection blur, start/end opacity, distance,
and angle vocabulary. Imported/source-bound text effect graphs, 3-D, WordArt,
and arbitrary effect lists remain opaque or fail closed.
