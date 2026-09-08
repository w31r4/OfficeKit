# Proposal: add authored text soft edges to PPJ

Add a bounded `softEdge` field to the PPJ text-style profile so authored text
runs and paragraph default runs can own the direct DrawingML
`a:effectLst/a:softEdge` effect.

The profile reuses the existing soft-edge radius vocabulary. Imported/source-
bound text effect graphs, 3-D, WordArt, and arbitrary effect lists remain
opaque or fail closed.
