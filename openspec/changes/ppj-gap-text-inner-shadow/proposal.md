# Proposal: add authored text inner shadows to PPJ

Add a bounded `innerShadow` field to the PPJ text-style profile so authored
text runs and paragraph default runs can own the direct DrawingML
`a:effectLst/a:innerShdw` effect.

The profile reuses the existing inner-shadow color, blur, distance, angle, and
opacity vocabulary. Imported/source-bound text effect graphs, reflection,
WordArt, and arbitrary effect lists remain opaque or fail closed.
