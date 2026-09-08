# Proposal: author a bounded glow effect

Add a typed PPJ `glow` effect for authored shape styles and picture styles.
The field lowers to the direct DrawingML `a:effectLst/a:glow` owner while
remaining independent from the existing outer-shadow field.

This closes the smallest useful F-05 effect gap: color, radius, and opacity
are carried through the PPJ compiler, native wire, and source-free PPTX
authoring path. Reflection, inner shadow, soft edge, 3-D effects, and complex
imported effect graphs remain outside the slice.
