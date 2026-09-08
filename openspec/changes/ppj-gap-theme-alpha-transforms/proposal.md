# Proposal: preserve bounded theme alpha transforms

Extend authored PPJ `design.theme.accentTransforms` with `alphaMod` and
`alphaOff` for accent1 through accent6. The fields lower directly to
DrawingML `a:alphaMod` and `a:alphaOff`, distinct from the absolute alpha
carried by an authored RGBA color.

This closes two concrete F-15 color-transform fields with authored XML and a
minimal native regression. Imported theme editing and the full color/effect
cascade remain outside the slice.
