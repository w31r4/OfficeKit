# Proposal: preserve bounded theme saturation transforms

Extend authored PPJ `design.theme.accentTransforms` with `satMod` and
`satOff` for accent1 through accent6. The fields lower directly to DrawingML
`a:satMod` and `a:satOff` while preserving the base accent role and the other
bounded transform fields.

This closes one concrete F-15 color-transform family with authored XML and a
minimal native regression. Imported theme editing and the full color/effect
cascade remain outside the slice.
