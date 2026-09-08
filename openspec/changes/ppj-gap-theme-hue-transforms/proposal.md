# Proposal: preserve bounded theme hue transforms

Extend authored PPJ `design.theme.accentTransforms` with `hueMod` and
`hueOff` for accent1 through accent6. The modulation lowers to DrawingML
`a:hueMod`; the offset lowers to `a:hueOff` using DrawingML's 1/60000-degree
angle unit.

This closes one bounded hue-transform family with authored XML and a minimal
native regression. Imported theme editing and the full color/effect cascade
remain outside the slice.
