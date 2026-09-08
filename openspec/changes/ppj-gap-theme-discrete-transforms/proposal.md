# Proposal: preserve bounded discrete theme color transforms

Extend authored PPJ `design.theme.accentTransforms` with the direct
DrawingML color operations `gray`, `comp`, `inv`, `gamma`, and `invGamma` for
accent1 through accent6. Each operation is an explicit boolean presence and
lowers to its corresponding empty DrawingML transform child.

This closes the no-argument color-transform operations with a minimal native
regression. Imported theme editing and the full color/effect cascade remain
outside the slice.
