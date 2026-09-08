# Proposal: preserve bounded theme channel transforms

Extend authored PPJ `design.theme.accentTransforms` with `redMod`/`redOff`,
`greenMod`/`greenOff`, and `blueMod`/`blueOff` for accent1 through accent6.
The fields lower directly to DrawingML `a:redMod`/`a:redOff`,
`a:greenMod`/`a:greenOff`, and `a:blueMod`/`a:blueOff` while preserving the
base accent role and the other bounded transform fields.

This closes the RGB-channel color-transform family with authored XML and a
minimal native regression. Imported theme editing and the full color/effect
cascade remain outside the slice.
