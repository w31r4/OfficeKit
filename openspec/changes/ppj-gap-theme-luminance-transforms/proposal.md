# Proposal: preserve bounded theme luminance transforms

Extend authored PPJ `design.theme.accentTransforms` with `lumMod` and
`lumOff` for accent1 through accent6. The fields lower directly to the
corresponding DrawingML color transforms while keeping the base accent role
and existing tint/shade fields separate.

This closes two concrete F-15 color-transform fields with authored XML and a
minimal native regression. Imported theme editing and the full color/effect
cascade remain outside the slice.
