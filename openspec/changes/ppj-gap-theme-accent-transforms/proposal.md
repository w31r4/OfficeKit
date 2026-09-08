# Proposal: preserve bounded theme accent transforms

Add authored PPJ `design.theme.accentTransforms` for the native DrawingML
`a:tint` and `a:shade` transforms on accent1 through accent6. The transform
is applied by the host to the existing accent role and remains distinct from
the base `accentColors` value.

This closes one concrete F-15 color-transform field with authored XML and
embedded recovery evidence. Full theme transform/effect schemes and
source-bound theme editing remain outside the slice.
