# Source-bound accent1 gray

## Why

PPJ already has an authored `gray` transform, but a direct imported
`a:gray` leaf is still opaque. A strict single-leaf source-bound owner closes
one concrete F-15 theme transform gap without claiming arbitrary theme-graph
rewrites.

## Scope

- Project a canonical imported presentation whose shared ThemePart contains an
  accent1 direct RGB leaf with exactly one direct `a:gray` child as
  `design.theme.accentTransforms.accent1.gray`.
- Issue `setThemeAccent1Gray` for that field.
- Compile `gray: false` by removing the existing direct `a:gray` child; retain
  `gray: true` as a no-op. The field is exposed only when that direct owner
  already exists, so adding a missing source leaf stays source-owned.
- Keep no-op bytes, unrelated ZIP parts, other theme roles, and unsupported
  topology source-owned; reject deletion, sibling/combined edits, invalid
  types, and capability tampering.
- Add the schema, semantic map, capability registry, generated/reference docs,
  focused native regression, and OpenSpec evidence.

## Out of scope

This does not add gray support for accent2 through accent6, create an absent
`a:gray` child, edit gray together with another transform, infer inherited
theme state, or claim PowerPoint host rendering equivalence.
