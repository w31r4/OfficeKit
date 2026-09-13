# Source-bound accent1 hueOff

## Why

PPJ already has an authored `hueOff` transform, but a direct imported
`a:hueOff` leaf is still opaque. A strict single-leaf source-bound owner closes
one concrete F-15 theme transform gap without pretending to edit an arbitrary
theme graph.

## Scope

- Project a canonical imported presentation whose shared ThemePart contains an
  accent1 direct RGB leaf with exactly one direct `a:hueOff/@val` child as
  `design.theme.accentTransforms.accent1.hueOff`, expressed in degrees.
- Issue `setThemeAccent1HueOff` for that field and compile one authorized
  change by rewriting only the existing `a:hueOff/@val` token.
- Convert PPJ degrees to DrawingML's 1/60000-degree integer with
  away-from-zero rounding, and convert the native integer back to degrees on a
  fresh projection.
- Keep no-op bytes, unrelated ZIP parts, other theme roles, and unsupported
  topology source-owned; reject deletion, sibling/combined edits, bad ranges,
  and capability tampering.
- Add the schema, semantic map, capability registry, generated/reference docs,
  focused native regression, and OpenSpec evidence.

## Out of scope

This does not add hueOff support for accent2 through accent6, edit hueMod or
other transform siblings in the same source-bound owner, infer inherited theme
state, or claim PowerPoint host rendering equivalence.
