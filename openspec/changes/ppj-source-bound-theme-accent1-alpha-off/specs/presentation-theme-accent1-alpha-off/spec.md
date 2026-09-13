## Purpose

Expose a bounded, source-preserving PPJ owner for a direct imported accent1 alpha-offset leaf while retaining unsupported theme topology as source-owned data.

## ADDED Requirements

### Requirement: Project a strict imported alpha-offset leaf

A source-derived PPJ projection SHALL expose `design.theme.accentTransforms.accent1.alphaOff` only when every slide master resolves to one unique ThemePart and the accent1 color is a direct six-digit RGB leaf with exactly one direct `a:alphaOff` child whose `val` is an integer from -100000 through 100000, with no extra attributes or descendants. The projection SHALL issue `setThemeAccent1AlphaOff` with only `accentTransforms.accent1.alphaOff` in the theme native reference.

#### Scenario: Direct alpha offset is projected
- **WHEN** the canonical accent1 RGB leaf contains only one direct `a:alphaOff/@val` with value -25000
- **THEN** the projection exposes `design.theme.accentTransforms.accent1.alphaOff` as `-0.25` and issues the field-qualified alpha-offset capability

#### Scenario: Unsupported alpha-offset topology remains source-owned
- **WHEN** the accent1 color is missing, inherited, alpha-bearing, has extra attributes or descendants, has another transform, or has an alpha-offset value outside -100000 through 100000
- **THEN** the projection omits `accentTransforms.accent1.alphaOff` and does not issue the alpha-offset capability

### Requirement: Patch only the owned alpha-offset token

A source-bound PPJ compile SHALL accept a changed existing `design.theme.accentTransforms.accent1.alphaOff` only with the freshly projected theme native reference and its `setThemeAccent1AlphaOff` capability. The export SHALL replace only the owning ThemePart's direct `a:accent1Color/a:srgbClr/a:alphaOff/@val`, report that ThemePart as changed, and preserve the base RGB, other accents and color roles, all other transform leaves, descendants, relationships, and package members.

#### Scenario: Alpha offset edits round-trip
- **WHEN** a projected `alphaOff` value changes from `-0.25` to `0.4`
- **THEN** only the owning ThemePart changes, its direct `a:alphaOff/@val` becomes `40000`, and a fresh projection returns `0.4`

#### Scenario: No-op preserves package bytes
- **WHEN** the projected PPJ is compiled without changing the alpha-offset field
- **THEN** the compile succeeds with no changed parts and byte-identical PPTX output

### Requirement: Reject edits outside the alpha-offset owner

A source-bound PPJ compile SHALL fail closed when the alpha-offset field is deleted, added, changed together with another theme field or transform, outside the `-1..1` fraction range, or authorized by a missing, stale, or tampered capability.

#### Scenario: Unsupported ownership is rejected
- **WHEN** a request deletes `alphaOff`, edits a sibling transform, combines it with a color or font edit, supplies an out-of-range fraction, or changes the capability fields/hash
- **THEN** compilation fails without producing an edited PPTX
