## Purpose

Expose a bounded, source-preserving PPJ owner for a direct imported accent1 alpha-modulation leaf while retaining unsupported theme topology as source-owned data.

## ADDED Requirements

### Requirement: Project a strict imported alpha-modulation leaf

A source-derived PPJ projection SHALL expose `design.theme.accentTransforms.accent1.alphaMod` only when every slide master resolves to one unique ThemePart and the accent1 color is a direct six-digit RGB leaf with exactly one direct `a:alphaMod` child whose `val` is an integer from 0 through 100000, with no extra attributes or descendants. The projection SHALL issue `setThemeAccent1AlphaMod` with only `accentTransforms.accent1.alphaMod` in the theme native reference.

#### Scenario: Direct alpha modulation is projected
- **WHEN** the canonical accent1 RGB leaf contains only one direct `a:alphaMod/@val` with value 60000
- **THEN** the projection exposes `design.theme.accentTransforms.accent1.alphaMod` as `0.6` and issues the field-qualified alpha-modulation capability

#### Scenario: Unsupported alpha topology remains source-owned
- **WHEN** the accent1 color is missing, inherited, alpha-bearing, has extra attributes or descendants, has another transform, or has an alpha-modulation value outside 0 through 100000
- **THEN** the projection omits `accentTransforms.accent1.alphaMod` and does not issue the alpha-modulation capability

### Requirement: Patch only the owned alpha-modulation token

A source-bound PPJ compile SHALL accept a changed existing `design.theme.accentTransforms.accent1.alphaMod` only with the freshly projected theme native reference and its `setThemeAccent1AlphaMod` capability. The export SHALL replace only the owning ThemePart's direct `a:accent1Color/a:srgbClr/a:alphaMod/@val`, report that ThemePart as changed, and preserve the base RGB, other accents and color roles, all other transform leaves, descendants, relationships, and package members.

#### Scenario: Alpha modulation edits round-trip
- **WHEN** a projected `alphaMod` value changes from `0.6` to `0.4`
- **THEN** only the owning ThemePart changes, its direct `a:alphaMod/@val` becomes `40000`, and a fresh projection returns `0.4`

#### Scenario: No-op preserves package bytes
- **WHEN** the projected PPJ is compiled without changing the alpha-modulation field
- **THEN** the compile succeeds with no changed parts and byte-identical PPTX output

### Requirement: Reject edits outside the alpha-modulation owner

A source-bound PPJ compile SHALL fail closed when the alpha-modulation field is deleted, added, changed together with another theme field or transform, outside the `0..1` fraction range, or authorized by a missing, stale, or tampered capability.

#### Scenario: Unsupported ownership is rejected
- **WHEN** a request deletes `alphaMod`, edits a sibling transform, combines it with a color or font edit, supplies an out-of-range fraction, or changes the capability fields/hash
- **THEN** compilation fails without producing an edited PPTX
