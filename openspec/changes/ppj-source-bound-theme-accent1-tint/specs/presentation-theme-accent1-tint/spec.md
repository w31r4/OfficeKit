## Purpose

Expose one complete, field-qualified source-bound PPJ owner for an existing accent1 tint while retaining the imported theme XML and leaving every other transform and theme field source-owned.

## ADDED Requirements

### Requirement: Canonical accent1 tint projection

A source-derived PPJ projection SHALL expose `design.theme.accentTransforms.accent1.tint` when every slide master resolves to one unique ThemePart and the accent1 color is a direct six-digit RGB leaf with exactly one direct `a:tint` child whose `val` is an integer from 0 through 100000, with no extra attributes or descendants. The projection SHALL issue `setThemeAccent1Tint` with only `accentTransforms.accent1.tint` in the theme native reference.

#### Scenario: Existing direct accent1 tint is projected with authority

- **WHEN** an imported presentation has one shared ThemePart whose accent1 color contains a direct RGB value and one direct tint value
- **THEN** `design.theme.accentTransforms.accent1.tint` equals the source tint divided by 100000 and the theme native reference issues `setThemeAccent1Tint` for `accentTransforms.accent1.tint`

#### Scenario: Unsupported transform ownership stays source-owned

- **WHEN** the source has no unique ThemePart, a missing or transformed RGB owner, an alpha child, more than one transform, an unsupported transform type, or extra attributes/descendants
- **THEN** the projection omits `accentTransforms.accent1.tint` and does not issue the accent1 tint capability

### Requirement: Source-bound accent1 tint edit

A source-bound PPJ compile SHALL accept a changed existing `design.theme.accentTransforms.accent1.tint` only with the freshly projected theme native reference and its `setThemeAccent1Tint` capability. The export SHALL replace only the owning ThemePart's direct `a:accent1Color/a:srgbClr/a:tint/@val`, report that ThemePart as changed, and preserve the base RGB, all other accents and color roles, other transform leaves, descendants, relationships, and package members.

#### Scenario: Accent1 tint edit changes one theme part

- **WHEN** a caller changes the projected tint fraction and retains the issued native reference
- **THEN** compilation succeeds, reports only the canonical ThemePart as changed, and a second projection returns the new tint fraction

#### Scenario: Other transforms or authority are rejected

- **WHEN** a caller deletes the tint, adds or changes another transform or theme field in the same compile, changes the base accent color, or tampers with the capability
- **THEN** compilation fails without rewriting the source package

#### Scenario: No-op preserves source bytes

- **WHEN** a caller compiles the unchanged projected PPJ against the exact source package
- **THEN** output bytes and changed-part list are identical to the source and empty respectively
