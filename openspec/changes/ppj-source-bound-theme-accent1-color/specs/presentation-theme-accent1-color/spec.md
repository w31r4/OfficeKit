## Purpose

Expose one complete source-bound PPJ field for an existing presentation theme accent color while keeping the rest of the theme color graph source-owned.

## ADDED Requirements

### Requirement: Canonical accent color projection

A source-derived PPJ projection SHALL expose `design.theme.accentColors` from six existing `a:accentNColor/a:srgbClr/@val` values when every slide master resolves to one unique ThemePart and all six accent nodes are direct six-digit RGB colors without descendants. The projection SHALL issue a theme native reference whose capability targets only `accentColors.accent1`.

#### Scenario: Existing direct accent palette is projected with authority

- **WHEN** an imported presentation has one shared ThemePart with six strict direct RGB accent colors
- **THEN** `design.theme.accentColors.accent1` equals the source accent1 color and the theme native reference issues `setThemeAccent1Color` for `accentColors.accent1`

#### Scenario: Unsupported accent ownership stays source-owned

- **WHEN** a master has no ThemePart, masters resolve to multiple ThemeParts, an accent node is missing, or any accent uses a non-RGB child, alpha, or transform
- **THEN** the projection does not issue an accent1-color capability and source-bound compilation cannot claim ownership of `accentColors.accent1`

### Requirement: Source-bound accent1 edit

A source-bound PPJ compile SHALL accept a changed existing `design.theme.accentColors.accent1` only with the freshly projected theme native reference and its `setThemeAccent1Color` capability. The export SHALL replace only the owning ThemePart's `a:accent1Color/a:srgbClr/@val`, report that ThemePart as changed, and preserve all other accent slots, color roles, descendants, relationships, and package members.

#### Scenario: Accent1 edit changes one theme part

- **WHEN** a caller changes the projected accent1 value and retains the issued native reference
- **THEN** compilation succeeds, reports only the canonical ThemePart as changed, and a second projection returns the new accent1 value

#### Scenario: Other accents or authority are rejected

- **WHEN** a caller changes accent2 through accent6, removes accent1, changes the native reference, changes another theme field in the same compile, or changes the capability fields
- **THEN** compilation fails with a source-bound validation error without rewriting the input source package

#### Scenario: No-op preserves source bytes

- **WHEN** a caller compiles the unchanged projected PPJ against its exact source package
- **THEN** the output bytes and changed-part list are identical to the validated source and empty respectively
