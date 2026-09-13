## Purpose

Expose one complete source-bound PPJ field for the existing presentation
theme's Latin minor typeface while keeping other font slots source-owned.

## ADDED Requirements

### Requirement: Canonical minor font projection

A source-derived PPJ projection SHALL expose
`design.theme.fontScheme.minor` from the existing
`a:minorFont/a:latin/@typeface` when every slide master resolves to one unique
ThemePart with non-empty major and minor Latin typefaces. The projection SHALL
issue a theme native reference whose capability targets `fontScheme.minor`.

#### Scenario: Existing minor font is projected with authority

- **WHEN** an imported presentation has one shared ThemePart with non-empty major and minor Latin typefaces
- **THEN** `design.theme.fontScheme.minor` equals the source minor Latin typeface and the theme native reference issues `setThemeMinorFont` for `fontScheme.minor`

#### Scenario: Unsupported font ownership stays source-owned

- **WHEN** a master has no ThemePart, masters resolve to multiple ThemeParts, or either required Latin typeface is missing
- **THEN** the projection does not issue a minor-font capability and source-bound compilation cannot claim ownership of `fontScheme.minor`

### Requirement: Source-bound minor font edit

A source-bound PPJ compile SHALL accept a changed existing
`design.theme.fontScheme.minor` only with the freshly projected theme native
reference and its `setThemeMinorFont` capability. The export SHALL replace
only the owning ThemePart's `a:minorFont/a:latin/@typeface`, report that
ThemePart as changed, and preserve all other theme slots, descendants,
relationships, and package members.

#### Scenario: Minor font edit changes one theme part

- **WHEN** a caller changes the projected minor Latin typeface and retains the issued native reference
- **THEN** compilation succeeds, reports only the canonical ThemePart as changed, and a second projection returns the new minor typeface

#### Scenario: Other slots or authority are rejected

- **WHEN** a caller changes major together with minor, changes East Asian or complex-script slots, removes the minor value, changes the native reference, or changes capability fields
- **THEN** compilation fails with a source-bound validation error without rewriting the input source package

#### Scenario: No-op preserves source bytes

- **WHEN** a caller compiles the unchanged projected PPJ against its exact source package
- **THEN** the output bytes and changed-part list are identical to the validated source and empty respectively
