## Purpose

Expose one complete, source-bound PPJ field for the existing presentation theme name while preserving the rest of the imported theme graph as opaque source-owned content.

## ADDED Requirements

### Requirement: Canonical theme name projection

A source-derived PPJ projection SHALL expose `design.theme.name` from the existing presentation theme root when every slide master resolves to one unique ThemePart with one non-empty `a:theme/@name` value. The theme object SHALL include a native reference whose capability targets only the theme name field.

#### Scenario: Existing shared theme name is projected

- **WHEN** an imported presentation has one shared ThemePart and a non-empty theme name
- **THEN** the PPJ theme name equals the source `a:theme/@name` value and its native reference issues `setThemeName` for `name`

#### Scenario: Ambiguous theme ownership stays source-owned

- **WHEN** masters resolve to multiple ThemeParts, a master has no ThemePart, or the theme name is missing
- **THEN** the projection keeps the bounded fallback theme metadata and issues no theme-name capability

### Requirement: Source-bound theme name edit

A source-bound PPJ compile SHALL accept a changed existing `design.theme.name` only with the freshly projected theme native reference and its `setThemeName` capability. The export SHALL replace only the existing `a:theme/@name` value in the unique ThemePart, report that ThemePart as the changed part, and preserve all other theme XML and package members.

#### Scenario: Name edit patches only the theme part

- **WHEN** a caller changes the projected theme name and retains the issued native reference
- **THEN** compilation succeeds, reports only the canonical ThemePart as changed, and a second projection returns the new name

#### Scenario: Name deletion or capability tampering fails closed

- **WHEN** a caller removes the projected name, changes another theme property, changes the native reference, or removes `setThemeName`
- **THEN** compilation fails with a source-bound validation error and does not modify the input source bytes

#### Scenario: No-op preserves source bytes

- **WHEN** a caller compiles the unchanged projected PPJ against its exact source package
- **THEN** the output bytes and changed-part list are identical to the validated source and empty respectively
