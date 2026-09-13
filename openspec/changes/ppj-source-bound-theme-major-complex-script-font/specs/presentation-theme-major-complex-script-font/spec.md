## Purpose

Expose one complete source-bound PPJ field for the existing presentation
theme's major complex-script typeface while keeping other font slots
source-owned.

## ADDED Requirements

### Requirement: Canonical major complex-script font projection

A source-derived PPJ projection SHALL expose
`design.theme.fontScheme.majorComplexScript` from the existing
`a:majorFont/a:cs/@typeface` when every slide master resolves to one unique
ThemePart with non-empty major/minor Latin typefaces and a non-empty major
complex-script typeface. The projection SHALL issue a theme native reference
whose capability targets `fontScheme.majorComplexScript`.

#### Scenario: Existing major complex-script font is projected with authority

- **WHEN** an imported presentation has one shared ThemePart with the required Latin faces and a major complex-script typeface
- **THEN** `design.theme.fontScheme.majorComplexScript` equals the source value and the theme native reference issues `setThemeMajorFontComplexScript` for `fontScheme.majorComplexScript`

#### Scenario: Unsupported script ownership stays source-owned

- **WHEN** a master has no ThemePart, masters resolve to multiple ThemeParts, a required Latin face is missing, or the major complex-script slot is missing
- **THEN** the projection does not issue a major complex-script capability and source-bound compilation cannot claim ownership of `fontScheme.majorComplexScript`

### Requirement: Source-bound major complex-script font edit

A source-bound PPJ compile SHALL accept a changed existing
`design.theme.fontScheme.majorComplexScript` only with the freshly projected
theme native reference and its `setThemeMajorFontComplexScript` capability. The
export SHALL replace only the owning ThemePart's
`a:majorFont/a:cs/@typeface`, report that ThemePart as changed, and preserve all
other font slots, descendants, relationships, and package members.

#### Scenario: Major complex-script font edit changes one theme part

- **WHEN** a caller changes the projected major complex-script typeface and retains the issued native reference
- **THEN** compilation succeeds, reports only the canonical ThemePart as changed, and a second projection returns the new value

#### Scenario: Other slots or authority are rejected

- **WHEN** a caller changes a Latin, East Asian, or minor complex-script slot together with major complex-script, removes the complex-script value, changes the native reference, or changes capability fields
- **THEN** compilation fails with a source-bound validation error without rewriting the input source package

#### Scenario: No-op preserves source bytes

- **WHEN** a caller compiles the unchanged projected PPJ against its exact source package
- **THEN** the output bytes and changed-part list are identical to the validated source and empty respectively
