## Purpose

Expose one complete source-bound PPJ field for the existing presentation
theme's major East Asian typeface while keeping other font slots source-owned.

## ADDED Requirements

### Requirement: Canonical major East Asian font projection

A source-derived PPJ projection SHALL expose
`design.theme.fontScheme.majorEastAsia` from the existing
`a:majorFont/a:ea/@typeface` when every slide master resolves to one unique
ThemePart with non-empty major/minor Latin typefaces and a non-empty major East
Asian typeface. The projection SHALL issue a theme native reference whose
capability targets `fontScheme.majorEastAsia`.

#### Scenario: Existing major East Asian font is projected with authority

- **WHEN** an imported presentation has one shared ThemePart with the required Latin faces and a major East Asian typeface
- **THEN** `design.theme.fontScheme.majorEastAsia` equals the source value and the theme native reference issues `setThemeMajorFontEastAsia` for `fontScheme.majorEastAsia`

#### Scenario: Unsupported script ownership stays source-owned

- **WHEN** a master has no ThemePart, masters resolve to multiple ThemeParts, a required Latin face is missing, or the major East Asian slot is missing
- **THEN** the projection does not issue a major East Asian capability and source-bound compilation cannot claim ownership of `fontScheme.majorEastAsia`

### Requirement: Source-bound major East Asian font edit

A source-bound PPJ compile SHALL accept a changed existing
`design.theme.fontScheme.majorEastAsia` only with the freshly projected theme
native reference and its `setThemeMajorFontEastAsia` capability. The export
SHALL replace only the owning ThemePart's `a:majorFont/a:ea/@typeface`, report
that ThemePart as changed, and preserve all other font slots, descendants,
relationships, and package members.

#### Scenario: Major East Asian font edit changes one theme part

- **WHEN** a caller changes the projected major East Asian typeface and retains the issued native reference
- **THEN** compilation succeeds, reports only the canonical ThemePart as changed, and a second projection returns the new value

#### Scenario: Other slots or authority are rejected

- **WHEN** a caller changes a Latin slot together with major East Asian, changes another East Asian or complex-script slot, removes the East Asian value, changes the native reference, or changes capability fields
- **THEN** compilation fails with a source-bound validation error without rewriting the input source package

#### Scenario: No-op preserves source bytes

- **WHEN** a caller compiles the unchanged projected PPJ against its exact source package
- **THEN** the output bytes and changed-part list are identical to the validated source and empty respectively
