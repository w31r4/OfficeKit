## Purpose

Expose one complete source-bound PPJ field for the existing presentation
theme's minor East Asian typeface while keeping other font slots source-owned.

## ADDED Requirements

### Requirement: Canonical minor East Asian font projection

A source-derived PPJ projection SHALL expose
`design.theme.fontScheme.minorEastAsia` from the existing
`a:minorFont/a:ea/@typeface` when every slide master resolves to one unique
ThemePart with non-empty major/minor Latin typefaces and a non-empty minor East
Asian typeface. The projection SHALL issue a theme native reference whose
capability targets `fontScheme.minorEastAsia`.

#### Scenario: Existing minor East Asian font is projected with authority

- **WHEN** an imported presentation has one shared ThemePart with the required Latin faces and a minor East Asian typeface
- **THEN** `design.theme.fontScheme.minorEastAsia` equals the source value and the theme native reference issues `setThemeMinorFontEastAsia` for `fontScheme.minorEastAsia`

#### Scenario: Unsupported script ownership stays source-owned

- **WHEN** a master has no ThemePart, masters resolve to multiple ThemeParts, a required Latin face is missing, or the minor East Asian slot is missing
- **THEN** the projection does not issue a minor East Asian capability and source-bound compilation cannot claim ownership of `fontScheme.minorEastAsia`

### Requirement: Source-bound minor East Asian font edit

A source-bound PPJ compile SHALL accept a changed existing
`design.theme.fontScheme.minorEastAsia` only with the freshly projected theme
native reference and its `setThemeMinorFontEastAsia` capability. The export
SHALL replace only the owning ThemePart's `a:minorFont/a:ea/@typeface`, report
that ThemePart as changed, and preserve all other font slots, descendants,
relationships, and package members.

#### Scenario: Minor East Asian font edit changes one theme part

- **WHEN** a caller changes the projected minor East Asian typeface and retains the issued native reference
- **THEN** compilation succeeds, reports only the canonical ThemePart as changed, and a second projection returns the new value

#### Scenario: Other slots or authority are rejected

- **WHEN** a caller changes a Latin or major East Asian slot together with minor East Asian, changes a complex-script slot, removes the East Asian value, changes the native reference, or changes capability fields
- **THEN** compilation fails with a source-bound validation error without rewriting the input source package

#### Scenario: No-op preserves source bytes

- **WHEN** a caller compiles the unchanged projected PPJ against its exact source package
- **THEN** the output bytes and changed-part list are identical to the validated source and empty respectively
