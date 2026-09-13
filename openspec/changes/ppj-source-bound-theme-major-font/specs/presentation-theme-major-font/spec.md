## Purpose

Expose one complete source-bound PPJ field for the existing presentation theme's Latin major typeface while keeping every other theme font slot source-owned.

## ADDED Requirements

### Requirement: Canonical major font projection

A source-derived PPJ projection SHALL expose `design.theme.fontScheme.major` from the existing `a:majorFont/a:latin/@typeface` when every slide master resolves to one unique ThemePart with non-empty major and minor Latin typefaces. The projection SHALL retain the other font scheme values as observed metadata and issue a theme native reference whose capability targets only `fontScheme.major`.

#### Scenario: Existing major font is projected with authority

- **WHEN** an imported presentation has one shared ThemePart with non-empty major and minor Latin typefaces
- **THEN** `design.theme.fontScheme.major` equals the source major Latin typeface and the theme native reference issues `setThemeFontScheme` for `fontScheme.major`

#### Scenario: Unsupported font ownership stays source-owned

- **WHEN** a master has no ThemePart, masters resolve to multiple ThemeParts, or either required Latin typeface is missing
- **THEN** the projection does not issue a major-font capability and source-bound compilation cannot claim ownership of `fontScheme.major`

### Requirement: Source-bound major font edit

A source-bound PPJ compile SHALL accept a changed existing `design.theme.fontScheme.major` only with the freshly projected theme native reference and its `setThemeFontScheme` capability. The export SHALL replace only the owning ThemePart's `a:majorFont/a:latin/@typeface`, report that ThemePart as changed, and preserve all other theme slots, descendants, relationships, and package members.

#### Scenario: Major font edit changes one theme part

- **WHEN** a caller changes the projected major Latin typeface and retains the issued native reference
- **THEN** compilation succeeds, reports only the canonical ThemePart as changed, and a second projection returns the new major typeface

#### Scenario: Other font slots or authority are rejected

- **WHEN** a caller changes minor, East Asian, or complex-script slots, removes the major value, changes the native reference, or changes the capability fields
- **THEN** compilation fails with a source-bound validation error without rewriting the input source package

#### Scenario: No-op preserves source bytes

- **WHEN** a caller compiles the unchanged projected PPJ against its exact source package
- **THEN** the output bytes and changed-part list are identical to the validated source and empty respectively
