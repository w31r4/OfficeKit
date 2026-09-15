## ADDED Requirements

### Requirement: Canonical accent1 RGBA projection

A source-derived PPJ projection SHALL preserve `design.theme.accentColors.accent1` as `#RRGGBBAA` when the canonical shared ThemePart has six strict direct accent colors and accent1 is a direct `a:srgbClr/@val` with exactly one canonical `a:alpha/@val`; the other accents remain six-digit direct RGB. The existing theme native reference SHALL issue `setThemeAccent1Color` for `accentColors.accent1`.

#### Scenario: Existing canonical alpha is projected with authority

- **WHEN** an imported presentation has one shared ThemePart, six direct accent colors, and accent1 has one canonical alpha child
- **THEN** `design.theme.accentColors.accent1` contains the source RGB plus the two-digit alpha suffix and the native reference retains the accent1-color capability

#### Scenario: Unsupported alpha topology stays source-owned

- **WHEN** accent1 has no alpha, multiple alpha children, a non-canonical alpha value, another transform, a scheme color, or a missing/ambiguous ThemePart
- **THEN** the source-bound projection does not claim an RGBA accent1 owner

### Requirement: Source-bound accent1 RGBA edit

A source-bound PPJ compile SHALL accept a changed existing `design.theme.accentColors.accent1` RGBA value only with the projected hash-bound native reference and `setThemeAccent1Color` capability. It SHALL update only `a:accent1Color/a:srgbClr/@val` and the existing `a:alpha/@val`, report the canonical ThemePart as changed, and preserve all other package members.

#### Scenario: RGBA edit changes one theme part

- **WHEN** a caller changes the projected accent1 RGB and alpha bytes while retaining the issued authority
- **THEN** compilation succeeds, reports only the canonical ThemePart, and a second projection returns the new `#RRGGBBAA` value

#### Scenario: Alpha topology or authority changes are rejected

- **WHEN** a caller changes alpha presence, changes another accent/theme field, removes accent1, supplies an out-of-range value, or tampers with capability fields
- **THEN** compilation fails without rewriting the source package

#### Scenario: No-op preserves source bytes

- **WHEN** the unchanged RGBA projection is compiled against its exact source package
- **THEN** output bytes and changed parts are identical to the source and empty respectively
