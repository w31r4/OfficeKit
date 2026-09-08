## ADDED Requirements

### Requirement: Authored PPJ declares explicit non-accent theme color roles

An authored PPJ program MAY declare `design.theme.colorRoles`. When present,
it MUST contain six opaque RGB strings named `dark1`, `light1`, `dark2`,
`light2`, `hyperlink`, and `followedHyperlink`, and the compiler MUST lower
them to the corresponding native DrawingML color-scheme roles.

#### Scenario: Explicit roles lower to native color-scheme elements

- **WHEN** `design.theme.colorRoles` declares six values
- **THEN** the generated theme contains those values in the corresponding
  `a:dk1`/`a:lt1`/`a:dk2`/`a:lt2`/`a:hlink`/`a:folHlink` color roles

#### Scenario: Embedded authored PPJ recovers the declaration

- **WHEN** an authored PPTX containing explicit non-accent theme color roles is
  projected through the OfficeKit embedded PPJ path
- **THEN** the recovered PPJ contains the same six named values

### Requirement: Existing authored theme defaults remain compatible

When `design.theme.colorRoles` is absent, authored compilation MUST retain the
existing clean-room defaults for the six roles. The existing `colors` catalog
and explicit accent/font owners remain independent compatibility paths.

#### Scenario: Existing PPJ omits the role map

- **WHEN** a valid program contains its existing theme declaration but no
  `colorRoles` object
- **THEN** compilation continues to produce the previous bounded theme

### Requirement: Source-bound theme color role maps fail closed

Source-bound PPJ MUST NOT use `design.theme.colorRoles` to introduce or edit a
native theme graph. A source-bound request carrying the field MUST fail before
writing package parts with diagnostic code
`ppj.sourceBound.themeColorRoles`.

#### Scenario: Source-bound request attempts a theme change

- **WHEN** an imported source-bound PPJ adds or changes
  `design.theme.colorRoles`
- **THEN** the request is rejected and the source package is not rewritten
