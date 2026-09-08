## ADDED Requirements

### Requirement: Authored PPJ declares explicit theme accent roles

An authored PPJ program MAY declare `design.theme.accentColors`. When present,
it MUST contain six non-transparent RGB strings named `accent1` through
`accent6`, and the compiler MUST lower them to the corresponding native
DrawingML accent color roles.

#### Scenario: Explicit roles override the legacy color array

- **WHEN** `design.theme.accentColors` declares six values while
  `design.theme.colors` contains different values
- **THEN** the generated theme contains the explicit values in
  `a:accent1Color` through `a:accent6Color`

#### Scenario: Embedded authored PPJ recovers the declaration

- **WHEN** an authored PPTX containing an explicit theme accent map is
  projected through the OfficeKit embedded PPJ path
- **THEN** the recovered PPJ contains the same six named values

### Requirement: Legacy theme accent behavior remains compatible

When `design.theme.accentColors` is absent, authored compilation MUST retain
the existing behavior: the first six `design.theme.colors` values supply the
native accent roles.

#### Scenario: Existing PPJ omits the explicit map

- **WHEN** a valid program contains only its existing `design.theme.colors`
  catalog
- **THEN** compilation continues to produce the same bounded accent palette

### Requirement: Source-bound theme accent maps fail closed

Source-bound PPJ MUST NOT use `design.theme.accentColors` to introduce or edit
a native theme graph. A source-bound request carrying the field MUST fail
before writing package parts with diagnostic code
`ppj.sourceBound.themeAccentColors`.

#### Scenario: Source-bound request attempts a theme change

- **WHEN** an imported source-bound PPJ adds or changes
  `design.theme.accentColors`
- **THEN** the request is rejected and the source package is not rewritten
