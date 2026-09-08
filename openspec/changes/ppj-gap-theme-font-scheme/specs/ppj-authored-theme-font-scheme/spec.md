## ADDED Requirements

### Requirement: Authored PPJ declares explicit theme font roles

An authored PPJ program MAY declare `design.theme.fontScheme`. When present,
it MUST contain non-empty string fields `major` and `minor`, and the compiler
MUST lower them to the corresponding native `a:majorFont` and `a:minorFont`
families. The bounded writer uses each family for its Latin, East Asian, and
complex-script typeface slots.

#### Scenario: Explicit roles override the legacy font array

- **WHEN** `design.theme.fontScheme` declares `major: "Aptos Display"` and
  `minor: "Aptos"`, while `design.fonts` contains different families
- **THEN** the generated theme contains the explicit major/minor families in
  all three bounded script slots

#### Scenario: Embedded authored PPJ recovers the declaration

- **WHEN** an authored PPTX containing an explicit theme font scheme is
  projected through the OfficeKit embedded PPJ path
- **THEN** the recovered PPJ contains the same `major` and `minor` values

### Requirement: Legacy theme font behavior remains compatible

When `design.theme.fontScheme` is absent, authored compilation MUST retain the
existing behavior: the first `design.fonts` family supplies the major role and
the second family supplies the minor role, falling back to the first family
when no second family exists.

#### Scenario: Existing PPJ omits the explicit scheme

- **WHEN** a valid program contains only its existing `design.fonts` catalog
- **THEN** compilation continues to produce the same bounded major/minor theme
  font behavior

### Requirement: Source-bound theme font schemes fail closed

Source-bound PPJ MUST NOT use `design.theme.fontScheme` to introduce or edit a
native theme graph. A source-bound request carrying the field MUST fail before
writing package parts with diagnostic code `ppj.sourceBound.themeFontScheme`.

#### Scenario: Source-bound request attempts a theme change

- **WHEN** an imported source-bound PPJ adds or changes
  `design.theme.fontScheme`
- **THEN** the request is rejected and the source package is not rewritten
