# PPJ authored theme color-role alpha

## ADDED Requirements

### Requirement: accept alpha-bearing authored theme roles

The PPJ schema MUST accept the existing six-digit RGB spelling and a
`#RRGGBBAA` spelling for every authored accent and non-accent theme color
role. Values outside those forms MUST be rejected.

#### Scenario: six-digit compatibility

- **WHEN** an authored theme role uses `#102030`
- **THEN** validation succeeds and the native role has no `a:alpha` child

#### Scenario: alpha is lowered

- **WHEN** an authored theme role uses `#10203080`
- **THEN** validation succeeds
- **AND** native `theme1.xml` contains `a:srgbClr/@val="102030"`
- **AND** the same color contains one `a:alpha` corresponding to `0x80`

### Requirement: preserve authored package evidence

The authored compiler MUST embed the exact input PPJ and MUST keep the output
valid under the Office 2021 Open XML validator.

#### Scenario: minimal round trip

- **WHEN** the bounded program is compiled
- **THEN** the embedded PPJ bytes equal the input bytes
- **AND** Open XML validation reports no errors

### Requirement: keep imported theme ownership closed

The source-bound compiler MUST NOT treat an authored theme-role alpha as a
permission to rebuild an imported theme graph.

#### Scenario: source-bound rejection

- **WHEN** a source-bound PPJ adds `design.theme.colorRoles` or
  `design.theme.accentColors`
- **THEN** compilation fails closed with the existing source-bound theme-owner
  diagnostic
