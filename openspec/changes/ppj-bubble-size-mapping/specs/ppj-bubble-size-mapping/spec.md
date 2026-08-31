## ADDED Requirements

### Requirement: Bounded authored bubble size mapping
PPJ SHALL express deliberate, finite visible bubble sizing without exposing an
arbitrary transform language.

#### Scenario: Order-of-magnitude evidence uses logarithmic radii
- **WHEN** a bubble chart declares `bubbleSizeScale: log` and a valid radius
  range
- **THEN** the compiler maps the shared positive size domain into stable
  editable circles within that range

#### Scenario: Bubble series remain comparable
- **WHEN** a chart contains multiple bubble series
- **THEN** all series use one shared minimum and maximum size domain

#### Scenario: Explicit sizing is absent
- **WHEN** a normal bubble chart does not declare the new fields
- **THEN** it continues to compile as the existing native ChartPart

#### Scenario: Sizing request is invalid
- **WHEN** the range is reversed, equal, outside the point budget, attached to a
  chart without bubbles, or paired with nonpositive data
- **THEN** validation rejects before PPTX output

#### Scenario: Authored program is recovered
- **WHEN** an explicitly sized bubble chart is built and imported again
- **THEN** OfficeKit restores the exact embedded PPJ and does not infer semantics
  from the generated group alone
