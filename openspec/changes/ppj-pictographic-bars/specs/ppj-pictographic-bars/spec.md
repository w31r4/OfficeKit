## ADDED Requirements

### Requirement: Authored PPJ pictographic bars
PPJ SHALL express bounded pictographic bars and columns using existing named
icon or preset-geometry symbols and compile them into editable DrawingML.

#### Scenario: Valid pictographic bar is authored
- **WHEN** one bar or column series declares a valid symbol and exact unit-sized
  non-negative values
- **THEN** the compiler emits the corresponding stable editable symbols plus
  category and value labels

#### Scenario: Symbol catalog is reused
- **WHEN** a symbol names an OfficeKit icon or preset geometry
- **THEN** the compiler resolves the existing pinned catalog without network
  access or raw path input

#### Scenario: Authored pictographic bar is reimported
- **WHEN** its OfficeKit-authored PPTX retains the embedded program
- **THEN** PPJ import restores the exact chart, stable IDs, values, unit and
  symbol declaration

### Requirement: Pictographic expansion fails closed
PPJ SHALL reject data or options outside the bounded exact profile.

#### Scenario: Value requires a fractional symbol
- **WHEN** a series value is not an exact multiple of its declared unit
- **THEN** validation rejects instead of clipping or rounding a symbol

#### Scenario: Expansion exceeds its budget
- **WHEN** one category would emit more than 32 symbols or the chart more than
  192 symbols
- **THEN** validation rejects before native output is written

#### Scenario: Imported group resembles a pictographic chart
- **WHEN** an imported PPTX contains repeated icon-like shapes without an
  embedded PPJ program
- **THEN** OfficeKit projects the editable group without inferring chart data or
  issuing a pictographic edit capability
