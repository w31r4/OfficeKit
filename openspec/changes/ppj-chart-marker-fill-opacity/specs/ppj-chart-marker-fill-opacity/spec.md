# PPJ chart-marker fill opacity

## ADDED Requirements

### Requirement: Chart-marker fill retains alpha

OfficeKit SHALL compile, import and project the bounded opacity of a direct
solid chart-marker fill.

#### Scenario: Build and project a translucent marker

- **GIVEN** a valid PPJ line, scatter or radar series with an alpha-bearing
  marker fill
- **WHEN** OfficeKit builds and projects the deck
- **THEN** the marker contains the equivalent direct native alpha
- **AND** projected PPJ retains the equivalent alpha-bearing color

#### Scenario: Preserve unsupported marker paint

- **GIVEN** an imported marker with a paint graph outside the bounded profile
- **WHEN** OfficeKit imports the presentation
- **THEN** the containing chart remains source-preserved and read-only
