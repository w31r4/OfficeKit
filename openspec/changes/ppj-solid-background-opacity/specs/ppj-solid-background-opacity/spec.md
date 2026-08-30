# PPJ solid-background opacity

## ADDED Requirements

### Requirement: Solid slide backgrounds retain opacity

OfficeKit SHALL compile, import, project and source-bound edit the bounded
opacity of a PPJ solid slide background.

#### Scenario: Build and project a translucent background

- **GIVEN** a valid PPJ page with a translucent solid background
- **WHEN** OfficeKit builds and projects the deck
- **THEN** the native slide background contains the equivalent direct alpha
- **AND** projected PPJ retains the equivalent opacity

#### Scenario: Preserve unsupported background color transforms

- **GIVEN** an imported slide background with transforms outside the bounded
  direct-alpha profile
- **WHEN** OfficeKit imports or edits an unrelated object
- **THEN** the background remains opaque-preserved
- **AND** OfficeKit does not flatten or normalize it
