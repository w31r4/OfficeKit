# PPJ chart-series color opacity

## ADDED Requirements

### Requirement: Compact series colors retain alpha

OfficeKit SHALL compile alpha-bearing PPJ chart-series `color` values through
the native bounded solid-fill profile.

#### Scenario: Build and project a translucent series color

- **GIVEN** a valid PPJ category series with `color: "#0A84FF80"`
- **WHEN** OfficeKit builds and projects the deck
- **THEN** the native series paint retains the alpha value
- **AND** projected PPJ contains an equivalent structured solid fill

