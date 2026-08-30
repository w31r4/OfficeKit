# PPJ chart text style specification

## ADDED Requirements

### Requirement: Bounded chart typography

PPJ SHALL support optional font size, Latin font family, East Asian font family,
bold, italic, color, and color opacity for chart titles and axis tick labels.

#### Scenario: Authored chart

- **WHEN** a valid PPJ chart declares the bounded text style
- **THEN** the compiler writes editable DrawingML text properties and a
  re-import projects the same executable values

### Requirement: Source-bound style edit

A canonical imported chart SHALL issue `setChartTextStyle`; changing only the
declared title or axis text-style fields SHALL patch the chart part. A stale,
irregular, or unissued edit SHALL fail closed.

#### Scenario: Third-party round trip

- **WHEN** an imported canonical chart style is changed with its fresh
  capability
- **THEN** only the required chart part changes and a second projection reports
  the requested values

### Requirement: Agent discoverability

The schema, capability registry, generated PPJ manual, and chart guidance SHALL
describe the same fields and native boundary.

#### Scenario: Maintainer synchronization

- **WHEN** the Presentation Skill maintainer checks the repository
- **THEN** the generated manual matches the chart text-style schema and registry
