# PPJ chart text-style completion specification

## ADDED Requirements

### Requirement: Complete chart typography hierarchy

PPJ SHALL support the bounded chart text style for chart legends, plot-level
data labels, and axis titles in addition to chart titles and axis tick labels.

#### Scenario: Authored data page

- **WHEN** a valid PPJ chart declares the new style locations
- **THEN** the compiler writes editable native text properties and re-import
  projects the same values

### Requirement: Source-bound style continuation

A canonical imported chart SHALL expose the new locations through
`setChartTextStyle`; changing only issued style values SHALL patch the chart
part, while irregular or stale text graphs SHALL fail closed.

#### Scenario: Imported chart edit

- **WHEN** an Agent changes legend, data-label, or axis-title typography using a
  fresh PPJ projection
- **THEN** a second projection reports the requested values without rebuilding
  the surrounding slide

### Requirement: Discoverable language surface

The PPJ schema, capability registry, generated language manual, and focused
chart guidance SHALL describe the same style locations and native boundary.

#### Scenario: Maintainer check

- **WHEN** the Presentation Skill maintainer runs
- **THEN** generated documentation and capability metadata remain synchronized
