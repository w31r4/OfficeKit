## Purpose

Expose the direct ChartML chart-space rounded-corners setting as a bounded PPJ field that can survive authored and source-bound round trips.

## ADDED Requirements

### Requirement: Chart rounded corners are represented and edited
The presentation and spreadsheet chart models SHALL accept an optional boolean `roundedCorners`; authored output SHALL emit direct `c:roundedCorners/@val` when present, and projection SHALL preserve its presence and value.

#### Scenario: Authored chart round trip
- **WHEN** a bounded chart contains `roundedCorners: true`
- **THEN** the exported ChartPart contains `c:roundedCorners val="1"` and re-projection returns `true`.

#### Scenario: Source-bound toggle and removal
- **WHEN** a recognized source chart's PPJ field is changed or removed
- **THEN** only its ChartPart is changed, the XML boolean is replaced or removed, and a second projection reflects the new presence and value.

#### Scenario: Unsupported topology fails closed
- **WHEN** the chart has an external workbook, duplicate rounded-corners nodes, or an unsupported chart topology
- **THEN** no rounded-corners edit capability is issued and the source remains preserved.
