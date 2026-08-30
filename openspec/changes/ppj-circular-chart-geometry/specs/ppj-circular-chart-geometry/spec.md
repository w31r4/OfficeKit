## ADDED Requirements

### Requirement: Circular plot geometry
PPJ SHALL express a bounded first-slice angle for pie and doughnut charts and a
bounded center-hole size for doughnut charts.

#### Scenario: Authored doughnut uses deliberate geometry
- **WHEN** a doughnut declares `startAngle` and `holeSize`
- **THEN** the compiler writes native `c:firstSliceAng` and `c:holeSize` values
  and projection restores both PPJ properties

#### Scenario: Geometry is applied to an incompatible chart
- **WHEN** `startAngle` is used outside pie/doughnut or `holeSize` is used
  outside doughnut
- **THEN** compilation rejects before a PPTX is written

### Requirement: Source-bound circular edit
An imported editable pie or doughnut SHALL expose one bounded plot capability
and SHALL patch only the modeled scalar properties inside the existing
ChartPart.

#### Scenario: Imported doughnut geometry changes
- **WHEN** a hash-bound PPJ changes only `startAngle` or `holeSize`
- **THEN** the existing chart is patched, reimport restores the new values and
  unrelated OPC parts remain unchanged

#### Scenario: Plot capability is absent or stale
- **WHEN** a requested circular edit lacks the issued capability or no longer
  matches its source revision
- **THEN** compilation rejects without rebuilding or flattening the chart

