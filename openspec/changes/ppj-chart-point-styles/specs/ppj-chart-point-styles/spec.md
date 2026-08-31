## ADDED Requirements

### Requirement: Native sparse chart point styles
PPJ SHALL express bounded visual overrides for existing measured points without
changing chart or series topology.

#### Scenario: Authored chart highlights one point
- **WHEN** a supported native chart series declares sorted sparse point styles
- **THEN** the compiler writes canonical `c:dPt` nodes and projection restores
  the same point indexes, fills, strokes and legal explosion values

#### Scenario: Point style does not address real data
- **WHEN** an index is duplicated, unsorted, out of range or refers to a
  missing value
- **THEN** PPJ validation rejects before a PPTX is written

#### Scenario: Imported point style changes locally
- **WHEN** a source-bound PPJ changes only proved sparse point-style state
- **THEN** `setChartFill` patches the existing ChartPart, reimport restores the
  requested state and unrelated OPC parts remain unchanged

#### Scenario: Native point graph exceeds the bounded profile
- **WHEN** a native data point contains markers, picture options, 3D state,
  extensions, effects, unknown children or irregular paint/line topology
- **THEN** OfficeKit preserves the graph but does not expose an editable PPJ
  point-style capability
