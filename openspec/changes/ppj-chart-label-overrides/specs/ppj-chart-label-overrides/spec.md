## ADDED Requirements

### Requirement: Native series and point data-label overrides
PPJ SHALL express bounded native label defaults for one chart series and sparse
overrides for existing non-missing points without changing chart topology.

#### Scenario: Authored chart uses series and point label policy
- **WHEN** one native chart series declares label defaults and sorted sparse
  point overrides
- **THEN** the compiler writes canonical `c:ser/c:dLbls` and `c:dLbl` state
  and projection restores the same PPJ semantics

#### Scenario: Point override does not address real data
- **WHEN** a point label index is duplicated, unsorted, out of range or refers
  to a missing value
- **THEN** PPJ validation rejects before a PPTX is written

#### Scenario: Imported label state changes locally
- **WHEN** a source-bound PPJ changes only proved series or point label state
- **THEN** `setChartLabels` patches the existing ChartPart, reimport restores
  the requested state and unrelated OPC parts remain unchanged

#### Scenario: Native label graph exceeds the bounded profile
- **WHEN** an imported label uses custom text, manual layout, shape/effect
  state, leader-line graphs, extensions or a source-linked number format
- **THEN** OfficeKit preserves the native graph but does not expose it as an
  editable PPJ label capability
