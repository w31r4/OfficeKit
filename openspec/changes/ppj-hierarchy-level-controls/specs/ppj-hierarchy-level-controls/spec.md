## ADDED Requirements

### Requirement: PPJ SHALL separate hierarchy data depth from display depth

OfficeKit SHALL permit a bounded treemap or sunburst series to retain its full
validated hierarchy while declaring how many leading levels are rendered.

#### Scenario: A dense hierarchy reveals only its first two levels

- **WHEN** a valid treemap or sunburst series declares `levels: 2`
- **THEN** OfficeKit SHALL keep every declared node in the authored PPJ
- **AND** SHALL compile only roots and direct children into editable native
  elements
- **AND** SHALL allocate the visible geometry across those two levels

#### Scenario: Levels are limited to hierarchy charts

- **WHEN** another chart family declares `levels`, or a sunburst declares more
  than six levels
- **THEN** validation SHALL reject the program before package compilation

### Requirement: Hierarchy level recovery SHALL remain exact and honest

OfficeKit SHALL restore authored `levels` from its embedded PPJ snapshot and
SHALL NOT infer hidden descendants from a snapshot-free native group.

#### Scenario: Snapshot removal exposes only native visible state

- **WHEN** an authored limited hierarchy is imported without its PPJ snapshot
- **THEN** the visible native elements SHALL remain editable as a group
- **AND** OfficeKit SHALL NOT claim that omitted semantic descendants exist
