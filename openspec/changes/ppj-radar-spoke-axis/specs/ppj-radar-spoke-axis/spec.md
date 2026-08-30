# PPJ radar spoke axis

## ADDED Requirements

### Requirement: Radar authors control one semantic spoke coordinate system

OfficeKit SHALL accept a bounded radar-only `spokeAxis` object and SHALL lower
its visibility, scale, numeric labels, spoke lines and spider grid to one
editable native standard-radar ChartPart.

#### Scenario: Build and recover a styled risk scale

- **GIVEN** a valid PPJ standard radar chart with a styled `spokeAxis`
- **WHEN** OfficeKit builds and reimports the presentation
- **THEN** the package contains paired native category/value axes with the
  declared scale, tick-label visibility and line styles
- **AND** the PPJ snapshot recovers the exact authored `spokeAxis`
- **AND** a snapshot-free canonical projection reports the same semantic
  coordinate-system fields when no unsupported distinction is present

### Requirement: Tick-label visibility is presence aware

OfficeKit SHALL represent hidden numeric chart tick labels without hiding the
axis grid and SHALL preserve an omitted native/default label position.

#### Scenario: Hide radar scale labels but keep the spider grid

- **GIVEN** `spokeAxis.label` is false and `spokeAxis.gridLine` is visible
- **WHEN** OfficeKit builds the chart
- **THEN** the value axis contains `c:tickLblPos val="none"`
- **AND** its major gridline container remains present

### Requirement: Ambiguous and unsupported radar axis state fails closed

OfficeKit SHALL reject semantic conflicts and SHALL not normalize unsupported
native radar-axis graphs into a guessed `spokeAxis`.

#### Scenario: A radar declares both semantic and generic axes

- **GIVEN** a PPJ radar element declares `spokeAxis` and `xAxis` or `yAxis`
- **WHEN** OfficeKit checks the program
- **THEN** validation fails before package mutation

#### Scenario: A source radar has a custom native label position

- **GIVEN** a third-party radar uses an unsupported high/low tick-label
  position or another unknown axis child graph
- **WHEN** OfficeKit projects the chart
- **THEN** it does not claim a fully editable semantic `spokeAxis`
- **AND** the source package remains preserved
