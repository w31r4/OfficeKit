## ADDED Requirements

### Requirement: PPJ SHALL express bounded mixed numeric Cartesian charts

When a combo chart contains scatter or bubble data, OfficeKit SHALL require all
series to use explicit numeric X coordinates and SHALL compile the shared
value/value coordinate system into editable native presentation objects.

#### Scenario: Bubble evidence and trend line share one numeric plot

- **WHEN** a valid PPJ combo contains a bubble series and a line series with
  complete finite X/Y values
- **THEN** OfficeKit SHALL compile bubbles and line segments at the same numeric
  scale with stable editable child IDs
- **AND** a second import SHALL recover the exact authored PPJ when the embedded
  program snapshot remains present

#### Scenario: A categorical disguise is rejected

- **WHEN** a numeric combo supplies shared categories, missing X coordinates or
  a categorical secondary axis
- **THEN** validation SHALL reject the program before package compilation

### Requirement: PPJ SHALL support bounded candlestick overlays

OfficeKit SHALL permit a valid OHLC/HLC series to be followed by a bounded set
of aligned line, area or column overlays and SHALL preserve the price marks as
the visual foreground.

#### Scenario: Moving average overlays an editable candlestick plot

- **WHEN** a candlestick chart adds an aligned line series
- **THEN** OfficeKit SHALL include the line values in the shared scale
- **AND** SHALL emit the line as editable native segments above the candle body

#### Scenario: Invalid overlay topology is rejected

- **WHEN** an overlay uses numeric X coordinates, carries OHLC channels, has a
  mismatched point count or requires a zero baseline outside the value domain
- **THEN** validation SHALL reject the program without flattening or guessing

### Requirement: Snapshot-free projection SHALL remain honest

OfficeKit SHALL project mixed numeric and candlestick overlay output without an
authored snapshot as editable native groups and SHALL NOT infer chart semantics
from arbitrary DrawingML children.

#### Scenario: Removing the authored snapshot does not invent chart data

- **WHEN** an OfficeKit-authored mixed chart is imported after its embedded PPJ
  snapshot has been removed
- **THEN** its native children SHALL remain discoverable and editable as a group
- **AND** OfficeKit SHALL NOT claim that the group is a semantic mixed chart
