## Purpose

Allow PPJ authors to express native logarithmic value-axis scaling and edit it without changing chart data or unrelated package content.

## ADDED Requirements

### Requirement: Optional logarithmic value-axis base

PPJ SHALL support optional axis logBase as a finite number from 2 through 1000 inclusive or a size grammar token resolving to that range. It SHALL support value axes, numeric scatter/bubble x axes, and combo secondary value axes. Category axes SHALL reject the field. Explicit logarithmic min/max values SHALL be positive.

Radar charts SHALL accept the same field on spokeAxis and recover it through the native value axis, including after removal and subsequent addition.

#### Scenario: Author and project logarithmic scaling
- **WHEN** a supported value axis specifies logBase directly or through a size token
- **THEN** native scaling contains that base and projection without embedded PPJ recovers it

#### Scenario: Invalid scale
- **WHEN** logBase is out of range, used on a category axis, or combined with nonpositive explicit limits
- **THEN** compilation fails with a diagnostic

### Requirement: Source-bound logarithmic scaling edits

The compiler SHALL add, change, and remove logBase on an editable native value axis. Removing it SHALL restore linear scaling. These edits SHALL change only the target ChartPart, retaining all unrelated package bytes.

#### Scenario: Edit and remove base
- **WHEN** a projected axis base is changed or removed
- **THEN** native XML and fresh projection reflect the new value or absence, with unrelated package entries unchanged

### Requirement: Ambiguous native scaling remains protected

Native charts with duplicate, malformed, out-of-range, or decorated logBase nodes SHALL NOT advertise editable chart axes and SHALL preserve their source content in a no-op round trip.

#### Scenario: Ambiguous native owner
- **WHEN** an imported axis contains multiple logBase nodes or an unsupported logBase attribute or child
- **THEN** projection withholds axis editing and a no-op compile retains the native chart bytes

### Requirement: Shared spreadsheet compatibility

The JavaScript spreadsheet adapter SHALL retain imported logarithmic scaling when rewriting other supported chart fields.

#### Scenario: Edit an imported workbook chart title
- **WHEN** a logarithmic chart imported through the shared codec receives a title edit
- **THEN** its original logBase remains present in the outgoing wire axis
