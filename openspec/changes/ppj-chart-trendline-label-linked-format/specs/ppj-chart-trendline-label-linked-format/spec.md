## Purpose

Preserve and edit the complete native number-format linkage state of trendline labels through PPJ without inventing workbook or host behavior.

## ADDED Requirements

### Requirement: Explicit and omitted native format flags

PPJ SHALL expose `label.numberFormatSourceLinked` as boolean or null, requiring `numberFormat`. True SHALL write native sourceLinked=1, false or an omitted PPJ field SHALL write 0, and null SHALL omit that native attribute. Projection SHALL retain true and null and canonicalize false to an omitted PPJ field. Existing wire messages without the new field SHALL continue emitting explicit false.

#### Scenario: Three native states round trip
- **WHEN** an ordinary or categorical combo trendline label is authored with true, false or null and a valid number-format string/token
- **THEN** the requested native flag state SHALL survive fresh projection, with false canonically omitted in PPJ

### Requirement: Format linkage lifecycle preserves other data

Supported labels SHALL allow toggling linkage, retaining native attribute absence, changing format, removing the complete format pair and recreating it through the existing analytics operation. No-op SHALL preserve all input bytes; changes SHALL retain other label/chart state and non-target ZIP entries. Documentation SHALL distinguish native flag preservation from host interpretation and workbook synchronization.

#### Scenario: Edit and remove a linked format
- **WHEN** a source-bound label format is toggled, changed, removed and recreated
- **THEN** only the target ChartPart SHALL change and fresh projection SHALL match requested canonical state
- **AND** changing another label property SHALL retain a native omitted sourceLinked attribute

### Requirement: Invalid format states cannot be silently discarded

Invalid PPJ types, linkage without a format, unknown wire enum values and malformed native flags SHALL fail closed. Unsupported source labels SHALL preserve exact bytes on no-op and reject analytics replacement without output.

#### Scenario: Bad linkage values are rejected
- **WHEN** a PPJ field contains a string instead of boolean/null, lacks numberFormat, or an imported native flag is invalid
- **THEN** compilation or the attempted analytics edit SHALL fail without emitting a file
