# PPJ standard radar chart

## ADDED Requirements

### Requirement: Authored PPJ radar charts compile to native editable charts

OfficeKit SHALL compile `chartType: "radar"` to the bounded native standard
radar profile while preserving PPJ element identity, categories, series,
strokes, markers, labels, axes and accessibility state.

#### Scenario: Build and reproject a standard radar chart

- **GIVEN** a valid source-free PPJ radar chart with finite values
- **WHEN** OfficeKit builds the PPJ and projects the resulting PPTX
- **THEN** the PPTX contains a standard native radar plot
- **AND** the projected PPJ reports `chartType: "radar"`
- **AND** stable IDs and semantic values remain available for continuation

### Requirement: Unsupported radar topology fails closed

OfficeKit SHALL NOT flatten or guess at radar variants outside its bounded
native profile.

#### Scenario: Import a non-standard radar plot

- **GIVEN** a third-party PPTX containing an unsupported radar style or unknown
  radar child graph
- **WHEN** OfficeKit projects the file
- **THEN** the source graph remains preserved
- **AND** unsupported mutations are not issued as typed editable capabilities

### Requirement: Radar continuation uses the existing chart edit contract

OfficeKit SHALL allow bounded data and supported series-style changes without
rebuilding unrelated package parts.

#### Scenario: Change one radar value

- **GIVEN** an editable source-bound standard radar chart
- **WHEN** one existing series value changes without topology changes
- **THEN** OfficeKit patches the bound ChartPart
- **AND** unrelated package parts remain unchanged
