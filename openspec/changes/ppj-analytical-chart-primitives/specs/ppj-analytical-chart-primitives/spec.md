## ADDED Requirements

### Requirement: PPJ SHALL expose bounded analytical chart semantics

The language MUST represent supported axis bounds and titles, marker state,
label content, trendlines, and error bars as typed finite JSON.

#### Scenario: Evidence chart is authored

- **WHEN** a source-free PPJ defines bounded analytical chart state
- **THEN** build emits editable native chart semantics
- **AND** reimport projects equivalent recognized state

### Requirement: Compatibility chart fields SHALL remain unambiguous

Legacy scalar marker and data-label fields MUST remain valid, while conflicting
legacy and structured spellings MUST fail before output.

#### Scenario: Two label spellings conflict

- **WHEN** a chart defines incompatible legacy and structured label state
- **THEN** build returns an unsupported-feature diagnostic
- **AND** no PPTX is emitted

### Requirement: Source-bound data edits SHALL NOT mutate chart style

An issued data capability MUST own only the recognized series names, categories,
and values. Other series and chart state MUST remain equal to the source PPJ.

#### Scenario: Agent changes marker under setChartData

- **WHEN** a source-bound PPJ changes a marker while presenting the mutation as chart data
- **THEN** compilation fails closed
- **AND** the source package remains unchanged

