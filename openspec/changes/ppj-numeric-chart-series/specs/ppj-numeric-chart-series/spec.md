## ADDED Requirements

### Requirement: PPJ SHALL author numeric-X chart series

PPJ SHALL represent scatter-series X values and bubble-series X values and sizes
as finite typed vectors with the same point count as Y values.

#### Scenario: Bubble evidence chart is built

- **WHEN** a source-free bubble chart defines equal-length `xValues`, `values`
  and positive `bubbleSizes`
- **THEN** build emits an editable native bubble chart
- **AND** reimport projects equivalent numeric channels

### Requirement: Numeric chart topology SHALL fail before export

PPJ MUST reject missing, mismatched or inapplicable numeric-channel vectors.

#### Scenario: Category chart carries X values

- **WHEN** a column chart series defines `xValues`
- **THEN** validation fails with a chart-family diagnostic
- **AND** no PPTX is emitted

### Requirement: Existing source data capability SHALL remain narrow

The source-bound `setChartData` capability MUST NOT authorize changes to numeric
X or bubble-size caches.

#### Scenario: Bubble size is changed under setChartData

- **WHEN** a projected source-bound bubble chart changes `bubbleSizes`
- **THEN** compilation fails closed
- **AND** the source package remains unchanged
