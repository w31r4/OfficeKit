## ADDED Requirements

### Requirement: Native chart label and bubble formatting
PPJ SHALL express direct plot-level data-label number formats and bounded native
bubble scale semantics without inventing renderer-only pixel geometry.

#### Scenario: Authored chart formats labels
- **WHEN** chart data labels declare a number format
- **THEN** the compiler writes a direct source-unlinked native number format
  and projection restores the PPJ property

#### Scenario: Authored bubble chart controls scale
- **WHEN** a bubble chart declares `bubbleScale` and `bubbleSizeMode`
- **THEN** the compiler writes native bubble scale and size-representation
  values and projection restores both properties

#### Scenario: Bubble semantics are used on another chart
- **WHEN** bubble scale fields are used outside a bubble chart
- **THEN** compilation rejects before a PPTX is written

### Requirement: Native chart axis direction and lines
PPJ SHALL express axis reversal plus bounded axis-line and major-gridline intent.

#### Scenario: Authored value axis is reversed and styled
- **WHEN** an axis declares reverse orientation and bounded line styles
- **THEN** native axis scaling and line graphs reflect those values and
  projection restores them

#### Scenario: Imported axis formatting changes
- **WHEN** a hash-bound PPJ changes only proved axis direction or line fields
- **THEN** the existing ChartPart is patched, reimport restores the values and
  unrelated OPC parts remain unchanged

#### Scenario: Native line graph is outside the bounded profile
- **WHEN** an imported axis or grid line uses unsupported native children
- **THEN** OfficeKit preserves it but does not issue an editable axis capability
