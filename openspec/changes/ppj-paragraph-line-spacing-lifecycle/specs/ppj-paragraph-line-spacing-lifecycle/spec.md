## Purpose

Complete PPJ editing of direct paragraph line spacing while retaining native units, absence and surrounding imported content.

## ADDED Requirements

### Requirement: Complete line-spacing lifecycle

Ordinary text/shape paragraphs SHALL support mutually exclusive `text.paragraphs[].style.lineSpacing` (finite positive points up to 1584) and `lineSpacingMultiplier` (finite positive multiplier up to 132). Points SHALL round to hundredths and multipliers to 1/100000 with ties to even; rounded values SHALL be positive. Higher-priority paragraph styles SHALL select both unit and value. Assignment, unit switching, removal/restoration and removal of a style containing only line spacing SHALL round-trip through native PPTX and fresh PPJ. Invalid types, out-of-range values, zero, values that round to zero and simultaneous unit declarations SHALL reject without output.

#### Scenario: Set, switch, remove and restore
- **WHEN** an ordinary paragraph changes line spacing in either unit, including native minimum values, fractional values and maxima, then removes and restores it
- **THEN** native state and fresh PPJ SHALL retain the selected unit, rounded positive value or absence
- **AND** every changed unit field SHALL require its exact source-bound authority, including both fields when switching units

#### Scenario: Style precedence selects the unit
- **WHEN** a higher-priority style declares line spacing in a different unit
- **THEN** only that style's unit and value SHALL be applied

### Requirement: Preserve source paragraph state

Line-spacing edits SHALL preserve before/after spacing, unchanged native numeric spelling, other paragraph content, runs, neighboring paragraphs and non-target XML/ZIP. Duplicate slots/children, unknown attributes or descendants and missing, zero, invalid or out-of-range native values SHALL remain source-owned: no-op preserves bytes, unrelated scalar edits preserve their XML, and replacement rejects without output.

#### Scenario: Mixed or unmodeled spacing
- **WHEN** a line-spacing or unrelated scalar edit encounters mixed or unmodeled spacing content
- **THEN** only requested modeled state SHALL change and unmodeled replacement SHALL reject without output

### Requirement: Discoverability and preview evidence

Schema authority, Help, registry, generated references and text guidance SHALL agree on the lifecycle, positive-value and unit rules. Preview SHALL report the existing partial spacing support explicitly for both unit forms.

#### Scenario: Preview line spacing
- **WHEN** a paragraph declares line spacing in either unit
- **THEN** preview assessment SHALL expose a support-limit diagnostic without claiming host layout fidelity
