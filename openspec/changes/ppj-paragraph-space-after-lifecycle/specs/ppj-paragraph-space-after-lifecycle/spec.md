## Purpose

Complete PPJ editing of the direct space after a paragraph while retaining native units, absence and surrounding imported content.

## ADDED Requirements

### Requirement: Complete space-after lifecycle

Ordinary text/shape paragraphs SHALL support mutually exclusive `text.paragraphs[].style.spaceAfter` (finite 0..1584pt) and `spaceAfterMultiplier` (finite 0..132). Points SHALL round to hundredths and multipliers to 1/100000 with ties to even. Higher-priority paragraph styles SHALL select both unit and value. Assignment, switching units, removal/restoration and removal of a style containing only that spacing SHALL round-trip through native PPTX and fresh PPJ. Zero SHALL remain distinct from absence and retain its unit. Invalid types, bounds and simultaneous unit declarations SHALL reject without output.

#### Scenario: Set, switch, remove and restore
- **WHEN** an ordinary paragraph changes space after in either unit, including zero, fractional values and endpoints, then removes and restores it
- **THEN** native state and fresh PPJ SHALL retain the selected unit, rounded value or absence
- **AND** every changed unit field SHALL require its exact source-bound authority, including both fields when switching units

#### Scenario: Style precedence chooses the unit
- **WHEN** a higher-priority style declares space after in a different unit
- **THEN** only that style's unit and value SHALL be applied

### Requirement: Preserve source paragraph state

Space-after edits SHALL preserve space before, line spacing, unchanged native numeric spelling, other paragraph content, runs, neighboring paragraphs and non-target XML/ZIP. Duplicate slots/children, unknown attributes or descendants and missing/invalid/out-of-range native values SHALL remain source-owned: no-op preserves bytes, unrelated scalar edits preserve their XML, and replacement rejects without output.

#### Scenario: Mixed or unmodeled spacing
- **WHEN** a space-after or unrelated scalar edit encounters mixed or unmodeled spacing content
- **THEN** only the requested modeled state SHALL change and unmodeled replacement SHALL reject without output

### Requirement: Discoverability and preview evidence

Schema authority, Help, registry, generated references and text guidance SHALL agree on this lifecycle and its unit rules. Preview SHALL report the existing partial spacing support explicitly, including zero values.

#### Scenario: Preview space after
- **WHEN** a paragraph declares space after in either unit
- **THEN** preview assessment SHALL expose a support-limit diagnostic without claiming host layout fidelity
