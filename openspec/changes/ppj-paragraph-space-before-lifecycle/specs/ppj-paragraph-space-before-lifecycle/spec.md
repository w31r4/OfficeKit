## Purpose

Allow PPJ to edit the direct space before a paragraph in either native unit, with complete presence semantics and preservation of surrounding imported content.

## ADDED Requirements

### Requirement: Complete space-before value and lifecycle

Ordinary text/shape paragraph styles SHALL support mutually exclusive `spaceBefore` (finite 0..1584pt) and `spaceBeforeMultiplier` (finite 0..132). Points SHALL round to hundredths and multipliers to 1/100000 with ties to even. Higher-priority paragraph styles SHALL select both the unit and value. Explicit zero SHALL retain its native unit. Assignment, unit switching, removal/restoration and removal of a style containing only that spacing SHALL round-trip through PPTX and fresh PPJ. Both fields together, invalid types and out-of-range values SHALL reject without output.

#### Scenario: Change unit, clear and restore
- **WHEN** an ordinary paragraph changes between points and multiplier, including zero and rounded fractional values, then removes and restores its space before
- **THEN** the native node and fresh PPJ SHALL retain the chosen unit, rounded value or absence
- **AND** each changed unit field SHALL require its exact source-bound authority, so a unit switch requires authority for both fields

#### Scenario: Style precedence selects a unit
- **WHEN** a higher-priority style declares a different space-before unit than a lower-priority style
- **THEN** only the higher-priority unit and value SHALL be used

### Requirement: Preserve imported paragraph state

Edits SHALL preserve line spacing, space after, other paragraph content, native spelling of unchanged spacing, direct runs, neighboring paragraphs and non-target XML/ZIP. Duplicate slots/children, unknown attributes or nested content, missing/invalid/out-of-range native values SHALL remain source-owned: no-op preserves bytes, unrelated scalar edits preserve their XML, and replacement rejects without output.

#### Scenario: Mixed and unmodeled spacing
- **WHEN** a paragraph space-before edit or unrelated scalar edit encounters other source state
- **THEN** only the requested modeled state SHALL change
- **AND** unmodeled space before SHALL remain preserved and reject replacement

### Requirement: Metadata and preview evidence

Schema authority, Help, registry, generated references and text guidance SHALL describe units, bounds, exclusivity and lifecycle consistently. Preview assessment SHALL report the existing partial spacing support explicitly, including zero.

#### Scenario: Preview a space-before value
- **WHEN** a paragraph declares space before in either unit
- **THEN** preview assessment SHALL expose a support-limit diagnostic without claiming host layout fidelity
