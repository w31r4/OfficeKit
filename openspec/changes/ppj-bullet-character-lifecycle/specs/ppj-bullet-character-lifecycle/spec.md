## Purpose

Let PPJ author and replace one character list marker without losing its original source styling, adjacent text or unmodeled presentation content.

## ADDED Requirements

### Requirement: One valid character scalar

`bullet.character` SHALL be exactly one Unicode scalar compatible with XML. Supplementary-plane scalars SHALL count as one character. Empty values, multiple scalars, lone UTF-16 surrogates and XML-incompatible characters SHALL be rejected. The field SHALL remain required for `type: character`.

#### Scenario: Author ordinary and supplementary characters

- **WHEN** PPJ authors a valid one-scalar marker, including a supplementary-plane symbol
- **THEN** native export and fresh projection SHALL preserve that exact character

#### Scenario: Invalid or missing character

- **WHEN** character selection is missing, null, empty, multiple scalars or invalid for XML
- **THEN** compilation SHALL reject it without emitting a candidate

### Requirement: Preserve source during character edits

Existing character-marked text/shape paragraphs SHALL support replacing and restoring `bullet.character` under its exact field authority. The edit SHALL preserve bullet font/color/size, other paragraph properties, runs, neighbors, unknown marker XML and non-target package content. No-op requests SHALL retain source bytes. Character-only authority SHALL NOT convert another marker kind or replace an unmodeled marker.

#### Scenario: Replace and restore an existing symbol

- **WHEN** a projected character marker changes and is restored from the actual candidate and fresh projection
- **THEN** only the target marker's character attribute SHALL change and reproject to each requested value

#### Scenario: Unmodeled parseable marker or missing authority

- **WHEN** source marker choices are ambiguous or their character is unmodeled, or the request lacks the exact character authority
- **THEN** replacement SHALL fail closed while no-op and unrelated edits preserve the parseable source content

### Requirement: Preview evidence is bounded

Field references SHALL describe single-scalar semantics and the source editing boundary. Preview SHALL preserve the marker in its supported direct-style profile and SHALL diagnose incomplete glyph/layout or unresolved marker styling explicitly.

#### Scenario: Direct character marker in preview

- **WHEN** a marker has the direct style and indentation needed by the bounded preview profile
- **THEN** preview SHALL retain the requested character while keeping its text-layout limitation
