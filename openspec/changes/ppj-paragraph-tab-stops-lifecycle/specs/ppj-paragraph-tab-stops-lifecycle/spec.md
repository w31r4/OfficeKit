## Purpose

Represent and independently edit a paragraph's ordered direct tab stops, preserving source content and making removal, precision and unsupported layout behavior explicit.

## ADDED Requirements

### Requirement: Bounded ordered tab stop choice
`text.paragraphs[].style.tabStops` SHALL accept at most 32 entries with finite nonnegative point positions in the native signed-32-bit EMU range, rounded to the nearest EMU with ties to even. Positions SHALL remain strictly increasing after rounding. Alignment SHALL be left, center, right or decimal, with omitted alignment selecting left. `tabStops` and `noTabStops: true` SHALL be mutually exclusive in each style layer; the highest-priority layer containing either SHALL select the complete choice. False, wrong types, out-of-range positions, invalid alignments, unordered/duplicate positions and quantization collisions SHALL reject.

#### Scenario: Authored precision and precedence
- **WHEN** a bounded tab list is authored, including zero or the largest native position, with lower-priority tab settings
- **THEN** native projection SHALL recover the selected ordered values at EMU precision, and a higher-priority clear SHALL suppress the lower-priority list

### Requirement: Complete independent source lifecycle
Ordinary source-bound text/shape owners SHALL support tab-list addition, assignment, clearing, removal and restoration. An empty array, removal of a previously projected list, or `noTabStops: true` SHALL remove the modeled direct list and project back to absence. Each changed PPJ property SHALL require exact authority, including both sides of a choice switch. Only changed paragraphs' tab lists SHALL be patched; other paragraph state, runs, neighbors and non-target XML/ZIP SHALL remain. No-op SHALL retain exact source bytes.

#### Scenario: Remove and restore from actual source
- **WHEN** a list is removed from a fresh projection of its original source, or restored from a fresh projection of the resulting candidate
- **THEN** native XML and re-projection SHALL agree with the requested list or absence, and unrelated content SHALL remain

#### Scenario: Missing field authority
- **WHEN** any changed tabStops or noTabStops property lacks its exact source capability
- **THEN** the request SHALL reject without a candidate

### Requirement: Unmodeled lists retain source fidelity
Malformed numeric or alignment tokens, unknown attributes/children, duplicate lists and unsupported order/budgets SHALL remain source-owned as a whole. They SHALL NOT be projected as partial usable tab lists or prevent independent modeled text/style editing. No-op and unrelated edits SHALL preserve their XML; explicit replacement or clearing SHALL reject. Equivalent modeled lists SHALL retain native number spelling and omitted default alignment.

#### Scenario: Imported malformed list
- **WHEN** source tab state contains an invalid number, foreign metadata, duplicate list or out-of-order stops
- **THEN** projection SHALL omit the list without a parsing failure, unrelated edits SHALL retain it, and explicit replacement or clearing SHALL fail

### Requirement: Tab layout remains explicitly reviewable
Schema, Help, capability registry, generated references and text guidance SHALL describe the same units, ordering, choice and removal semantics. Preview SHALL retain explicit diagnostics for unresolved tab-stop layout instead of claiming exact host typography.

#### Scenario: Preview with explicit tabs
- **WHEN** a paragraph with direct tab stops is previewed
- **THEN** unresolved tab layout SHALL remain visible in diagnostics and SHALL NOT be classified as fully supported
