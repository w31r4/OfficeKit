## Purpose

Allow PPJ users to edit direct paragraph levels without losing source formatting, native presence or unrelated document content.

## ADDED Requirements

### Requirement: Direct paragraph level lifecycle

Ordinary source-bound text and shape owners SHALL independently assign, remove and restore `text.paragraphs[].style.level` under its exact field authority. Values SHALL be integers from 0 through 8. Explicit zero SHALL retain native `lvl=0`, while field or level-only style removal SHALL remove the direct attribute. A direct authored paragraph level SHALL override its owner default.

#### Scenario: Change one paragraph level
- **WHEN** a user sets, removes or restores the direct level of one paragraph
- **THEN** fresh projection SHALL retain the value or absence while preserving runs, neighbors, other properties and non-target XML/ZIP content

#### Scenario: Reject invalid or unauthorized changes
- **WHEN** a request lacks level authority or supplies a noninteger, out-of-range or wrong-typed value
- **THEN** compilation SHALL reject it without output

### Requirement: Preserve unmodeled source levels

Malformed or out-of-range native levels SHALL remain source-owned. No-op and unrelated scalar edits SHALL preserve them, while replacing them through the modeled level field SHALL reject without output. Unchanged valid native level spelling SHALL also remain unchanged.

#### Scenario: Unknown source level during an unrelated edit
- **WHEN** an imported paragraph with an unmodeled native level has its default bold assigned or removed
- **THEN** its original level attribute SHALL survive and remain absent from modeled PPJ

### Requirement: Expose the field and its layout limits

Help, schema, capability metadata and generated Agent guidance SHALL describe the field lifecycle. Preview SHALL retain explicit partial diagnostics for unresolved level-driven layout.

#### Scenario: Discover a level field
- **WHEN** a user reads the generated reference or previews a level-bearing paragraph
- **THEN** its 0..8 range, deletion semantics and incomplete layout evidence SHALL remain explicit
