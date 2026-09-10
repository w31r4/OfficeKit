## Purpose

Represent the direct Latin word-break setting in PPJ while preserving explicit
boolean presence and non-target source content during paragraph edits.

## ADDED Requirements

### Requirement: Boolean state and precedence

PPJ SHALL accept boolean `text.paragraphs[].style.latinLineBreak`. True permits
word-internal Latin line breaking and false disables that setting. Export and
snapshot-free re-import SHALL preserve true, false and absence separately, using
existing paragraph style precedence. Text-box wrap, explicit break inlines,
hanging punctuation, indentation and writing direction SHALL remain independent.
Non-boolean values SHALL reject. Absence SHALL NOT synthesize a direct default.

#### Scenario: Explicit false overrides a lower-priority true
- **WHEN** a higher-priority paragraph style declares false and a lower-priority style declares true
- **THEN** export/re-import retains false; removing the higher-priority declaration selects true

### Requirement: Independent source lifecycle

Ordinary recognized text/shape owners SHALL support add/set/remove/restore under
exact `text.paragraphs[].style.latinLineBreak` authority. Removing a style containing
only this field SHALL clear the modeled direct attribute. Unchanged boolean
spelling, other paragraph properties, runs, explicit breaks, neighbors and
non-target XML/ZIP SHALL remain preserved.

#### Scenario: Original-source deletion and restoration
- **WHEN** a fresh projection of original source bytes removes the field and a later edit restores it
- **THEN** re-import reflects both states and only the target attribute changes

#### Scenario: Missing exact authority
- **WHEN** the field changes without its exact capability
- **THEN** compilation rejects without an output artifact

### Requirement: Unknown native values stay source-owned

Native 0/1/false/true SHALL be recognized. Unknown values SHALL be omitted from
semantic state, retained on no-op and unrelated edits, and reject replacement
by a modeled value.

#### Scenario: Unknown token and an independent default-bold edit
- **WHEN** source contains an unknown Latin-line-break token
- **THEN** no-op preserves the original file and changing default bold preserves that token
- **AND** assigning latinLineBreak rejects without an artifact

### Requirement: Discovery and preview limits

Help, registry, schema and Agent guidance SHALL describe the same lifecycle.
Preview SHALL explicitly diagnose both boolean values as partial/unmapped until
measured word breaking is implemented.

#### Scenario: Explicit false in preview
- **WHEN** preview assesses latinLineBreak false or true
- **THEN** it reports the unmapped field without claiming host word-breaking fidelity
