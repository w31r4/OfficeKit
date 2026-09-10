## Purpose

Represent the direct East Asian line-break setting in PPJ while preserving its
boolean presence and unrelated source content during paragraph edits.

## ADDED Requirements

### Requirement: Boolean state and independent semantics

PPJ SHALL accept boolean `text.paragraphs[].style.eastAsianLineBreak` to control
application of East Asian typography and line-start/end rules. True, false and
absence SHALL survive export and snapshot-free re-import separately, using
existing paragraph style precedence. Absence SHALL NOT synthesize a direct
default. Latin word breaking, hanging punctuation, text-box wrap, explicit break
inlines, indentation and direction SHALL remain independent. Non-booleans SHALL reject.

#### Scenario: Explicit false overrides lower-priority true
- **WHEN** a higher-priority paragraph style declares false and a lower-priority style declares true
- **THEN** export/re-import retains false; removing the higher-priority declaration selects true

### Requirement: Independent source lifecycle

Ordinary recognized text/shape owners SHALL support add/set/remove/restore under
exact `text.paragraphs[].style.eastAsianLineBreak` authority. Removing a style
containing only the field SHALL clear the modeled direct attribute. Unchanged
boolean spelling, other properties, runs, breaks, neighboring paragraphs and
non-target XML/ZIP SHALL remain preserved.

#### Scenario: Original-source deletion and restoration
- **WHEN** a fresh projection of original bytes removes the field and a later edit restores it
- **THEN** re-import reflects both states and only the target attribute changes

#### Scenario: Missing exact authority
- **WHEN** the field changes without its exact capability
- **THEN** compilation rejects without an output artifact

### Requirement: Unknown native state stays source-owned

Native 0/1/false/true SHALL be recognized. Unknown values SHALL be omitted from
semantic state, retained by no-op and unrelated edits, and reject replacement.

#### Scenario: Unknown token and an independent default-bold edit
- **WHEN** source contains an unknown East-Asian-line-break token
- **THEN** no-op preserves the original file and changing default bold preserves the token
- **AND** assigning eastAsianLineBreak rejects without an artifact

### Requirement: Discovery and preview limits

Help, registry, schema and Agent guidance SHALL describe the same lifecycle.
Preview SHALL explicitly diagnose both boolean values as partial/unmapped until
measured East Asian line breaking and kinsoku evaluation are implemented.

#### Scenario: Explicit false in preview
- **WHEN** preview assesses eastAsianLineBreak false or true
- **THEN** it reports the unmapped field without claiming host line-break fidelity
