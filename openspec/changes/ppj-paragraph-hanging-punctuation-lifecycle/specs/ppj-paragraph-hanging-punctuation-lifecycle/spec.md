## Purpose

Represent direct hanging-punctuation settings in PPJ while preserving explicit
boolean presence and unrelated content during source-bound paragraph edits.

## ADDED Requirements

### Requirement: Boolean state and precedence

PPJ SHALL accept boolean `text.paragraphs[].style.hangingPunctuation`, preserving
true, false and absence through export and snapshot-free re-import. Existing
paragraph style precedence SHALL apply. Hanging indent, margins, alignment and
writing direction SHALL remain independent. Non-boolean values SHALL reject.

#### Scenario: Explicit false overrides an inherited true
- **WHEN** a higher-priority paragraph style declares false and a lower-priority style declares true
- **THEN** export/re-import retains false; removing the higher-priority declaration selects true

### Requirement: Independent source lifecycle

Ordinary recognized text/shape owners SHALL support adding, setting, removing
and restoring the field under exact `text.paragraphs[].style.hangingPunctuation`
authority. Removing a style containing only this field SHALL clear the modeled
direct attribute. Unchanged native boolean spelling, other paragraph attributes,
runs, neighboring paragraphs and non-target XML/ZIP SHALL be preserved.

#### Scenario: Original-source removal and restoration
- **WHEN** an edit removes the field from a fresh projection of original source bytes and a subsequent edit restores it
- **THEN** re-import reflects both states and only the target attribute changes

#### Scenario: Missing exact authority
- **WHEN** an edit changes this field without its exact capability
- **THEN** compilation rejects without an output artifact

### Requirement: Unknown native state is preserved

Native boolean spellings 0/1/false/true SHALL be recognized. Unknown native
values SHALL be omitted from semantic projection, retained on no-op and unrelated
edits, and reject replacement by a modeled value.

#### Scenario: Unknown token with an independent default-bold edit
- **WHEN** a source paragraph has an unknown hanging-punctuation token
- **THEN** no-op preserves the original file and changing default bold preserves the token
- **AND** setting hangingPunctuation rejects without an artifact

### Requirement: Discovery and preview diagnostics

Help, registry, schema and Agent guidance SHALL describe the same field lifecycle.
Preview SHALL explicitly report partial/unmapped handling for both boolean values
until measured punctuation placement and line breaking are implemented.

#### Scenario: Explicit disabled punctuation setting
- **WHEN** preview assesses hangingPunctuation false or true
- **THEN** it reports the unmapped field without claiming measured punctuation layout
