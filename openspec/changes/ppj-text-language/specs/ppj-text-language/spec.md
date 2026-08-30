# PPJ text-language specification

## ADDED Requirements

### Requirement: Authored run language

PPJ SHALL compile valid run and default-run language tags to editable native
DrawingML language attributes.

#### Scenario: Mixed Chinese and English text

- **WHEN** adjacent PPJ runs declare `zh-CN` and `en-US`
- **THEN** native output and re-import preserve both explicit language tags

### Requirement: Imported direct language

A safe direct run language SHALL project into PPJ and issue one source-bound
`fontLanguage` leaf.

#### Scenario: Continue an imported language-tagged run

- **WHEN** an Agent changes an issued `fontLanguage` leaf with matching hashes
- **THEN** only the target `a:rPr/@lang` token changes and second import recovers
  the requested tag

### Requirement: Honest failure boundary

Malformed, inherited or ambiguous language state SHALL remain source owned and
shall not be normalized or guessed.

#### Scenario: Unsupported source language graph

- **WHEN** a source run cannot prove one canonical direct language token
- **THEN** PPJ preserves the source and omits the editing leaf
