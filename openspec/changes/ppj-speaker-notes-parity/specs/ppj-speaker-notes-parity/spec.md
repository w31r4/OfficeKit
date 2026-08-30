## ADDED Requirements

### Requirement: Structured imported speaker notes
The PPTX projector SHALL preserve every supported paragraph, run, and formatting
value in imported speaker notes instead of flattening the notes to plain text.

#### Scenario: Multi-run imported notes
- **WHEN** a source slide contains a relationship-free notes body with multiple
  formatted runs
- **THEN** PPJ represents it with `pages[].notes.paragraphs[].runs[]` and stable
  local paragraph and run IDs

### Requirement: Evidence-bound notes capability
The projector SHALL issue `setNotes` only when the existing notes body is
editable or the source graph can safely accept one canonical notes leaf.

#### Scenario: Irregular notes body
- **WHEN** imported notes contain a field, hyperlink, picture bullet, or unknown
  relationship
- **THEN** the semantic text may be preserved but PPJ does not advertise a
  notes mutation capability

### Requirement: Local source-bound notes edit
A capable PPJ notes text change SHALL lower through the existing native notes
codec while preserving source-owned graph and formatting state.

#### Scenario: Edit one rich run
- **WHEN** an Agent changes one run's text without changing paragraph/run counts
  or styles
- **THEN** build changes only the target notes part and required package metadata,
  and reimport recovers the edited run and unchanged formatting

### Requirement: Bounded notes addition
A source page whose native graph proves speaker notes addable SHALL accept one
plain-text PPJ notes value.

#### Scenario: Add plain speaker notes
- **WHEN** a capable source page without NotesSlide receives a string notes value
- **THEN** build creates the canonical bounded notes graph and reimport recovers
  the value

### Requirement: Unsupported notes mutations fail closed
PPJ SHALL reject source-bound notes deletion, plain/rich conversion, rich style
or topology changes, and addition of structured notes to an absent notes part.

#### Scenario: Change imported notes styling
- **WHEN** an Agent changes a run style as well as its text
- **THEN** build rejects the unsupported style mutation before package export

### Requirement: Agent discoverability
The generated PPJ reference and presentation guidance SHALL explain how notes
support live delivery and which imported mutations remain bounded.

#### Scenario: Agent prepares a live deck
- **WHEN** an Agent reads the PPJ text or delivery guidance
- **THEN** it can find `pages[].notes`, rich notes structure, and the source-bound
  mutation boundary without reading protobuf or C# source
