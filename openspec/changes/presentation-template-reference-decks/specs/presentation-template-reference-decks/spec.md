# Presentation template reference decks

## Requirements

### Requirement: authored reference deck

The system SHALL package a presentation template only when its spec names an
absolute `.pptx` `referencePath` containing a structurally valid Office Open
XML package. The published bytes SHALL be copied to `assets/reference.pptx`.

#### Scenario: valid authored deck

- **WHEN** the creator receives a valid OfficeKit-authored reference deck and
  the required visual examples
- **THEN** it SHALL publish schema v4 metadata and a hash-bound reference asset

#### Scenario: malformed or missing deck

- **WHEN** the creator receives a missing, non-PPTX, malformed, or oversized
  reference path
- **THEN** it SHALL fail before publishing a template tree

### Requirement: source separation

The system SHALL keep external source decks and their extracted evidence out of
the published template. The authored reference deck SHALL not be interpreted
as a fixed layout registry.

### Requirement: discoverable reference

Search SHALL return an absolute `referencePath` for schema v4 candidates and
verify its declared SHA-256 before returning the candidate. Schema v3 entries
SHALL be rejected and rebuilt with the Creator; no presentation compatibility
path is retained.

### Requirement: restoration threshold

Migration evidence SHALL record independent visual and functional restoration
indices. A template SHALL be labelled restored only when both indices are at
least 95 and the cited render, inspect, edit, and re-import evidence exists.
