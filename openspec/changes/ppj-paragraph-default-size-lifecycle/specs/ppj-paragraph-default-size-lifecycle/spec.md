## Purpose

Provide independently authorized paragraph default font-size editing and removal in source-bound PPJ without changing other paragraph or run state.

## ADDED Requirements

### Requirement: Default size has independent presence

Ordinary text/shape paragraph defaultText.size SHALL support assignment, removal and restoration under its exact setTextParagraphStyle field authority while retaining fixed topology and non-target source state.

#### Scenario: Delete and restore size
- **WHEN** size or a size-only defaultText/style wrapper is omitted from a fresh source projection
- **THEN** native sz and freshly projected size are absent, and a later explicit size can restore them without altering other defaults, runs, paragraphs or unrelated ZIP members

#### Scenario: Precision and invalid values
- **WHEN** a finite size from 1 to 768pt is assigned
- **THEN** export and fresh projection use nearest-even hundredth-point rounding, and out-of-range sizes are rejected without output

#### Scenario: Missing authority
- **WHEN** size changes without its exact field authority, or an unrelated unsupported default field changes
- **THEN** compilation rejects the request without candidate output
