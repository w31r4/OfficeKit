## Purpose

Provide independent direct paragraph letter-spacing edits in source-bound PPJ while preserving native precision, optional presence and unrelated source content.

## ADDED Requirements

### Requirement: Direct paragraph default letter spacing lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.letterSpacing assignment, explicit zero, deletion, spacing-only defaultText/style wrapper removal and restoration under exact setTextParagraphStyle field authority. Values SHALL be finite within -768..768pt and round to native hundredths with ties to even.

#### Scenario: Assign remove and restore signed spacing
- **WHEN** a positive, negative or zero spacing is assigned, removed directly or via its spacing-only wrapper, and restored
- **THEN** native spc and fresh PPJ presence follow the request while other defaults, direct runs, unknown attributes, paragraphs and non-target XML/ZIP content remain unchanged

#### Scenario: Reject invalid or unauthorized changes
- **WHEN** spacing is out of range, lacks exact field authority or accompanies unsupported default-style changes
- **THEN** compilation rejects without output

#### Scenario: Preserve unmodeled native spacing
- **WHEN** source spc is outside the modeled range
- **THEN** no-op and unrelated scalar assignment/removal preserve it, while replacement through defaultText.letterSpacing rejects without output
