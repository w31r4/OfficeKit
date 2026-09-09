## Purpose

Allow direct paragraph kerning threshold edits in source-bound PPJ with independent presence, native precision and preservation of unrelated source content.

## ADDED Requirements

### Requirement: Direct default kerning lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.kerning assignment, explicit zero, deletion, kerning-only defaultText/style wrapper removal and restoration under exact setTextParagraphStyle field authority. Values SHALL be finite within 0..768pt and round to native hundredths with ties to even.

#### Scenario: Assign remove and restore a threshold
- **WHEN** a threshold is assigned, removed directly or via its kerning-only wrapper, and restored
- **THEN** native kern and fresh PPJ presence follow the request, with zero retained explicitly and other defaults, direct runs, unknown attributes, paragraphs and non-target XML/ZIP content preserved

#### Scenario: Reject invalid or unauthorized changes
- **WHEN** a threshold is out of range, lacks exact field authority or accompanies unsupported default-style changes
- **THEN** compilation rejects without output

#### Scenario: Preserve an unmodeled source threshold
- **WHEN** native kern is outside the modeled range
- **THEN** no-op and unrelated scalar assignment/removal preserve it, while replacement through defaultText.kerning rejects without output
