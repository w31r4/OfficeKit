## Purpose

Enable independently authorized direct paragraph baseline edits in source-bound PPJ while retaining percentage precision, optional presence and unrelated source content.

## ADDED Requirements

### Requirement: Direct paragraph default baseline lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.baseline assignment, explicit zero, deletion, baseline-only defaultText/style wrapper removal and restoration under exact setTextParagraphStyle field authority. Values SHALL be finite within -400..400 percent and round to native thousandths with ties to even, retaining canonical explicit zero.

#### Scenario: Assign remove and restore signed baseline
- **WHEN** a positive, negative or zero baseline is assigned, removed directly or through its baseline-only wrapper, and restored
- **THEN** native baseline and fresh PPJ presence follow the request while other defaults, direct runs, unknown attributes, paragraphs and non-target XML/ZIP content remain unchanged

#### Scenario: Reject invalid or unauthorized changes
- **WHEN** baseline is out of range, lacks exact field authority or accompanies unsupported default-style changes
- **THEN** compilation rejects without output

#### Scenario: Preserve unmodeled source baseline
- **WHEN** source baseline is outside the modeled range
- **THEN** no-op and unrelated scalar assignment/removal preserve it, while replacement through defaultText.baseline rejects without output
