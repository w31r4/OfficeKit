## Purpose

Provide independently authorized paragraph default strike edits in source-bound PPJ with explicit cancellation, optional presence and preservation of original text and native content.

## ADDED Requirements

### Requirement: Direct paragraph default strike lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.strike assignment, explicit cancellation, deletion, strike-only defaultText/style wrapper removal and restoration under exact setTextParagraphStyle field authority. True/false SHALL map to sngStrike/noStrike, and noStrike/sngStrike/dblStrike strings SHALL be accepted. Fresh projection SHALL use canonical strings; noStrike SHALL remain distinct from omission.

#### Scenario: Assign cancel remove and restore strike
- **WHEN** boolean or string strike is assigned, explicitly canceled, removed directly or through its strike-only wrapper, and restored
- **THEN** native strike and fresh PPJ presence follow the canonical request while original text, direct run strike, other defaults, unknown attributes and non-target XML/ZIP content remain unchanged

#### Scenario: Reject invalid or unauthorized changes
- **WHEN** strike is outside the accepted boolean/enum values, lacks exact field authority or accompanies unsupported default-style changes
- **THEN** compilation rejects without output

#### Scenario: Preserve unmodeled native strike
- **WHEN** source strike uses an unmodeled token
- **THEN** no-op preserves original bytes, unrelated scalar assignment/removal preserves that source value, and replacement rejects without output
