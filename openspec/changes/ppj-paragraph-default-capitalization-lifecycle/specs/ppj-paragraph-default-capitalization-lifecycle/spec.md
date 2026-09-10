## Purpose

Provide independent direct paragraph capitalization editing in source-bound PPJ with explicit cancellation, optional presence and preservation of original text and source content.

## ADDED Requirements

### Requirement: Direct paragraph default capitalization lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.capitalization values none, small and all, deletion, capitalization-only defaultText/style wrapper removal and restoration under exact setTextParagraphStyle field authority. Explicit none SHALL remain distinct from omission.

#### Scenario: Assign cancel remove and restore capitalization
- **WHEN** capitalization is assigned, explicitly canceled with none, removed directly or through its capitalization-only wrapper, and restored
- **THEN** native cap and fresh PPJ presence follow the request while original text, direct run capitalization, other defaults, unknown attributes and non-target XML/ZIP content remain unchanged

#### Scenario: Reject invalid or unauthorized changes
- **WHEN** capitalization is outside the enum, lacks exact field authority or accompanies unsupported default-style changes
- **THEN** compilation rejects without output

#### Scenario: Retain invalid native capitalization
- **WHEN** source cap uses an unmodeled token
- **THEN** no-op preserves original bytes, unrelated scalar assignment/removal retains the invalid source value without introducing validation warnings, and replacement rejects without output
