## Purpose

Provide independently authorized paragraph default underline edits in source-bound PPJ while retaining optional presence, original text and native decoration content.

## ADDED Requirements

### Requirement: Direct paragraph default underline lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.underline assignment, explicit cancellation, deletion, underline-only defaultText/style wrapper removal and restoration under exact setTextParagraphStyle field authority. All existing underline tokens SHALL remain accepted: none, words, sng, dbl, heavy, dotted, dottedHeavy, dash, dashHeavy, dashLong, dashLongHeavy, dotDash, dotDashHeavy, dotDotDash, dotDotDashHeavy, wavy, wavyHeavy and wavyDbl; single/double and sng/dbl aliases SHALL use native sng/dbl and project as single/double.

#### Scenario: Assign cancel remove and restore underline
- **WHEN** an accepted underline token or alias is assigned, canceled with none, removed directly or through its underline-only wrapper, and restored
- **THEN** native underline and fresh PPJ retain canonical values and optional presence, with none distinct from omission and original text, direct run styles, other defaults, unknown attributes and non-target XML/ZIP unchanged

#### Scenario: Reject invalid or unauthorized changes
- **WHEN** underline has an invalid type/token, lacks exact field authority or accompanies unsupported default-style changes
- **THEN** compilation rejects without output

#### Scenario: Preserve native decoration content
- **WHEN** source underline has an unknown token or underline fill/line effect children with or without a u attribute
- **THEN** no-op preserves original bytes, unrelated scalar assignment/removal preserves that source content, and replacing underline rejects without output
