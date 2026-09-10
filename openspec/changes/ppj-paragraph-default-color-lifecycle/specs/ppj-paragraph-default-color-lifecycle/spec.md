## Purpose

Provide independently authorized direct paragraph text color edits with RGB/theme identity, optional alpha presence and preservation of surrounding native content.

## ADDED Requirements

### Requirement: Direct paragraph default color lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.color assignment, deletion, color-only defaultText/style wrapper removal and restoration under exact setTextParagraphStyle field authority. RGB/RGBA and color-token inputs SHALL follow existing resolution; untransformed standard source-bound theme tokens SHALL retain scheme identity unless a declared grammar color overrides the token. Explicit direct alpha, including zero and one, SHALL remain distinct from omission.

#### Scenario: Assign delete and restore color
- **WHEN** a valid color is assigned, removed directly or through its color-only wrapper, and restored
- **THEN** native and fresh PPJ retain the requested color/presence while original text, direct run paint, other defaults and non-target XML/ZIP remain unchanged

#### Scenario: Preserve adjacent source precision
- **WHEN** one paragraph color changes
- **THEN** other paragraphs retain their exact native theme bindings and alpha values without quantization through PPJ projection

#### Scenario: Reject invalid or incompatible changes
- **WHEN** color is invalid, lacks exact authority, conflicts with gradient or accompanies unsupported default fields
- **THEN** compilation rejects without output

#### Scenario: Preserve source-owned paint
- **WHEN** source paint contains luminance transforms, duplicate fill nodes or unmodeled fill topology
- **THEN** no-op preserves source bytes, unrelated scalar edits preserve the paint, and attempted solid-color replacement rejects; deletion of projected luminance-transformed color also rejects
