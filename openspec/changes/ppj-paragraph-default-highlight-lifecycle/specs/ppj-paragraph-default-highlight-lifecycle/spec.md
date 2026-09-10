## Purpose

Provide independently authorized paragraph default highlight edits with explicit presence, RGB/theme color fidelity and preservation of original text and source content.

## ADDED Requirements

### Requirement: Direct paragraph default highlight lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.highlight assignment, deletion, highlight-only defaultText/style wrapper removal and restoration under exact setTextParagraphStyle field authority. Opaque RGB and declared color-token/tint/shade inputs SHALL resolve through existing color rules. Simple native theme highlights SHALL project as token objects; changed untransformed standard theme tokens SHALL retain scheme identity unless shadowed by declared grammar tokens.

#### Scenario: Assign remove and restore highlight
- **WHEN** an accepted RGB, color token or simple theme highlight is assigned, deleted or restored
- **THEN** native/fresh PPJ presence and color follow the request, and original text, direct run styles, other defaults and non-target XML/ZIP remain unchanged

#### Scenario: Edit one paragraph
- **WHEN** a paragraph highlight is changed while another paragraph uses a theme highlight
- **THEN** the unchanged paragraph retains its native scheme binding

#### Scenario: Reject invalid or unauthorized changes
- **WHEN** highlight is invalid, resolves nonopaque, lacks field authority or accompanies unsupported default edits
- **THEN** compilation rejects without output

#### Scenario: Preserve unmodeled native highlight
- **WHEN** source highlight contains color transforms, duplicate nodes or unknown nested elements
- **THEN** no-op preserves original bytes, unrelated scalar assignment/removal preserves the source graph and replacement rejects without output


#### Scenario: Reject source content the SDK cannot retain
- **WHEN** a source color node contains illegal character data that the native SDK loses during editing
- **THEN** no-op preserves original bytes and attempted edits reject without output under source binding validation
