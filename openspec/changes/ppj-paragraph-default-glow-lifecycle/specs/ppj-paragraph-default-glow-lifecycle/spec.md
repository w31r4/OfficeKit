## Purpose

Provide a complete direct paragraph default glow lifecycle while preserving sibling effect state, source color identity and native presence.

## ADDED Requirements

### Requirement: Paragraph default glow lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.glow assignment, deletion, glow-only defaultText/style wrapper removal and restoration under exact setTextParagraphStyle field authority. Radius SHALL accept finite 0..1000pt and round to native EMU precision. Explicit opacity zero/one and omission SHALL remain distinct; RGB/RGBA, color tokens, opacity tokens and untransformed standard theme identity SHALL retain their meanings. Declared grammar colors SHALL take precedence during source-bound edits; tint/shade SHALL resolve to RGB.

#### Scenario: Assign delete and restore glow
- **WHEN** a valid glow is assigned, deleted directly or through its glow-only wrapper, and restored
- **THEN** native output and fresh PPJ retain the requested presence, radius and color/opacity while direct runs, original text, other paragraphs and non-target XML/ZIP remain unchanged

#### Scenario: Retain mixed effect siblings
- **WHEN** a direct glow is edited or removed beside other effects in one flat list
- **THEN** existing modeled sibling values remain stable and sibling XML/attributes are preserved; only the glow and an empty attribute-free list wrapper can be removed

#### Scenario: Retain unsupported source owners
- **WHEN** the source has duplicate glow/list nodes, an effect DAG or an unmodeled glow color/radius/descendant
- **THEN** no-op preserves source bytes, unrelated scalar edits preserve the source state, and attempted glow replacement rejects without output

#### Scenario: Reject invalid requests
- **WHEN** glow input is invalid, uses a wrong-kind token, lacks exact field authority or accompanies unsupported default-field edits
- **THEN** compilation rejects without output
