## Purpose

Provide independently authorized paragraph default gradients with native presence, reversible solid-color transitions and preservation of surrounding source content.

## ADDED Requirements

### Requirement: Paragraph default gradient lifecycle

Ordinary text/shape paragraphs SHALL support assignment, deletion, gradient-only defaultText/style wrapper removal and restoration of defaultText.gradient under its exact setTextParagraphStyle field authority. The existing linear and centered radial profiles, ordered duplicate stop positions, RGB/color-token stops and explicit stop opacity zero/one SHALL retain their meanings and native precision. Linear angles SHALL normalize after native rounding. Declared grammar stop colors SHALL take precedence in source-bound edits.

#### Scenario: Change remove and restore a gradient
- **WHEN** a valid linear or radial gradient is assigned, deleted directly or through its gradient-only wrapper, and restored
- **THEN** native output and fresh PPJ retain the requested gradient or absence while original text, direct runs, neighboring paragraph paint and non-target XML/ZIP remain unchanged

#### Scenario: Switch between solid and gradient paint
- **WHEN** color is removed and gradient assigned, or gradient removed and color assigned
- **THEN** the mutually exclusive paint changes only when both fields have authority, and missing either authority rejects without output

#### Scenario: Preserve source-owned gradient state
- **WHEN** a source gradient has unsupported scaling, theme stops, duplicate fill nodes or unknown nested content
- **THEN** no-op preserves original bytes and unrelated scalar edits preserve that paint while attempted replacement rejects

#### Scenario: Reject invalid paint
- **WHEN** a gradient is invalid, conflicts with color, lacks authority or accompanies an unsupported field change
- **THEN** compilation rejects without output; replacing a source solid color carrying unmodeled transforms with a gradient also rejects
