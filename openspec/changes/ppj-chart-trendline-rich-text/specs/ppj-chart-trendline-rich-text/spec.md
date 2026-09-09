## Purpose

Allow agents to express native trendline labels with ordered paragraphs, styled runs and line breaks while preserving unsupported imported text graphs.

## ADDED Requirements

### Requirement: Structured trendline label text

`label.text` SHALL accept string/token or `{paragraphs:[{style?,runs:[{text,style?}|{break:true,style?}],endStyle?}]}`. Styles SHALL use existing chart text styles, with alignment permitted only on paragraph style. Text runs SHALL support string grammar tokens, whitespace and empty text; empty paragraphs SHALL be supported. A single unstyled nonempty literal run of at most 255 characters SHALL canonicalize to the existing string form. Resource limits SHALL be 4096 paragraphs, 16384 runs/breaks and 1,048,576 UTF-16 characters; control characters SHALL require typed breaks rather than literal controls.

#### Scenario: Styled text and breaks round trip
- **WHEN** an ordinary or categorical combo trendline label contains multiple styled runs, a break and another paragraph
- **THEN** native text and fresh PPJ projection SHALL preserve content order, whitespace, paragraph defaults, run/end styles and explicit false overrides

### Requirement: Full text-state edit lifecycle

Supported labels SHALL permit structured text replacement, editing, conversion to literal text, removal for automatic content and recreation. No-op SHALL preserve original bytes. Text-only edits SHALL preserve other label/chart state and every non-target ZIP entry.

#### Scenario: Transition between structured and automatic labels
- **WHEN** a source-bound rich label is edited, replaced with literal text, removed and recreated
- **THEN** only its ChartPart SHALL change and fresh projection SHALL match canonical requested text state

### Requirement: Unsupported graphs remain opaque

Formula references, fields, hyperlinks, unsupported body/list/character properties, malformed order/attributes and invalid typed runs SHALL fail closed without flattening. Invalid PPJ or wire state SHALL produce no edited file.

#### Scenario: Unknown formatting cannot disappear
- **WHEN** imported rich text carries an unknown property or relationship-bearing child
- **THEN** no-op SHALL preserve input bytes and analytics replacement SHALL be rejected without output
