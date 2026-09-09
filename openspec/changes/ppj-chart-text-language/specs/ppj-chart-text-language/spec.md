## Purpose

Let PPJ preserve native chart text language tags across authoring, source-bound editing and projection using the established language-tag contract.

## ADDED Requirements

### Requirement: Chart language field

Every existing chartTextStyle owner SHALL accept optional language as a bounded language tag or string grammar token. The field SHALL use the established 2–63-character language-tag lexical profile, preserve explicit casing/value and native attribute presence, and count as meaningful style by itself. The native chart font-family-only profile SHALL remain restricted to font family.

#### Scenario: Language-only text is editable
- **WHEN** a supported chart title, legend, axis, data label or trendline style declares language, including an explicit en-US
- **THEN** native character properties and fresh PPJ projection SHALL retain that language without requiring another style property

### Requirement: Language lifecycle and propagation

Authored and source-bound chart styles SHALL support language addition, replacement, omission/deletion and recreation. Rich paragraph/run/end styles and vector chart title defaults SHALL retain language; more-specific vector run language SHALL take precedence. Existing per-field grammar precedence SHALL also apply to language. No-op SHALL preserve original bytes and language-only native chart edits SHALL preserve non-target ZIP entries.

#### Scenario: Change then remove language
- **WHEN** supported chart text language is changed, removed and restored against fresh source projections
- **THEN** only the target ChartPart SHALL change and native/projection state SHALL match the requested language state

### Requirement: Invalid and unknown language state remains safe

Invalid literal/token/wire/native language values SHALL fail closed. Unsupported imported properties SHALL not become editable merely because lang is now recognized. The implementation SHALL not infer dictionary availability or host proofreading success from a valid language tag.

#### Scenario: Reject an invalid language tag
- **WHEN** an authored or edited language is empty, malformed or outside the lexical budget
- **THEN** compilation SHALL reject it without an output file, and an imported unsupported owner SHALL retain input bytes on no-op
