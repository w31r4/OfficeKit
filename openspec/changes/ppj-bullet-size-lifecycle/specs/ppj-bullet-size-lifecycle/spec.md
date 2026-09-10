## Purpose

Represent the direct size choice of a list marker and edit it independently while preserving surrounding text and source document content.

## ADDED Requirements

### Requirement: Explicit bounded size choice
Character, numbered and picture bullets SHALL accept at most one of `size`, `sizePercent` and `sizeFollowText: true`. Size SHALL be 1–768 points, quantized to 0.01 points. SizePercent SHALL be a ratio of 0.25–4, quantized to 0.00001; one means 100%. Absence, explicit ratio one and follow-text SHALL remain distinct. Invalid types/ranges, false follow-text and conflicting choices SHALL reject.

#### Scenario: Native choice and precision
- **WHEN** a point size, relative ratio or explicit follow-text choice is authored and projected without embedded PPJ
- **THEN** PPJ retains the selected choice and native precision, including its distinction from absence

### Requirement: Independent source size lifecycle
Ordinary source-bound text/shape markers SHALL support size addition, assignment, switching, removal and restoration. Each changed property SHALL require its exact field authority, including both sides of a switch. Only the target size declaration SHALL change; marker, font/color, runs, neighboring paragraphs and non-target XML/ZIP SHALL remain. No-op SHALL retain source bytes.

#### Scenario: Edit switch delete restore
- **WHEN** an authorized size edit is compiled from source bytes and their fresh projection
- **THEN** native re-projection recovers the requested choice or absence, and unrelated source content remains

#### Scenario: Insufficient authority
- **WHEN** a changed size property lacks exact authority
- **THEN** compilation rejects without a candidate

### Requirement: Preserve unmodeled native size
Invalid numeric values, duplicate choices and unknown size metadata SHALL remain source-owned. No-op and unrelated edits SHALL preserve their XML; replacement SHALL reject. Parseable invalid numeric tokens SHALL not abort projection.

#### Scenario: Malformed imported size
- **WHEN** a source marker contains an invalid number, out-of-range value or ambiguous size choice
- **THEN** PPJ omits the unmodeled size choice and preserves the original source through unrelated edits

### Requirement: Discoverable preview boundary
Schema, Help, registry, generated references and text guidance SHALL agree on size units, bounds, choices and lifecycle. Preview SHALL retain existing direct point-size painting and explicitly diagnose unresolved relative or follow-text marker sizing.

#### Scenario: Relative-size preview
- **WHEN** a marker uses a percentage or follow-text size
- **THEN** preview reports the unresolved layout limitation instead of claiming complete rendering
