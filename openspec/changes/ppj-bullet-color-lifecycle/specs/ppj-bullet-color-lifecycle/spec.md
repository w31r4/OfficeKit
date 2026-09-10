## Purpose

Represent direct list-marker color without losing theme identity or native alpha precision, and edit the choice while preserving the surrounding source document.

## ADDED Requirements

### Requirement: Complete direct color choice
Character, numbered and picture bullets SHALL accept either `color` or `colorFollowText: true`, or neither. Color SHALL support hex RGB/RGBA, token color and `{ rgb: "#RRGGBB", alpha?: number }`. Alpha SHALL be in 0–1, quantized to native 0.00001 units; explicit zero/one and absence SHALL remain distinct. Direct standard theme tokens without tint/shade SHALL retain theme identity. Declared grammar colors SHALL take precedence and use bounded tint/shade resolution. Invalid input and conflicting choices SHALL reject.

#### Scenario: Precise native alpha
- **WHEN** RGB alpha 0.12345 is authored or imported and projected without embedded PPJ
- **THEN** PPJ retains the exact native alpha through its RGB object form and re-authoring retains that alpha

#### Scenario: Theme and follow-text identity
- **WHEN** a direct theme color or explicit follow-text color is authored or imported
- **THEN** PPJ recovers the native choice, distinguishing it from RGB and absence

### Requirement: Independent source lifecycle
Ordinary source-bound text/shape markers SHALL support color addition, assignment, switching, removal and restoration. Every changed color property SHALL require its exact field authority, including both sides of a switch. Only the target color declaration SHALL change; marker, font/size, runs, neighbors and non-target XML/ZIP SHALL remain. No-op SHALL retain source bytes.

#### Scenario: Owner-local edit and deletion
- **WHEN** an authorized color choice is changed or removed using source bytes and their fresh projection
- **THEN** only that paragraph color declaration changes and native re-projection recovers the requested choice or absence

#### Scenario: Missing authority
- **WHEN** a changed color property lacks exact authority
- **THEN** compilation rejects without a candidate

### Requirement: Preserve unmodeled source colors
Malformed, transformed, duplicate or unknown native color declarations SHALL remain source-owned; no-op and unrelated edits SHALL preserve them, and replacement SHALL reject. Parseable invalid native color tokens and alpha values SHALL not abort projection.

#### Scenario: Unknown native color
- **WHEN** a marker has an invalid theme token, invalid alpha, extra transform or duplicate color declaration
- **THEN** its PPJ color choice is absent, unrelated edits preserve the original XML and requested color replacement rejects

### Requirement: Discoverable preview boundary
Schema, Help, registry, generated references and text guidance SHALL describe the same color choice and precision. Preview SHALL preserve direct RGB color/alpha and explicitly diagnose unresolved theme or follow-text marker color.

#### Scenario: Unresolved inherited color
- **WHEN** marker color comes from a theme token or the paragraph text
- **THEN** preview reports the color/layout limitation instead of claiming complete support
