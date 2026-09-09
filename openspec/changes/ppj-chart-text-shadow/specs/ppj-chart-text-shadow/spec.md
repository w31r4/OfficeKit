## Purpose

Express direct outer shadows on chart text while preserving native attribute presence and source-bound edit boundaries.

## ADDED Requirements

### Requirement: Direct text shadow field
PPJ chart text styles SHALL accept `shadow` with required RGB/theme/grammar color and optional blur (0–1000 pt), distance (0–100000 pt), angle (-360–360 degrees), opacity (0–1 or opacity token), alignment and rotateWithShape. Theme colors SHALL retain theme identity; unsupported theme transforms SHALL be rejected. Grammar colors SHALL resolve their supported transforms. Omitted geometry/opacity/rotation attributes SHALL remain omitted; explicit zero/false SHALL survive. Angles SHALL normalize to native positive turns and geometry SHALL use native precision.

#### Scenario: Shared typography and vector defaults
- **WHEN** a shadow is authored on titles, legends, axes, data labels, trendline label styles or rich paragraph/run/end styles
- **THEN** native character effects and fresh projection SHALL retain it, and vector text SHALL carry it with explicit run shadow taking precedence over title defaults.

#### Scenario: Presence and tokens
- **WHEN** a color-only shadow is changed to explicit zero geometry, zero opacity and false rotation and then restored to color-only
- **THEN** each native optional attribute SHALL reflect exactly that presence, including after fresh projection, and color/opacity tokens SHALL resolve consistently in authored and source-bound modes.

### Requirement: Source-bound shadow lifecycle
Recognized chart text shadows SHALL support create, modify, delete and recreate operations. An unchanged projected program SHALL retain original bytes. Editing a shadow SHALL preserve literal text and every non-target ZIP entry.

#### Scenario: Ordinary and combination charts
- **WHEN** a line or categorical combination chart is projected and its text shadows are changed, removed and recreated
- **THEN** only its target ChartPart SHALL change and each fresh native projection SHALL recover the new state.

### Requirement: Unknown effects retain source ownership
Only a single direct outer-shadow effect with supported color/alpha and attributes SHALL be editable. Duplicate effects, empty/mixed effect lists, unknown attributes/children, unsupported transforms and invalid values SHALL retain source ownership without enabling chart edits.

#### Scenario: Unsupported native effect
- **WHEN** a native trendline rich run contains an unsupported effect graph
- **THEN** no-op SHALL preserve its original bytes and an analytics edit SHALL fail without output.
