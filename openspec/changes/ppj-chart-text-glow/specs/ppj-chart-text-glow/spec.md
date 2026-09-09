## Purpose

Express direct chart character glow and preserve it together with outer shadows during authoring and source-bound edits.

## ADDED Requirements

### Requirement: Chart glow field
Chart text styles SHALL accept `glow` using required `color` and `radius` (0–1000 points), with optional `opacity` (0–1 or opacity token). RGB and grammar colors SHALL resolve supported transforms; undeclared standard theme tokens SHALL retain native identity and reject unsupported tint/shade. Radius SHALL round to EMU precision with ties to even. Explicit zero radius/opacity and omitted alpha SHALL remain distinct. Glow SHALL be available to titles, legends, axes, data labels and trendline label paragraph/run/end styles.

#### Scenario: Color and opacity tokens
- **WHEN** a chart style uses a grammar color plus opacity token, or a theme color without opacity
- **THEN** authored and source-bound output SHALL preserve the resolved RGB/alpha or theme identity/alpha absence, and fresh projection SHALL reproduce the effective field.

### Requirement: Glow and shadow lifecycle
A direct glow SHALL coexist with one supported outer shadow in native order. Changing or deleting either field SHALL retain the other field and literal text. Removing both SHALL remove the direct effect list. An unchanged projected program SHALL retain original bytes, and glow edits SHALL change only the target ChartPart.

#### Scenario: Independent effects on line and combo charts
- **WHEN** glow is changed, removed, restored, then shadow is removed and restored on line and categorical combination chart text
- **THEN** every fresh projection SHALL recover each requested effect state and every non-target ZIP entry SHALL remain byte-identical.

### Requirement: Precedence and generated text
Glow SHALL follow declared chart text field precedence and propagate to vector text; explicit rich run glow SHALL override a title default independently of its shadow.

#### Scenario: Vector title override
- **WHEN** a vector chart has default glow/shadow and a run specifies a different glow
- **THEN** the run SHALL retain its own glow and the default shadow, while other text retains its appropriate defaults.

### Requirement: Unknown effects stay source-owned
Duplicate/reordered effects, missing or invalid native radius, unknown effect/attribute/child content, unsupported color transforms and malformed alpha SHALL remain outside the recognized profile. This expands the preceding shadow-only profile to include known glow plus shadow; it SHALL NOT authorize arbitrary mixed graphs.

#### Scenario: Unsupported native graph
- **WHEN** imported rich chart text contains an unsupported effect graph
- **THEN** no-op SHALL retain source bytes and an analytics edit SHALL fail without producing a file.
