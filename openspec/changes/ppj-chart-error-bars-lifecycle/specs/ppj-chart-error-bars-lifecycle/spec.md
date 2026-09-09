## Purpose

Allow imported presentation chart series to express the full presence lifecycle of the existing PPJ errorBars object while retaining unrelated source content.

## ADDED Requirements

### Requirement: Error-bar object lifecycle
PPJ SHALL support adding, modifying, removing and recreating `series[].errorBars` for editable column, bar and line series, including categorical combo series. The existing fixed-value, percentage, standard-deviation and standard-error modes SHALL retain their value, direction, type, cap and stroke semantics. Omitting the object from an edited projected series SHALL remove an existing supported object.

#### Scenario: Repeated source-bound edits
- **WHEN** an imported supported series loses its errorBars object, receives a new one, changes modes, and is projected again after each edit
- **THEN** native chart XML and the fresh PPJ projection reflect each requested state, including a zero scalar and standard-error without a scalar

### Requirement: Preserve source ownership
Changing errorBars presence SHALL preserve trendlines, other series, chart data and unrelated package entries. Unsupported native owners SHALL remain preserved or cause the edit to fail closed. A scalar insertion SHALL NOT replace custom plus/minus data that is absent from the PPJ projection.

#### Scenario: Unsupported or custom owner
- **WHEN** the native owner is duplicated or malformed, or a scalar insertion targets an unprojected custom owner
- **THEN** the edit is rejected and a no-op import/export preserves the original chart part

### Requirement: Keep existing validation
New errorBars objects SHALL obey the same type and scalar constraints as authored objects, and custom-data presence changes SHALL remain outside this scalar lifecycle capability.

#### Scenario: Invalid inserted object
- **WHEN** an inserted fixed-value object has a negative or missing value, or standard-error carries a scalar value
- **THEN** compilation rejects the request instead of emitting invalid error-bar XML

### Requirement: Honest local preview boundary
The local SVG preview SHALL report error-bar charts as partial while it does not draw their error bars.

#### Scenario: Error bars added to a locally previewed chart
- **WHEN** a supported column, bar, line or categorical combo series contains errorBars
- **THEN** the preview emits an explicit error-bars-not-rendered diagnostic instead of reporting complete chart support
