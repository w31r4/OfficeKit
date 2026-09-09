## Purpose

Express direct signed chart character baseline offsets in PPJ and preserve native presence through authored and source-bound editing.

## ADDED Requirements

### Requirement: Baseline units and precision

ChartTextStyle SHALL accept optional baseline as a finite percentage from -400 to 400, matching ordinary PPJ text. Native state SHALL use integer thousandths of a percent; finer inputs SHALL round to the nearest native unit with ties to even. Fresh projection SHALL return that native value as a percentage. Explicit zero SHALL remain distinct from omission, and a baseline-only style SHALL be meaningful. The global-font-family-only profile SHALL remain restricted.

#### Scenario: Fractional baseline and explicit reset
- **WHEN** a style requests baseline -25.125, then zero, then omits the field
- **THEN** native baseline SHALL be -25125, then explicitly 0, then absent, with corresponding fresh PPJ projection

### Requirement: Baseline edit lifecycle and propagation

Existing chart text-style consumers SHALL support baseline authoring, change, reset, removal and recreation, including title/legend/axis/data-label/trendline and rich paragraph/run/end styles. Declared field precedence SHALL apply. Vector title defaults SHALL yield to explicit run baseline, and vector label styles SHALL retain baseline. No-op SHALL preserve original bytes, and native chart style edits SHALL preserve non-target ZIP entries.

#### Scenario: Edit a native line or combo chart
- **WHEN** baseline is edited against a fresh native source projection
- **THEN** only the target ChartPart SHALL change and a second projection SHALL reflect the requested native-precision state

### Requirement: Invalid and unknown state remains fail closed

Out-of-range, non-finite or malformed authored/wire/native baseline values SHALL fail closed. Recognizing baseline SHALL not authorize unknown native character effects. Unrelated JS chart edits SHALL retain signed and explicit-zero baseline wire values.

#### Scenario: Unsupported imported character state
- **WHEN** a native chart style contains an invalid baseline or unknown character property
- **THEN** unsupported edits SHALL remain unavailable and no-op SHALL retain the source content
