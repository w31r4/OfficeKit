## Purpose

Express direct chart capitalization display state while preserving original text, native presence and the source-bound edit lifecycle.

## ADDED Requirements

### Requirement: Capitalization state preserves text and presence

Existing chartTextStyle consumers SHALL accept optional capitalization as none, small or all, using the ordinary PPJ text vocabulary. The field SHALL preserve native cap and explicit none independently of omission. A capitalization-only style SHALL be meaningful. Applying any mode SHALL preserve original literal characters. The global-font-family-only profile SHALL remain restricted.

#### Scenario: Reset then remove capitalization
- **WHEN** a style changes from all to none and then omits capitalization
- **THEN** native cap SHALL change from all to explicit none and then absence, while literal text remains identical

### Requirement: Capitalization lifecycle and propagation

Authored and source-bound chart styles SHALL support creation/change/reset/removal/recreation across title, legend, axis, data-label and trendline owners, including rich paragraph/run/end styles. Declared field precedence SHALL apply. Vector text SHALL retain the field and explicit title-run capitalization SHALL override chart defaults. No-op SHALL preserve original bytes and native chart style edits SHALL preserve non-target ZIP entries.

#### Scenario: Edit and reproject line and combo charts
- **WHEN** capitalization is edited against a fresh native projection
- **THEN** only the target ChartPart SHALL change and fresh PPJ projection SHALL reflect the requested field or absence

### Requirement: Invalid and unknown state remains fail closed

Empty, invalid or unknown authored/wire/native capitalization values SHALL fail closed. Recognizing cap SHALL not authorize unknown imported character properties. Unrelated JS chart edits SHALL retain optional capitalization state.

#### Scenario: Unknown native character state
- **WHEN** a native chart owner has malformed cap or another unsupported character property
- **THEN** unsupported style edits SHALL remain unavailable and no-op SHALL preserve source content
