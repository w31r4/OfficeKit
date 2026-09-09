## Purpose

Express direct chart character strike state in PPJ while preserving native attribute presence and the source-bound edit lifecycle.

## ADDED Requirements

### Requirement: Chart character strike vocabulary and presence

Every existing chartTextStyle consumer SHALL accept optional strike as true, false, noStrike, sngStrike or dblStrike. Boolean true SHALL produce sngStrike and false SHALL produce noStrike. Fresh chart style projection SHALL use the native string tokens and retain explicit noStrike. A strike-only style SHALL be meaningful; removing strike SHALL remove the direct attribute. The special chart-wide font-family profile SHALL stay restricted to font family.

#### Scenario: Cancel and then omit strike
- **WHEN** a supported chart style changes from true to false and then omits strike
- **THEN** its native state SHALL change from sngStrike to explicit noStrike and then to an absent strike attribute

### Requirement: Strike lifecycle and propagation

Authored and source-bound styles SHALL support creation, change, removal and recreation across chart title, legend, axis, data-label and trendline owners, including rich paragraph/run/end styles. Declared nested grammar precedence SHALL apply to strike. Vector text SHALL retain strike and explicit title-run strike SHALL override chart defaults. No-op SHALL preserve original bytes, and source-bound chart style edits SHALL preserve non-target ZIP entries.

#### Scenario: Edit and reproject native text styles
- **WHEN** a recognized line or combo chart's direct strike is edited against a fresh source projection
- **THEN** only the target ChartPart SHALL change and fresh projection SHALL retain the selected token or its absence

### Requirement: Invalid and unsupported strike remains fail closed

Unknown or invalid authored, wire and native strike values SHALL be rejected or remain source-owned. Recognizing strike SHALL not unlock unknown imported character effects. Unrelated JS chart edits SHALL retain the wire strike field.

#### Scenario: Preserve unsupported imported text
- **WHEN** imported chart character properties contain a malformed strike token or unknown character attribute
- **THEN** no unsupported style edit SHALL be authorized and no-op SHALL preserve the source content
