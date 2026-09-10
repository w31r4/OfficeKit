## Purpose

Represent a paragraph's direct default tab size in points while preserving
native coordinate precision and unrelated source content during editing.

## ADDED Requirements

### Requirement: Numeric coordinate and presence

PPJ SHALL accept `text.paragraphs[].style.defaultTabSize` in points over the full
signed 32-bit native EMU range, including negative values and zero. Conversion
SHALL use nearest EMU with ties-to-even. Explicit zero and absence SHALL remain
distinct under authored style precedence and snapshot-free re-import. Absence
SHALL NOT materialize a default. Non-numbers and out-of-range values SHALL reject.

#### Scenario: Signed endpoints and sub-EMU precision
- **WHEN** an authored paragraph uses a signed endpoint or fractional-point default tab size
- **THEN** export/re-import preserves the quantized EMU, both endpoints remain valid PPJ, and subsequent no-op retains source bytes

#### Scenario: Zero overrides lower-priority size
- **WHEN** a higher-priority paragraph style sets zero and a lower-priority style sets a nonzero size
- **THEN** zero is retained; removing the higher declaration selects the lower value

### Requirement: Source lifecycle and independent tab state

Ordinary recognized text/shape paragraphs SHALL support add/set/remove/restore
under exact `text.paragraphs[].style.defaultTabSize` authority. Removing a style
containing only this field SHALL remove the direct native attribute. Explicit
tabStops/noTabStops, indentation, direction, literal tabs, runs, neighboring
paragraphs and non-target XML/ZIP SHALL remain unchanged by the field edit.

#### Scenario: Original-source removal and restoration
- **WHEN** a fresh projection of original source removes defaultTabSize and a later request restores it
- **THEN** native presence and reprojected values reflect both edits while only the target attribute changes

#### Scenario: Exact authority missing
- **WHEN** defaultTabSize changes without its exact source capability
- **THEN** compilation rejects without an output artifact

### Requirement: Unknown native state stays source-owned

Recognized native integer spellings SHALL be preserved when unchanged. Unknown,
noninteger or overflowing native tokens SHALL be omitted from PPJ state, retained
by no-op and unrelated edits, and reject replacement with a modeled value.

#### Scenario: Unknown token and independent default-bold edit
- **WHEN** a source contains an unrecognized default-tab-size token
- **THEN** no-op and an unrelated default-bold edit retain the token, while assigning defaultTabSize rejects

### Requirement: Discovery and honest preview

Schema, Help, registry and Agent guidance SHALL describe the same units, range,
presence and source lifecycle. Preview SHALL explicitly diagnose zero and nonzero
defaultTabSize as partial/unmapped until measured tab placement is implemented.

#### Scenario: Explicit zero in preview
- **WHEN** preview assesses defaultTabSize zero or a nonzero value
- **THEN** it retains an explicit field diagnostic without claiming tab-layout fidelity
