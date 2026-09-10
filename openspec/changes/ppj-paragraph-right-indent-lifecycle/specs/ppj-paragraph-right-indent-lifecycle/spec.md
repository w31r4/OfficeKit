## Purpose

Represent the paragraph's physical right inset independently of other layout fields, with precise units and source-preserving addition, editing and removal.

## ADDED Requirements

### Requirement: Bounded right paragraph indent
Typed paragraph style SHALL accept rightIndent as a finite point value from zero through 4032, rounded to the nearest EMU with ties to even. It SHALL address the physical right paragraph margin independently of direction, left indent, hanging indent and text-box insets. Authored style precedence SHALL follow existing paragraph layers. Explicit zero and absence SHALL remain distinct.

#### Scenario: Author and reproject right indent
- **WHEN** a paragraph selects a bounded rightIndent above a lower-priority style value
- **THEN** native output and fresh projection SHALL preserve the selected value at EMU precision, including explicit zero

#### Scenario: Invalid input
- **WHEN** rightIndent has the wrong type or lies outside the supported point range
- **THEN** compilation SHALL reject without a candidate

### Requirement: Independent source lifecycle
Ordinary source-bound text/shape paragraphs with supported inline topology SHALL support rightIndent addition, assignment, removal and restoration under exact text.paragraphs[].style.rightIndent authority. Removing the field or a rightIndent-only style wrapper SHALL clear its modeled direct attribute. Unrelated paragraph state, runs, neighbors and non-target XML/ZIP SHALL remain.

#### Scenario: Remove from actual source and restore
- **WHEN** rightIndent is removed from a fresh projection of original source bytes and restored from a fresh projection of the resulting candidate
- **THEN** native attribute presence and PPJ projection SHALL match each request, with non-target source content preserved

#### Scenario: Missing field authority
- **WHEN** rightIndent changes without its exact issued capability
- **THEN** compilation SHALL reject without a candidate

### Requirement: Retain native spelling and unknown right margins
Equivalent modeled native integer spelling SHALL remain untouched during no-op and unrelated edits. Malformed, negative, overflowing or out-of-range native right margins SHALL stay source-owned and omitted from modeled rightIndent. Replacement SHALL reject while independent modeled edits SHALL remain available.

#### Scenario: Unmodeled source coordinate
- **WHEN** an imported paragraph has an invalid right-margin token
- **THEN** no-op and unrelated edits SHALL preserve it, and assigning rightIndent SHALL reject

#### Scenario: Explicit native removal against unknown state
- **WHEN** an ordinary source shape or grouped shape receives an explicit right-margin removal intent
- **THEN** an absent attribute SHALL permit idempotent clearing, and unknown native right-margin content SHALL reject even if its modeled projection is absent

### Requirement: Discoverability and explicit layout limits
Schema, Help, generated references, capability documentation and Agent guidance SHALL agree on the field's owner, units, precision and lifecycle. Preview SHALL diagnose unresolved right-indent layout as non-supported.

#### Scenario: Preview native right margin
- **WHEN** a preview scene contains an explicit right paragraph margin
- **THEN** machine-readable diagnostics SHALL retain its unresolved layout limitation
