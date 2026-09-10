## Purpose

Expose the supported automatic-list numbering formats in PPJ and allow source-bound paragraphs to change format while retaining their existing start value and source content.

## ADDED Requirements

### Requirement: Closed numbering format vocabulary

PPJ SHALL enumerate the 41 schemes already supported by the native codec. A numbered bullet SHALL contain exactly one of `scheme` or `format`. The aliases decimal, lower-alpha, upper-alpha, lower-roman and upper-roman SHALL resolve respectively to arabicPeriod, alphaLcPeriod, alphaUcPeriod, romanLcPeriod and romanUcPeriod. Fresh native projection SHALL emit canonical `scheme` syntax.

#### Scenario: Supported catalog and aliases

- **WHEN** PPJ authors a supported scheme or one of the five format aliases
- **THEN** native export and fresh projection SHALL retain the corresponding canonical scheme and startAt presence

#### Scenario: Missing or invalid format selection

- **WHEN** a numbered bullet omits both selections, supplies both, or uses an unsupported value
- **THEN** validation SHALL reject it without emitting a candidate

### Requirement: Source-bound numbering format edits

Existing numbered text/shape paragraphs SHALL support changing and restoring their format. Each changed paragraph SHALL require the exact `bullet.scheme` or `bullet.format` capability matching the requested spelling. An accompanying startAt edit SHALL independently require startAt authority. Edits SHALL preserve non-target marker attributes, start value spelling and presence, styling, runs, neighbors and package content.

#### Scenario: Change scheme and restore through fresh projection

- **WHEN** an existing numbered paragraph changes scheme or uses a format alias, then restores its original scheme
- **THEN** only the requested native attributes SHALL change and fresh projection SHALL retain each canonical scheme

#### Scenario: Unauthorized or unsupported marker replacement

- **WHEN** the requested format capability is absent, or a format-only request would create a numbered marker from another marker kind or an unmodeled source marker
- **THEN** the request SHALL fail closed while no-op and unrelated edits preserve the original source

### Requirement: Honest preview and discoverability

Schema, Help and field guidance SHALL agree on the vocabulary and alias mapping. Preview SHALL retain a partial or unavailable automatic-number diagnostic wherever it cannot paint the requested format.

#### Scenario: Alternate numbered format in preview

- **WHEN** a preview contains a supported numbering scheme or format alias
- **THEN** unsupported list painting SHALL be diagnosed without claiming complete numbering layout
