## ADDED Requirements

### Requirement: Source-bound page name editing
Every hash-bound imported page SHALL expose a capability to set or clear its
direct PPJ `name` without changing page topology.

#### Scenario: Rename one imported page
- **WHEN** an Agent changes `pages[].name` through an issued `setName` capability
- **THEN** build changes only that SlidePart and reimport recovers the exact name

### Requirement: Source-bound slide-show visibility editing
An imported page with canonical native visibility SHALL expose `setHidden` and
an explicit boolean PPJ state.

#### Scenario: Hide one visible page
- **WHEN** an Agent changes a capable page from `hidden: false` to `true`
- **THEN** build writes the canonical skipped-slide state without changing
  content, layout, notes, comments, transition, or relationships

### Requirement: Opaque visibility fails closed
An imported page with malformed or irregular native visibility SHALL preserve
that source state and SHALL NOT advertise `setHidden`.

#### Scenario: Invalid native show value
- **WHEN** a slide root contains an unsupported `show` lexical value
- **THEN** PPJ omits editable hidden state and a requested mutation is rejected

### Requirement: Agent discoverability
The generated PPJ reference and delivery guidance SHALL distinguish page
visibility from element visibility and from custom-show membership.

#### Scenario: Agent prepares a hidden appendix
- **WHEN** an Agent reads PPJ page guidance
- **THEN** it can find `pages[].hidden`, its source-bound boundary, and the fact
  that it does not alter sections or custom shows
