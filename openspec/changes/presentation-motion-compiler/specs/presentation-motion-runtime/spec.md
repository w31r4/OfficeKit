## ADDED Requirements

### Requirement: Typed animation surface
The Presentation model SHALL expose `slide.animations.add`, `remove`, `clear`,
and `items` for fade, wipe, fly, zoom, pulse, ordered starts, text builds, and
chart builds. It SHALL accept `animateChartBackground` and SHALL reject
unsupported fields or incompatible target/build combinations.

#### Scenario: Round-trip a chart build
- **WHEN** a source-free bar, line, or pie chart receives a typed chart build
- **THEN** the canonical timing and build records export and reimport with the
  same target, build level, order, and background flag.

#### Scenario: Reject an incompatible build
- **WHEN** a text build targets a chart or a chart build targets a shape
- **THEN** the operation fails before export with a typed capability error.

### Requirement: Bounded canonical timing
The compiler SHALL cap each slide at 32 semantic animations and 64 expanded
timing nodes. It SHALL lower delay and stagger into deterministic timing order,
or reject a request that cannot be represented safely.

#### Scenario: Enforce the semantic limit
- **WHEN** a slide receives a 33rd semantic animation
- **THEN** the add operation fails and prior animations remain unchanged.

#### Scenario: Preserve deterministic order
- **WHEN** two animations use `withPrevious` and a chart build supplies a stagger
- **THEN** export/reimport returns the same sequence and effective timing values.

### Requirement: Cross-slide Morph
The Presentation model SHALL support destination-side `setMorph({ from, pairs })`
for adjacent compatible non-chart objects and SHALL reject stale, non-adjacent,
cross-deck, duplicate-key, name-conflicting, or transition-conflicting pairs.

#### Scenario: Pair source and destination objects
- **WHEN** adjacent slides contain one compatible pair and the destination sets
  Morph
- **THEN** the codec writes matching `!!key` names on both slides and reimport
  returns the pair and source slide identity.

#### Scenario: Reject an unsafe pair
- **WHEN** a pair references a chart, a non-adjacent slide, or a duplicate key
- **THEN** the operation fails closed and does not emit a Morph extension.

### Requirement: Opaque imported timing
Unknown or non-canonical imported timing and Morph graphs SHALL remain byte
preserved. They SHALL expose non-editable capability and SHALL not be silently
replaced by a canonical graph.

#### Scenario: Inspect an opaque graph
- **WHEN** an imported slide contains an unsupported timing graph
- **THEN** inspect reports it as opaque and adding/removing timing fails without
  changing the source package.
