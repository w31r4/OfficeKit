## ADDED Requirements

### Requirement: Strategy-bound review evidence
`reviewArtifact({ authoringPlan })` SHALL extend its existing Presentation
design report with plan-bound communication, narrative, cognitive, and visual
risk evidence while reusing structural, layout, motion, and delivery evidence.

#### Scenario: Review a current strategy plan
- **WHEN** a candidate is reviewed with its current authoring plan
- **THEN** the report identifies the communication job, scenario, direction, delivery mode, plan revision, and issues for each deterministic quality layer

### Requirement: Deterministic errors remain narrow
Review SHALL block only verifiable contract violations such as plan/candidate
mismatch, invalid strategy values, required unresolved decisions, content or
font limits, strict design authority violations, and invalid native behavior.

#### Scenario: Valid but unconventional design
- **WHEN** a deck intentionally uses an unconventional composition without violating a declared invariant
- **THEN** deterministic review does not fail merely because the composition differs from a style preference

### Requirement: Visual heuristics stay advisory
Review MAY warn about repeated dominant geometry, large hollow containers,
repeated card surfaces, weak text-container hierarchy, missing structural
anchors, decorative competition, and repeated page silhouettes, but SHALL not
describe those signals as proof of bad aesthetics.

#### Scenario: Repeated circular motif
- **WHEN** the modeled deck repeatedly uses one dominant circular motif across multiple pages
- **THEN** review identifies affected pages and the repeated pattern as a warning for visual inspection

### Requirement: Facts and aesthetics remain honest boundaries
The runtime SHALL NOT claim to verify factual truth or aesthetic quality from
plan structure, object counts, text extraction, or warning absence.

#### Scenario: Warning-free report
- **WHEN** a report contains no deterministic or heuristic issue
- **THEN** it still records source verification and visual review as separate Agent or human evidence rather than declaring the presentation objectively correct or beautiful
