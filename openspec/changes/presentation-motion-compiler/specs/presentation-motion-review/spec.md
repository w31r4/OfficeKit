## ADDED Requirements

### Requirement: Motion inspection
Presentation inspection SHALL expose individual animation records and Morph pair
records with target identity, effect, phase, start, duration, build, order,
capability, and source revision evidence.

#### Scenario: Inspect a generated deck
- **WHEN** an Agent requests animation inspection after reimport
- **THEN** every emitted animation and Morph pair is discoverable without raw XML
  selectors.

### Requirement: Motion-aware review
`reviewArtifact()` SHALL return a motion section containing plan hash, animation
count, motion-unit count, Morph pair count, playback evidence, and issues. It
SHALL block invalid targets, incompatible builds, over-limit timing, broken Morph
pairs, plan mismatches, and reader-policy violations.

#### Scenario: Review structural evidence
- **WHEN** a generated deck has valid canonical timing and `playbackEvidence:
  "structural"`
- **THEN** review passes the structural motion checks but explicitly does not
  claim Keynote or PowerPoint playback.

#### Scenario: Report excessive choreography
- **WHEN** a live page exceeds the bounded motion budget or repeats pulse effects
- **THEN** review returns a warning with the affected slide and does not rewrite
  the deck automatically.

### Requirement: Composition warnings
Review SHALL warn when a new structured composition plan has no matching visual
carrier, an unintentionally tiny occupied canvas, repeated card-wall silhouettes,
or motion attached to an unfinished composition. Warnings SHALL remain distinct
from deterministic blocking failures.

#### Scenario: Detect an under-composed page
- **WHEN** a non-sparse page declares a chart carrier but contains only a title and
  a small text box
- **THEN** review reports an under-composition warning with page identity.

### Requirement: Focused real-world evidence
The change SHALL provide one financial data deck, one causal architecture deck,
and one brand Morph deck using the public package. Each artifact SHALL record
static review, structural motion evidence, and any observed Keynote playback;
PowerPoint evidence SHALL remain explicitly unverified on macOS.

#### Scenario: Reopen a dogfood artifact
- **WHEN** a real scenario is exported and reopened through the public package
- **THEN** content, motion records, task plan, and review evidence remain available
  for continuation.
