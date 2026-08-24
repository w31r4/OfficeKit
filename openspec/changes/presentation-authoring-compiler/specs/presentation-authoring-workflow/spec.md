## ADDED Requirements

### Requirement: One-prompt authoring compiler workflow
The Presentations Skill SHALL turn a materially clear one-line request into a
complete reviewed working draft without requiring the user to select internal
layouts, tokens, or APIs. It SHALL ask at most three questions only when the
missing answer changes audience, conclusion, evidence, or authoritative design
source.

#### Scenario: Clear request without visual direction
- **WHEN** the user provides a clear presentation goal but no style or template
- **THEN** the Agent writes a self-directed authoring plan, composes a complete deck, reviews it, and returns a working draft without selecting Grid by default

#### Scenario: Material ambiguity
- **WHEN** audience, decision, evidence, or template authority cannot be safely inferred
- **THEN** the Agent asks only the bounded questions needed before writing the plan

### Requirement: Per-deck design grammar
Every broad creation or redesign plan SHALL contain a concrete design grammar
covering palette roles, typography roles, spacing/grid, density rhythm, visual
motif, imagery, charts, and declared invariants. The grammar SHALL be derived
from `self-directed`, `design-system`, `template`, or `style-transfer` source
mode and MAY use zero to two mechanism packs.

#### Scenario: Generate from a mechanism pack
- **WHEN** the Agent selects an academic-research mechanism
- **THEN** it still chooses deck-specific fonts, colors, layout rhythm, and visual treatment instead of copying a fixed style recipe

#### Scenario: Follow a user template
- **WHEN** a user supplies an authoritative PPTX template
- **THEN** the Agent uses `designProfile` and template planning evidence, does not mix in Grid, and records source artifact identity in the plan

### Requirement: Grid is an explicit scaffold
The bundled Grid Layout library SHALL remain installable and usable, but SHALL
only be selected explicitly by the user, by an authoring plan, or by the frozen
A evaluation arm. Failure in another authoring route SHALL NOT silently switch
to Grid.

#### Scenario: Freeform plan fails validation
- **WHEN** a self-directed plan or candidate cannot pass its required checks
- **THEN** the Agent reports or repairs the failure without silently rebuilding the deck with Grid

### Requirement: Progressive Presentation task routes
The main Presentations Skill SHALL remain at most 350 lines and SHALL route to
exactly five task documents for creation, template creation, existing-deck
editing, continuation, and review/delivery. Advanced native-object procedures
SHALL remain in one-level references loaded only by a relevant route.

#### Scenario: Create a new deck
- **WHEN** a fresh Agent receives a net-new presentation request
- **THEN** it reads the create route and only the selected design/text/review references rather than the complete native-edit catalog

#### Scenario: Edit imported SmartArt
- **WHEN** the request targets imported SmartArt text
- **THEN** the edit-existing route directs the Agent to the advanced native reference and its bounded capability workflow

### Requirement: Audience-facing editorial continuity
Creation SHALL lock evidence before audience editing and slide-fit compression.
A local revision SHALL preserve untouched page copy and the task's voice,
title, density, font, and color roles. A global editorial or visual rewrite
SHALL require explicit scope.

#### Scenario: Revise one page
- **WHEN** the user asks to sharpen one conclusion slide
- **THEN** only the declared page copy and layout may change and review reports any non-target or design-grammar drift

### Requirement: Plan-bound Presentation review
`reviewArtifact` SHALL accept `authoringPlan` and `changedPageIds` for PPTX.
Its `design` report SHALL fail machine-provable plan mismatches and SHALL report
subjective repetition, rhythm, hierarchy, and title-form concerns only as
warnings. Visual review status SHALL remain explicit.

#### Scenario: Strict grammar violation
- **WHEN** a candidate uses a font or color outside a plan field declared strict
- **THEN** design review fails with the page and offending role/value

#### Scenario: Repetitive composition
- **WHEN** several pages share a suspiciously similar silhouette but satisfy all explicit invariants
- **THEN** design review emits a warning and does not claim aesthetic failure

#### Scenario: No image understanding
- **WHEN** visual review is unavailable
- **THEN** structural and design checks still run and the report retains `visualReview: unavailable`
