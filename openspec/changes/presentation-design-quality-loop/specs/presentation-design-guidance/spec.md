## ADDED Requirements

### Requirement: Selected scenario provides concrete design decisions
OfficeKit SHALL load only the chosen primary scenario guide, plus one secondary
guide when explicitly justified, and that guide SHALL cover evidence hierarchy,
page archetypes, visual carriers, data treatment, imagery, density rhythm,
geometry, line, surface, typography, and scenario-specific repair questions.

#### Scenario: Self-directed analysis deck
- **WHEN** the Agent selects `analysis-decision` as the primary scenario
- **THEN** it receives concrete analysis-specific decisions without loading all
  other scenario guides

### Requirement: Guidance remains generative rather than templated
Scenario guidance MUST describe communication purpose, relationships, ranges,
and anti-patterns without requiring fixed palettes, fonts, layout IDs, or page
sequences.

#### Scenario: Two decks share a scenario
- **WHEN** two unrelated briefs both select `management-report`
- **THEN** each deck can form a different design grammar while following the same
  management communication principles

#### Scenario: User supplies a brand system
- **WHEN** an authoritative design system conflicts with a scenario default
- **THEN** the user-provided system wins and the scenario guide fills only
  unspecified decisions

### Requirement: Visual carriers name an honest source strategy
Every planned page SHALL identify the carrier that communicates its claim and
whether it comes from a user asset, reference/template asset, sourced image,
generated image, or native OfficeKit composition.

#### Scenario: Image generation is unavailable
- **WHEN** a page can communicate its claim with a chart, diagram, table,
  typography, or native vector composition
- **THEN** the Agent selects that carrier without adding an unrelated image
  placeholder

#### Scenario: A specific image is essential
- **WHEN** the task requires understanding or showing a subject for which no
  usable asset is available
- **THEN** the Agent requests the asset or reports the no-image limitation
  instead of inventing evidence

### Requirement: Examples teach decisions through public primitives
OfficeKit SHALL provide public-API examples for de-defaulted charts, editorial
rules, image treatment, diagrams, asymmetry, and motif use, and those examples
MUST explain the information relationship they implement.

#### Scenario: Agent needs a financial chart
- **WHEN** the Agent searches Help for a styled evidence chart
- **THEN** it can find an example that integrates chart surface, axes, labels,
  series, source, and page hierarchy without prescribing one fixed palette

### Requirement: Weak pages use a layered repair order
The Presentations workflow SHALL repair claim/evidence, visual carrier,
composition/hierarchy, and styling before adding motion or decoration.

#### Scenario: Page feels visually empty
- **WHEN** the page lacks an informative relationship
- **THEN** the Agent changes the carrier or composition rather than filling space
  with decorative shapes
