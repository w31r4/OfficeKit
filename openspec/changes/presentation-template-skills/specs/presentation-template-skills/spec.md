## ADDED Requirements

### Requirement: A PowerPoint template has one distributable form
The system SHALL accept a reusable PowerPoint template only as a schema-v3
Template Skill containing `SKILL.md`, `artifact-template.json`, Agent metadata,
one preview PNG, and four to six hashed example PNGs.

#### Scenario: Valid style Skill
- **WHEN** a schema-v3 presentation template contains the required guide,
  metadata, preview, example roles, hashes, and provenance
- **THEN** template discovery accepts it without requiring a source PPTX or
  executable layout asset

#### Scenario: Prohibited source-backed contents
- **WHEN** a presentation template contains a PPTX, JavaScript, TypeScript,
  page DSL, SVG page skeleton, or fixed layout registry
- **THEN** validation rejects the template as outside the single supported form

### Requirement: Presentation schema v2 is unsupported
The system MUST reject schema-v2 metadata whose kind is `presentation`, while
continuing to accept existing schema-v2 document and spreadsheet templates.

#### Scenario: Old presentation template discovered
- **WHEN** discovery encounters a schema-v2 presentation template
- **THEN** it reports an invalid candidate with guidance to rebuild it through
  the presentation template creator and does not materialize its reference

### Requirement: Search returns style evidence
Presentation search SHALL preserve local ranking and `selectionMade: false`
while returning the candidate Skill, preview, example images, search traits,
and provenance without a reference-file or edit-profile contract.

#### Scenario: Search finds a presentation style
- **WHEN** an Agent searches for a presentation template whose English metadata
  matches the request
- **THEN** the result includes canonical guide and image paths with verified
  hashes and does not choose the template for the Agent

#### Scenario: Search finds no suitable style
- **WHEN** all candidates are avoided, invalid, or below the request's needs
- **THEN** search succeeds with no selection so the Agent can use self-directed
  design

### Requirement: A selected template conditions free composition
The Presentations workflow SHALL allow at most one selected Template Skill,
derive a deck-specific Design Grammar from its guide and examples, and compose
new pages with OfficeKit primitives rather than cloning a template artifact.

#### Scenario: Selected style is used
- **WHEN** an Agent selects a valid style Skill
- **THEN** it reads the guide and examples, records the selected ID and hashes,
  creates new page compositions, and reviews the rendered result

#### Scenario: Design-system conflict
- **WHEN** a user-provided design system conflicts with a selected template
- **THEN** the design system remains authoritative and the Agent does not blend
  incompatible template rules

### Requirement: Reference decks remain a separate workflow
The system SHALL describe uploaded PPTX files as reference decks or
source-continuation artifacts, never as catalog templates.

#### Scenario: User uploads a PPTX for reuse
- **WHEN** the user asks to create from or modify an uploaded presentation
- **THEN** the Agent uses reference-deck analysis or source-bound continuation
  without registering the PPTX as a template

### Requirement: The bundled library uses only the new form
The package SHALL distribute the seven existing presentation IDs and Grid ID
as eight schema-v3 style Skills and SHALL exclude their previous PPTX, preview,
fixed-layout, and runtime assets.

#### Scenario: Installed package inventory
- **WHEN** the packed OfficeKit distribution is inspected
- **THEN** all eight presentation templates validate as style Skills and no old
  presentation template artifact or embedded Grid implementation is present
