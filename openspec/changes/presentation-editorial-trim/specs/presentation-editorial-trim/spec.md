## ADDED Requirements

### Requirement: Focused presentation editorial Skill
OfficeKit SHALL ship a host-neutral presentation editorial Skill that can be
invoked by the Presentations workflow or explicitly for copy-only deck work.

#### Scenario: Create a new deck
- **WHEN** an Agent follows the Presentations create route
- **THEN** it uses the editorial Skill before composition and again after the
  rendered page is available

#### Scenario: Polish an existing deck
- **WHEN** a user asks only to improve the wording of declared slides
- **THEN** the editorial Skill edits only resolved text in that declared scope
  and leaves unsupported or opaque content unchanged

### Requirement: Facts and source wording remain protected
The editorial Skill MUST preserve names, numbers, dates, units, citations,
quotations, uncertainty, technical identifiers, and user-locked wording unless
the user explicitly authorizes a factual change.

#### Scenario: Shortening a qualified claim
- **WHEN** removing words would change the certainty, scope, or causal meaning of
  a claim
- **THEN** the Skill keeps the qualification and changes the composition or page
  structure instead

#### Scenario: Source text contains a flagged pattern
- **WHEN** quoted or source-preserved text contains a long dash, contrast form,
  or other editorial review pattern
- **THEN** the source characters remain unchanged

### Requirement: Copy layers receive different treatment
The editorial Skill SHALL distinguish titles and claims, visible support copy,
labels and sources, and speaker notes before rewriting.

#### Scenario: Live presentation
- **WHEN** the authoring plan delivery mode is `live`
- **THEN** visible copy is compressed for spoken delivery while necessary
  elaboration moves to speaker notes and sources remain visible

#### Scenario: Reader presentation
- **WHEN** the delivery mode is `reader`
- **THEN** visible copy retains the qualifiers and explanation required for
  independent reading

### Requirement: Language patterns guide judgment instead of mechanical bans
The editorial Skill SHALL review false contrast, defensive negation, empty
signposts, abstract noun chains, repeated triads, unsupported superlatives,
continuous metaphors, slogan fragments, and repeated title forms without
mechanically rejecting every matching string.

#### Scenario: Intentional comparison
- **WHEN** a contrast construction represents a real evidenced distinction
- **THEN** the Skill may retain it while checking that the same construction is
  not repeated mechanically across the deck

#### Scenario: Empty rhetorical setup
- **WHEN** a contrast construction invents an unheld position only to reveal the
  actual point
- **THEN** the Skill writes the actual point directly

### Requirement: Page-fit editing follows real composition
The post-render editorial pass SHALL respond to observed wrapping, overflow,
hierarchy, and page density instead of applying a fixed character target before
layout.

#### Scenario: Copy does not fit
- **WHEN** a rendered page cannot hold its protected claim and evidence at the
  declared readability floor
- **THEN** the Agent shortens unprotected repetition, splits the page, or changes
  composition before reducing type size

### Requirement: Deck voice remains varied and scoped
The editorial Skill SHALL review title sequence, terminology, certainty, and
sentence rhythm across the deck while preserving the scope of local edits.

#### Scenario: Repeated title formula
- **WHEN** several slides repeat the same rhetorical opening without a deliberate
  structural reason
- **THEN** the Skill varies the sentence form without changing each page's claim

#### Scenario: Local copy edit
- **WHEN** the user requests a wording change on selected pages
- **THEN** the Skill inherits the existing deck voice and does not rewrite
  unrelated pages
