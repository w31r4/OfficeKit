## ADDED Requirements

### Requirement: Stable imported page identity
Source-projected PPJ SHALL derive page identity from stable source ownership
rather than the page's current presentation position.

#### Scenario: Reimport after a page move
- **WHEN** a capable imported page is moved and the result is projected again
- **THEN** that page and all of its unchanged page-local elements retain their
  PPJ IDs

### Requirement: Bounded source page reorder
A source-bound page SHALL advertise `reorder/pageOrder` only when the native
presentation can preserve its page and section graph through a complete
retained-page permutation.

#### Scenario: Move one retained page
- **WHEN** an Agent reorders the `pages` array using only capable existing pages
- **THEN** build changes only `ppt/presentation.xml`, retains the source
  SlideParts and relationships, and reimport recovers the exact page order

### Requirement: Related routes retain page references
Modeled comments, sections, and custom shows SHALL continue to reference stable
page IDs across a page reorder.

#### Scenario: Reorder a deck with modeled routes
- **WHEN** an Agent moves pages and supplies a valid section partition
- **THEN** comments stay attached to their pages and custom-show membership
  remains unchanged unless explicitly edited

### Requirement: Opaque section graphs fail closed
A presentation with an opaque section graph SHALL NOT advertise page reorder.

#### Scenario: Reorder without issued capability
- **WHEN** an Agent changes page order on such a projection
- **THEN** build rejects the mutation without rewriting the source package

### Requirement: Agent discoverability
The generated PPJ reference SHALL explain that `pages[]` is presentation order,
stable IDs survive capable moves, and modeled sections may need a matching
partition update.

#### Scenario: Agent plans a page move
- **WHEN** an Agent reads the PPJ reference and an imported page nativeRef
- **THEN** it can distinguish `reorder/pageOrder` from element `reorder/zOrder`
