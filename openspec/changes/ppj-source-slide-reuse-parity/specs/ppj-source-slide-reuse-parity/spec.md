## ADDED Requirements

### Requirement: Source slide reuse projection
An imported page with a proven native clone graph SHALL advertise a bounded
page-duplication capability in PPJ.

#### Scenario: Agent discovers a reusable page
- **WHEN** an imported slide's clone analysis is known and supported
- **THEN** its page nativeRef advertises `duplicate` for `pageClone`

### Requirement: Declarative pending clone
A source-bound PPJ SHALL represent one pending exact source slide clone as a
finite `sourceClone` page descriptor.

#### Scenario: Reuse a source page
- **WHEN** an Agent inserts a fresh empty page immediately after a retained
  source page and references that page's issued duplicate capability
- **THEN** build creates a distinct SlidePart through the existing native clone
  writer while preserving the original source page

### Requirement: Pending clone immutability
The compiler SHALL reject edits to a pending source clone until it has been
built and reimported as an ordinary source-bound page.

#### Scenario: Agent tries to edit while cloning
- **WHEN** a `sourceClone` page also declares native elements, layout,
  background, notes, transition, animation, visibility, or a nativeRef
- **THEN** check or build rejects the program instead of reserializing or
  flattening the source slide

### Requirement: Clone authority fails closed
The compiler SHALL re-prove the source revision, source page, native object
hash, capability set, ownership graph, adjacency, and one-clone budget.

#### Scenario: Stale or invented clone request
- **WHEN** the page, capability ID, source revision, adjacency, or ownership
  evidence does not match the fresh source projection
- **THEN** build rejects the request without creating an output PPTX

### Requirement: Route topology remains stable
Pending source slide reuse SHALL not silently add the clone to sections,
custom shows, comments, or other source-owned route graphs.

#### Scenario: Clone and route edit are combined
- **WHEN** the same PPJ build both inserts a sourceClone and changes sections
  or custom shows
- **THEN** build rejects the route edit and preserves the source graphs

### Requirement: Agent discoverability
The generated PPJ reference SHALL explain how to inspect a page capability,
insert a sourceClone, build, reimport, and only then edit the new page.

#### Scenario: Fresh Agent continues a template
- **WHEN** an Agent reads the PPJ continuation guide
- **THEN** it can reuse one safe page without using MJS, raw OOXML, part paths,
  or a procedural clone API
