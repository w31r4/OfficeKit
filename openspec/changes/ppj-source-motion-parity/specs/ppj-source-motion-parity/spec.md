## ADDED Requirements

### Requirement: Imported motion capability
An imported page SHALL advertise `setAnimations/animations` only when its
native timing graph is canonical and editable, or absent and safely addable.

#### Scenario: Agent discovers a safe motion page
- **WHEN** the fresh source projection proves timing editable or addable and
  the page does not participate in Morph
- **THEN** the page nativeRef contains the bounded animation capability

### Requirement: Declarative animation replacement
The complete PPJ `pages[].animations[]` array SHALL be the requested canonical
object timing state for a capable imported page.

#### Scenario: Agent adds one animation
- **WHEN** an Agent adds a valid typed animation targeting an existing source
  element and keeps the page capability unchanged
- **THEN** build writes canonical native timing through the existing timing
  codec and second projection recovers the typed animation

#### Scenario: Agent removes canonical timing
- **WHEN** an Agent removes all projected animations from a timing-editable page
- **THEN** build removes the canonical native graph without changing unrelated
  SlidePart content

### Requirement: Stable target remapping
Source-bound animation lowering SHALL resolve stable PPJ element IDs to the
exact native semantic element IDs without exposing drawing IDs.

#### Scenario: Animation targets a group descendant
- **WHEN** the target is a projected child in an unchanged source group
- **THEN** build resolves its wire identity from the fresh parallel tree or
  fails closed on topology drift

### Requirement: Opaque and Morph timing remain source-owned
The compiler SHALL not replace an opaque timing graph or a Morph timing graph
through `animations[]`.

#### Scenario: Page has unsupported timing or Morph
- **WHEN** an Agent changes animations without an issued capability
- **THEN** check or build rejects the change and preserves the source package

### Requirement: Structural evidence is not playback evidence
Review and Skill guidance SHALL distinguish canonical timing round-trip from
actual desktop-host playback.

#### Scenario: Native round-trip succeeds
- **WHEN** build and second projection recover the requested animation
- **THEN** the result is reported as structural evidence until a real host has
  played the presentation
