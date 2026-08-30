## ADDED Requirements

### Requirement: Typed authored element state
The NativeAOT compiler SHALL accept optional PPJ `hidden` and `locked` state on
source-free typed elements and SHALL lower it without changing element ID,
z-order, content, relationships, accessibility, or frame.

#### Scenario: Locked hidden guide shape
- **WHEN** a valid PPJ page declares a shape with `hidden: true` and
  `locked: true`
- **THEN** build writes hidden non-visual state and the canonical full shape
  lock profile while retaining the shape's stable identity and geometry

### Requirement: Canonical cross-type locking
The writer SHALL map PPJ locking to type-appropriate standard DrawingML locks
for shapes, pictures/media, connectors, graphic-frame charts/tables, and
groups, and SHALL distinguish those locks from the native baseline required by
the object type.

#### Scenario: Unlocked chart retains baseline
- **WHEN** an authored chart declares `locked: false`
- **THEN** its graphic frame retains the canonical `noGrouping` baseline but
  does not receive selection, movement, resizing, aspect, or drill-down locks

### Requirement: Conservative imported projection
The PPTX projector SHALL expose imported `hidden` state and SHALL expose
`locked` plus a `setState` capability only when the exact native lock profile
is recognized as the canonical locked or unlocked state. Other lock profiles
SHALL remain source-preserved and uneditable through this boolean.

#### Scenario: Imported partial picture lock
- **WHEN** a third-party picture contains only `noCrop` plus an unknown lock
  extension
- **THEN** projection does not claim a PPJ locked value or issue `setState`, and
  an attempted state edit is rejected without changing the source package

### Requirement: Source-bound local state edit
The source-bound compiler SHALL require a fresh `setState` capability and the
PPTX exporter SHALL re-prove the source object and lock profile before changing
only hidden/lock state.

#### Scenario: Lock a recognized imported text box
- **WHEN** an Agent changes `locked` from false to true on a hash-bound imported
  text box whose capability includes `setState`
- **THEN** build mutates only the owning SlidePart's canonical lock nodes,
  preserves the text and all non-target parts, and reimport reports the same
  element ID with `locked: true`

### Requirement: Agent discoverability
The generated PPJ reference and focused shape/layer guidance SHALL explain the
cross-type meaning, the difference from slide hiding and protection, and the
source-bound failure boundary.

#### Scenario: Agent protects a background composition
- **WHEN** an Agent searches the PPJ guidance for locking or hiding a visual
  element
- **THEN** it can choose element `locked`/`hidden`, understand that arrays still
  control z-order, and avoid treating the fields as document security

