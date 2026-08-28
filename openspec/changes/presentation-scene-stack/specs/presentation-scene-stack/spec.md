## ADDED Requirements

### Requirement: Ordered direct-slide scene stack

OfficeKit SHALL retain one bottom-to-top ordered sequence for all direct slide elements while preserving type-specific collections as indexes.

#### Scenario: Mixed authored elements retain insertion order

- **WHEN** an Agent adds an image, shape, chart, table, connector, and text element in a declared order
- **THEN** inspection, preview, export, and reimport SHALL report the same cross-type order

#### Scenario: Group order remains local

- **WHEN** an Agent reorders a group child
- **THEN** only the group's child sequence SHALL change and the direct slide stack SHALL retain the group as one element

### Requirement: Common ordering operations

Every supported visual element SHALL expose one reorder capability and common bottom/top/before/after operations within its owner.

#### Scenario: Cross-type authored reorder

- **WHEN** an authored image is moved behind a scrim shape and editable text
- **THEN** the exported PresentationML shape tree SHALL place the image first, followed by the scrim and text

#### Scenario: Invalid owner or stale source

- **WHEN** a target belongs to another slide/group or an imported capability no longer matches the source revision
- **THEN** OfficeKit SHALL reject the operation without changing the model

### Requirement: Editable background-image authoring

OfficeKit SHALL provide a background-image convenience that creates an ordinary full-slide picture at the bottom of the scene stack.

#### Scenario: Photo, scrim, foreground composition

- **WHEN** an Agent sets a background image, adds a semi-transparent scrim, and adds editable foreground content
- **THEN** the picture SHALL remain editable and the rendered foreground SHALL remain visible above it

### Requirement: Source-bound imported reorder

OfficeKit SHALL preserve imported order by default and SHALL reorder existing native direct nodes only under an independently re-proved capability.

#### Scenario: Safe local reorder

- **WHEN** a source-bound direct element has a valid reorder capability and no dependency-sensitive graph
- **THEN** the codec SHALL move the existing node in the target SlidePart while preserving every unrelated part and native object

#### Scenario: Unsupported native graph

- **WHEN** order affects an opaque group, animation target, connector dependency, or ambiguous native identity
- **THEN** inspection SHALL explain the blocked reason and export SHALL fail closed rather than flatten or rebuild the slide

### Requirement: Layer-aware Agent guidance

The Presentations Skill and Presentation Template Creator SHALL teach relationship-first layer composition and review.

#### Scenario: Image-led page

- **WHEN** a page uses a full-bleed image
- **THEN** the Agent SHALL verify crop, scrim contrast, foreground readability, source/rights, stack order, and obstruction before delivery
