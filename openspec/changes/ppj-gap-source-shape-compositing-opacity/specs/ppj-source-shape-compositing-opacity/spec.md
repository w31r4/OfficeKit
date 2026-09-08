## ADDED Requirements

### Requirement: Source-bound compound shape opacity has one provable owner value

When an imported ordinary non-placeholder shape has no visible text and every
modeled visual paint owner has the same effective alpha, the projected PPJ
shape MUST expose `compositing.opacity` and its native reference MUST issue
`setOpacity` for `compositing.opacity`.

#### Scenario: Matching fill and outline alpha are projected

- **WHEN** a source-bound shape has a direct fill and visible outline whose
  effective alpha values are both `0.4`
- **THEN** its PPJ projection contains `compositing.opacity` equal to `0.4`
  and advertises `setOpacity` for `compositing.opacity`

#### Scenario: Mixed visual-owner alpha remains source-bound

- **WHEN** a source-bound shape's fill and outline have different effective
  alpha values
- **THEN** the projection omits `compositing` and does not issue the compound
  opacity capability

### Requirement: Source-bound compound opacity lowers to all modeled owners

When a source-bound PPJ changes only an eligible shape's
`compositing.opacity`, the compiler MUST require the exact
`setOpacity/compositing.opacity` capability and write the requested absolute
alpha to every modeled fill, gradient-stop, image-fill, outline, and shadow
owner.

#### Scenario: A shape opacity edit changes only its slide owner

- **WHEN** a projected eligible shape changes `compositing.opacity` from `1`
  to `0.4`
- **THEN** the compiled package preserves unrelated parts and the shape's
  modeled fill and outline both carry alpha `40000`

#### Scenario: Style and compound opacity cannot be changed ambiguously

- **WHEN** one source-bound transaction changes both `style` and
  `compositing.opacity` on an eligible shape
- **THEN** the compiler rejects the transaction instead of guessing the new
  compound owner set

### Requirement: Unsupported compound topology fails closed

Source-bound shapes with visible text, placeholder identity, line geometry, mixed
owner alpha, or an opaque custom image-fill graph MUST NOT be presented as a
compound opacity owner.

#### Scenario: Text-bearing shape keeps its existing direct semantics

- **WHEN** a source-bound shape contains a text body
- **THEN** the projector does not issue the source-bound compound opacity
  capability
