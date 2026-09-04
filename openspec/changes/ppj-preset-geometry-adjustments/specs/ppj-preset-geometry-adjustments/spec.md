## ADDED Requirements

### Requirement: PPJ compiles canonical preset adjustments

The authored compiler SHALL compile a complete ordered integer adjustment array
for each supported PPJ preset geometry into canonical literal DrawingML guides.
An omitted or empty array SHALL use the native preset defaults.

#### Scenario: Authored rounded rectangle

- **WHEN** a source-free `roundRect` declares `adjustments: [24000]`
- **THEN** build writes one `adj` guide with formula `val 24000` and re-import
  restores the same PPJ array

#### Scenario: Invalid adjustment arity

- **WHEN** a preset declares a non-empty array whose length differs from its
  canonical profile
- **THEN** check and build reject it before writing PPTX

### Requirement: Imported canonical adjustments are editable state

The PPTX projector SHALL restore canonical literal preset adjustment guides as
ordered PPJ values and SHALL issue `setGeometry` only for
`geometry.adjustments` on shapes, or the bounded `setImageMask` fields on
recognized picture masks, when the source object is otherwise safely editable.

#### Scenario: Capability-issued arrow edit

- **WHEN** an imported editable arrow has canonical `adj1` and `adj2` literal
  guides and the PPJ changes only those values
- **THEN** source-bound build updates the target guide values and preserves
  non-target OPC parts and source topology

#### Scenario: Picture preset identity replaces adjustment arity

- **WHEN** an imported editable picture has a canonical `roundRect` mask with
  one literal adjustment and the PPJ changes it to a supported `snipRoundRect`
  mask with its two literal adjustments through `setImageMask`
- **THEN** source-bound build replaces only the picture-owned `a:prstGeom`
  slot on the existing SlidePart, preserves the picture relationship, and a
  second projection restores the requested preset and ordered values

#### Scenario: Noncanonical formula remains source-owned

- **WHEN** an imported preset shape uses a formula-valued, partial, reordered,
  duplicated, or unknown adjustment guide
- **THEN** PPJ does not issue geometry capability and an attempted adjustment
  edit fails closed without normalizing the native graph

### Requirement: Agent documentation describes adjustment semantics

The generated PPJ reference SHALL identify the ordered-array contract, native
default behavior, numeric units, supported preset profiles, and imported
capability boundary.

#### Scenario: Primitive discovery

- **WHEN** an Agent searches the PPJ reference for rounded corners, arrows, or
  shape adjustments
- **THEN** it can find the writable PPJ field, an executable example, and the
  fail-closed imported boundary without consulting C# or raw OOXML
