## ADDED Requirements

### Requirement: PPJ exposes the complete bounded preset vocabulary

The PPJ schema, authored compiler, imported projector, and generated reference
SHALL share one versioned registry containing every non-connector DrawingML
preset supported by the pinned Office schema.

#### Scenario: Newly exposed native preset

- **WHEN** a source-free PPJ shape uses a registry preset that was absent from
  the former hand-picked subset
- **THEN** build writes the exact native preset token and re-import restores the
  same stable PPJ geometry and ordered adjustments

#### Scenario: Connector token

- **WHEN** a PPJ shape attempts to use a native connector preset
- **THEN** schema validation rejects it and routes line relationships to the
  typed connector element instead

### Requirement: Preset picture masks retain ordered adjustments

The authored compiler SHALL accept a complete ordered adjustment array on a
preset image mask and SHALL emit a canonical literal native adjustment list.
Projection SHALL restore the same PPJ state.

#### Scenario: Adjusted rounded picture mask

- **WHEN** an image declares a `roundRect` mask with one adjustment value
- **THEN** build, re-import, and a second build retain the preset and exact
  ordered value

### Requirement: Imported mask edits are capability-bound

An imported editable picture SHALL receive `setImageMask` only when its preset
identity and adjustment list are complete, literal, canonical, and safely
owned. The capability SHALL allow only `mask.adjustments` to change.

#### Scenario: Capability-issued mask adjustment

- **WHEN** PPJ changes only an imported picture's issued mask adjustment array
- **THEN** source-bound build updates the target `a:avLst` while preserving the
  source asset, crop, frame, relationships, and non-target package content

#### Scenario: Custom or formula mask

- **WHEN** an imported picture mask uses custom geometry or a noncanonical
  formula graph
- **THEN** PPJ preserves it through the bound source and refuses a mask edit

### Requirement: Agent guidance remains searchable and purposeful

The generated PPJ reference SHALL list every supported preset, native defaults,
parameter order, and picture-mask reuse. Human guidance SHALL state that preset
availability is not a reason to add decorative geometry without information
purpose.

#### Scenario: Agent discovers a mask profile

- **WHEN** an Agent searches the generated reference for a rounded image or a
  less-common native shape
- **THEN** it finds the valid PPJ preset, complete adjustment order, defaults,
  and the rule that geometry must serve the page's information structure
