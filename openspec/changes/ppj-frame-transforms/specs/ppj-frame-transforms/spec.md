## ADDED Requirements

### Requirement: PPJ owns bounded rectangular frame transforms

The authored compiler SHALL compile rotation and reflection for shape, image,
chart, table, and group frames using the existing PPJ `frame` state.

#### Scenario: Rotated chart

- **WHEN** a source-free chart frame declares rotation and reflection
- **THEN** the PPTX contains one editable native graphic-frame transform and
  re-projection restores the same PPJ state

#### Scenario: Rotated group

- **WHEN** a source-free group declares an outer rotation
- **THEN** the group child coordinate system and ordered children remain intact

### Requirement: Imported frame transforms remain safe and explicit

The projector SHALL expose recognized frame transforms and source-bound build
SHALL change them only under an issued `setFrame` capability.

#### Scenario: Capability-issued transform edit

- **WHEN** an imported modeled chart, table, group, shape, or image has a safe
  transform and its PPJ frame changes
- **THEN** the compiler updates only the target frame and required slide part

#### Scenario: Connector transform refusal

- **WHEN** a connector frame declares independent rotation or reflection
- **THEN** authored or source-bound build rejects the incompatible state
