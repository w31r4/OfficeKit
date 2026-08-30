## ADDED Requirements

### Requirement: PPJ authors native custom image masks
The authored PPJ compiler SHALL accept the bounded custom geometry form at `image.mask` and emit the same ordered path graph as native DrawingML custom geometry on the picture.

#### Scenario: Authored custom mask
- **WHEN** a PPJ image declares a valid literal custom mask with a finite view box and supported commands
- **THEN** the output PPTX contains an editable native picture whose shape properties own an `a:custGeom` matching that PPJ path graph

#### Scenario: Unsupported path topology
- **WHEN** a mask contains an unsupported command, formula graph, guide, handle, or exceeds a custom-geometry budget
- **THEN** validation or build fails before package mutation with a path-specific diagnostic

### Requirement: Canonical custom masks reproject to PPJ
The PPTX projector SHALL recover a custom image mask only when its native geometry maps exactly to the bounded literal PPJ geometry form.

#### Scenario: Canonical reimport
- **WHEN** an authored or third-party picture uses one supported literal custom geometry and otherwise matches the bounded picture profile
- **THEN** the projected image includes `mask.kind: "custom"` with stable ordered paths and retains the picture asset, frame, crop, opacity, border, and accessibility state

#### Scenario: Irregular native mask
- **WHEN** a picture mask contains unmodeled formulas, handles, connection sites, effects, or ambiguous geometry ownership
- **THEN** OfficeKit preserves it as opaque/source-owned content and does not simplify it into PPJ

### Requirement: Source-bound custom mask topology is immutable
Imported custom mask paths MUST remain unchanged during unrelated capability-issued picture edits.

#### Scenario: Unrelated picture edit
- **WHEN** an imported canonical custom-mask picture receives a proven frame, crop, opacity, accessibility, or same-format asset edit without changing its mask graph
- **THEN** the compiler preserves the custom geometry and applies only the declared mutation

#### Scenario: Path mutation without capability
- **WHEN** requested PPJ changes any custom-mask path, command, coordinate, order, or presence
- **THEN** the compiler rejects the mutation before altering the source package
