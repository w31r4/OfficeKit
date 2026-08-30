## ADDED Requirements

### Requirement: Authored visual fields SHALL NOT be silently ignored

The source-free PPJ compiler MUST either lower every present visual field to its
typed native semantic owner or reject the program before emitting a PPTX.

#### Scenario: Unsupported style is present

- **WHEN** a valid PPJ includes a visual field without an authored compiler owner
- **THEN** build fails with a path-specific unsupported-feature diagnostic
- **AND** no output PPTX is returned

### Requirement: PPJ custom paths SHALL compile through bounded native geometry

The compiler MUST normalize finite custom path coordinates and lower supported
move, line, quadratic, cubic, and close commands through the existing native
custom-geometry validator and writer.

#### Scenario: Custom path is authored

- **WHEN** a valid source-free PPJ contains a bounded custom path
- **THEN** build emits editable DrawingML custom geometry
- **AND** reimport recovers its path commands without rasterization

### Requirement: Typed paint SHALL remain closed and source-safe

Gradient fills and line opacity MUST use typed values with bounded stops, angle,
and opacity. Raw DrawingML, arbitrary paint nodes, and imported topology changes
MUST remain unavailable.

#### Scenario: Linear gradient is authored

- **WHEN** a valid PPJ applies a bounded linear gradient to an authored shape
- **THEN** build emits a native editable gradient fill
- **AND** reimport reports equivalent stop colors, positions, alpha, and angle

### Requirement: Data-visual styling SHALL be operational or explicit

Every PPJ chart and table style field MUST compile to a bounded native semantic
owner or fail with a path-specific unsupported diagnostic.

#### Scenario: Agent uses a supported chart style

- **WHEN** PPJ specifies supported legend, stacking, gap, axis, gridline, or area-fill state
- **THEN** the resulting native chart visibly and structurally reflects that state
