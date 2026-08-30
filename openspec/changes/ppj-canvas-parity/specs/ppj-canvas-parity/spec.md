## ADDED Requirements

### Requirement: Imported canvas capability
An exact imported PPTX canvas SHALL project a nativeRef that authorizes only
bounded width and height changes.

#### Scenario: Agent inspects an imported canvas
- **WHEN** PPTX projects to PPJ
- **THEN** `design.canvas.nativeRef` advertises `setCanvas` for
  `canvas.width` and `canvas.height`

### Requirement: Canvas-only source mutation
A capable source-bound canvas edit SHALL change the native presentation canvas
without scaling or rebuilding page content.

#### Scenario: Change one dimension
- **WHEN** an Agent changes a capable canvas width or height
- **THEN** build changes only `ppt/presentation.xml`, reimport recovers the
  requested dimensions, and all page-local object identities remain stable

### Requirement: Canvas mutations fail closed
The compiler SHALL reject a canvas change without the exact issued nativeRef or
with invalid native dimensions.

#### Scenario: Fabricated capability
- **WHEN** an Agent changes canvas state without an issued `setCanvas`
- **THEN** build rejects the program instead of returning unchanged PPTX bytes

### Requirement: Agent discoverability
The generated PPJ reference SHALL explain that source-bound canvas editing does
not scale, reflow, crop, or move existing content.

#### Scenario: Agent plans aspect-ratio conversion
- **WHEN** an Agent reads the canvas reference
- **THEN** it knows to recompose affected pages explicitly and render all pages
