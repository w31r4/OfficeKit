## Purpose

Expose the text-layout rectangle of a custom shape in PPJ while retaining numeric units, reference identity and source-bound edit fidelity.

## ADDED Requirements

### Requirement: Custom shapes express their text rectangle

Literal custom shapes SHALL accept geometry.textRectangle with left/top/right/bottom, each a shape-local point number or native built-in reference. Authored export and fresh projection SHALL retain literals, references, mixed edges and omission. Invalid or unresolved rectangles MUST reject. Unmodeled custom guide graphs SHALL remain source-owned. Image masks and clips MUST reject this shape-only state rather than discard it.

#### Scenario: Mixed rectangle
- **WHEN** a custom shape supplies numeric left/top and reference right/bottom
- **THEN** native output and fresh PPJ SHALL preserve both representations without replacing references with numbers

### Requirement: Source editors add change and remove rectangles

Issued setGeometry authority SHALL permit geometry.textRectangle addition, change and removal. Unchanged paths, text, frame and non-target package members SHALL be preserved. No-op source compilation SHALL preserve bytes.

#### Scenario: Independent source rectangle edit
- **WHEN** an original-source request changes a rectangle edge or removes the rectangle
- **THEN** fresh projection SHALL show the requested state and only the target SlidePart SHALL change

### Requirement: Preview retains text-layout limitations

Preview SHALL diagnose text-rectangle state it cannot fully render, without reporting complete visual support.

#### Scenario: Rectangle preview
- **WHEN** a custom rectangle is outside the painter's supported text-layout mapping
- **THEN** an explicit machine-readable limitation SHALL remain in the preview assessment
