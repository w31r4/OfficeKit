## ADDED Requirements

### Requirement: PPJ declares bounded text effects
PPJ SHALL allow a text style to declare either a solid `color` or a bounded direct-RGB `gradient`, plus one optional bounded direct outer `shadow`.

#### Scenario: Valid display text style
- **WHEN** a direct run or paragraph default text style declares an ordered linear or centered-radial gradient and a valid outer shadow
- **THEN** PPJ validation accepts the program and retains every stop and shadow parameter

#### Scenario: Conflicting or unbounded text style
- **WHEN** a style declares both solid color and gradient, an invalid stop graph, or invalid shadow geometry
- **THEN** validation fails before package output with a path-specific diagnostic

### Requirement: Text effects compile to native editable DrawingML
The authored compiler SHALL encode supported text gradient and shadow state directly on the relevant DrawingML run or default-run properties.

#### Scenario: Authored rich title
- **WHEN** a PPJ title run declares a supported gradient and shadow
- **THEN** the output contains canonical `a:gradFill` and `a:effectLst/a:outerShdw` children and the text remains native and editable

#### Scenario: Paragraph default effect
- **WHEN** a paragraph default text style declares a supported effect
- **THEN** the compiler writes the effect to `a:defRPr` so descendant runs inherit it without duplicated direct formatting

### Requirement: Import preserves proven and unknown text effects
OfficeKit SHALL project only canonical text gradient and outer-shadow graphs while preserving unsupported source graphs without semantic flattening.

#### Scenario: Canonical effect reimport
- **WHEN** a PPTX contains the bounded canonical text gradient and shadow profile
- **THEN** PPJ projection restores equivalent gradient stops and shadow parameters

#### Scenario: Unknown effect graph
- **WHEN** a third-party run contains theme-transformed gradients, glow, reflection or another unsupported effect graph
- **THEN** unrelated edits preserve the native graph and a conflicting effect mutation fails closed
