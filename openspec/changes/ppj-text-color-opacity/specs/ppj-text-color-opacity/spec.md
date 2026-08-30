## ADDED Requirements

### Requirement: PPJ compiles bounded text color opacity

The authored Presentation compiler SHALL accept alpha-bearing PPJ colors for
text runs and default text styles and SHALL emit one canonical direct
DrawingML alpha value.

#### Scenario: Translucent supporting text

- **WHEN** a PPJ text run uses an RGB color with alpha below one
- **THEN** the compiler emits editable native text with the requested direct
  RGB paint and opacity

#### Scenario: Opaque text remains canonical

- **WHEN** text color alpha equals one or is omitted
- **THEN** the compiler omits the native alpha child

### Requirement: Projection preserves recognized text alpha

The PPTX projector SHALL preserve recognized direct text alpha in PPJ and SHALL
not flatten unrecognized color transforms.

#### Scenario: Direct RGB alpha round trip

- **WHEN** an authored PPJ is compiled, imported, and projected again
- **THEN** its text color and alpha remain semantically equal

#### Scenario: Unsupported source mutation

- **WHEN** an Agent changes projected text opacity without an issued capability
- **THEN** source-bound build rejects the mutation and preserves the source
  package
