## Purpose

Expose one source-preserving horizontal reflection scale field for imported rich-text runs, so PPJ can edit the canonical DrawingML sx scalar without flattening the surrounding text effect graph.

## ADDED Requirements

### Requirement: Direct rich-text reflection scale X is a bounded PPJ field

The presentation importer and authoring compiler SHALL represent run.style.reflection.scaleX as a signed ratio using native 1/100000 precision. A source-bound direct run SHALL expose the field only when its reflection uses full-span positions and has exactly one supported transform, the canonical sx; scaleY, skew, alignment, rotateWithShape, fadeAngle, variable endpoints, malformed tokens, and complex effect lists SHALL remain source-owned.

#### Scenario: Author and project scale X

- **WHEN** a PPJ direct rich-text run declares a valid reflection with scaleX: 1.25
- **THEN** the authored PPTX stores a:reflection/@sx as 125000 and a projection exposes scaleX as 1.25

#### Scenario: Edit only the canonical scale token

- **WHEN** a projected textReflectionScaleX leaf changes from 1.25 to 0.75 with matching source preconditions
- **THEN** only the owning slide part changes, a:reflection/@sx becomes 75000, and a second projection returns scaleX: 0.75

#### Scenario: Preserve unsupported reflection topology

- **WHEN** a direct run reflection contains skew, a variable endpoint, fadeAngle combined with scaleX, or an ambiguous effect list
- **THEN** the run remains source-owned and no textReflectionScaleX leaf or edit capability is issued
