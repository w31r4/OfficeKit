## Purpose

This capability gives imported custom geometry one safe, independently editable guide value while preserving the formula graph and every unsupported geometry feature as source-owned data.

## ADDED Requirements

### Requirement: Expose literal custom-geometry guide values as bounded native leaves

The system SHALL expose an existing custom-geometry guide as a `customGeometryGuide` native leaf only when the guide is a direct, uniquely located `a:gd` entry whose formula is exactly `val N`, where `N` is a bounded signed integer. The leaf SHALL retain the guide's source position and value proof without presenting calculated guides, handles, connection sites, or path topology as editable semantics.

#### Scenario: Project a literal guide without flattening its geometry graph

- **WHEN** an imported custom geometry contains a uniquely located `a:gd` with a non-empty name, no children or extension attributes, and `fmla="val N"`
- **THEN** the projected element exposes a `customGeometryGuide` native leaf with the canonical integer value and retains the custom geometry as the source-bound geometry owner

#### Scenario: Edit one literal guide in place

- **WHEN** a source-bound PPJ edit changes the value of an issued `customGeometryGuide` leaf while retaining its identity and preconditions
- **THEN** the system changes only that guide's `fmla` value in the owning slide part, preserves the guide name and all other geometry XML, and a subsequent projection reports the new value

#### Scenario: Keep unsupported guide graphs source-owned

- **WHEN** a guide is calculated, malformed, duplicated, extension-bearing, child-bearing, or outside the bounded signed-integer range
- **THEN** the system SHALL NOT issue a `customGeometryGuide` leaf for that guide and SHALL preserve or fail closed on edits rather than guessing a partial geometry model
