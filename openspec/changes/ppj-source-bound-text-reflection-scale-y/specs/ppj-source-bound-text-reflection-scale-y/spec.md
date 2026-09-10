## ADDED Requirements

### Requirement: Direct rich-text reflection exposes vertical scale

The PPJ contract MUST represent `run.style.reflection.scaleY` as a finite signed ratio, preserving native 1/100000 precision. Authored direct runs and a reprojected native leaf MUST recover the same value after rounding.

#### Scenario: Authored reflection scaleY round-trips

- **WHEN** an authored direct run declares a valid reflection with `scaleY: 1.25`
- **THEN** the compiler writes the native `sy` value `125000`
- **AND** stripping and projecting the package exposes `reflection.scaleY` as `1.25` and a `textReflectionScaleY` leaf with the same value.

### Requirement: Only a proven direct owner receives the leaf

The native projection MUST issue `textReflectionScaleY` only for one direct run `a:reflection` with full-span positions, one canonical `sy`, and no fade, `sx`, skew, alignment, or `rotWithShape` transform. Other or ambiguous effect graphs MUST remain source-owned or fail closed.

#### Scenario: Unsupported transforms stay opaque

- **WHEN** the reflection has a variable span, another transform, an unknown child/attribute, or a duplicate effect owner
- **THEN** no `textReflectionScaleY` capability is issued.

### Requirement: Source-bound edits splice one native token

An edit through `textReflectionScaleY` MUST replace only the proven reflection `@sy` token, preserve other effects and run topology, and write only the owning SlidePart.

#### Scenario: ScaleY edit preserves the package

- **WHEN** a source-bound leaf changes from `1.25` to `-0.75`
- **THEN** only the owning SlidePart changes
- **AND** the Open XML package remains valid, all non-target parts remain byte-identical, and a second projection reports `-0.75`.
