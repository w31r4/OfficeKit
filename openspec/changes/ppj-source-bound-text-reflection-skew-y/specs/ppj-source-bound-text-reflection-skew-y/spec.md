## ADDED Requirements

### Requirement: Direct rich-text reflection exposes vertical skew

The PPJ contract MUST represent `run.style.reflection.skewY` as a finite signed degree, preserving native 1/60000-degree precision. Authored direct runs and a reprojected native leaf MUST recover the same value after rounding.

#### Scenario: Authored reflection skewY round-trips

- **WHEN** an authored direct run declares a valid reflection with `skewY: -12.5`
- **THEN** the compiler writes the native `ky` value `-750000`
- **AND** stripping and projecting the package exposes `reflection.skewY` as `-12.5` and a `textReflectionSkewY` leaf with the same value.

### Requirement: Only a proven direct owner receives the leaf

The native projection MUST issue `textReflectionSkewY` only for one direct run `a:reflection` with full-span positions, one canonical `ky`, and no fade, `sx`, `sy`, `kx`, alignment, or `rotWithShape` transform. Other or ambiguous effect graphs MUST remain source-owned or fail closed.

#### Scenario: Unsupported transforms stay opaque

- **WHEN** the reflection has a variable span, another transform, an unknown child/attribute, or a duplicate effect owner
- **THEN** no `textReflectionSkewY` capability is issued.

### Requirement: Source-bound edits splice one native token

An edit through `textReflectionSkewY` MUST replace only the proven reflection `@ky` token, preserve other effects and run topology, and write only the owning SlidePart.

#### Scenario: SkewY edit preserves the package

- **WHEN** a source-bound leaf changes from `-12.5` to `20`
- **THEN** only the owning SlidePart changes
- **AND** the Open XML package remains valid, all non-target parts remain byte-identical, and a second projection reports `20`.
