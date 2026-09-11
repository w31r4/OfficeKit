## ADDED Requirements

### Requirement: Direct rich-text reflection exposes horizontal skew

The PPJ contract MUST represent ` run.style.reflection.skewX ` as a finite degree value strictly between -90 and 90 after native 1/60000-degree rounding. Authored direct runs and a reprojected native leaf MUST recover the canonical value.

#### Scenario: Authored reflection skewX round-trips

- **WHEN** an authored direct run declares a valid reflection with `skewX: -12.5`
- **THEN** the compiler writes native `kx` value `-750000`
- **AND** stripping and projecting the package exposes `reflection.skewX` as -12.5 and a `textReflectionSkewX` leaf with the same native value.

### Requirement: Only a proven direct owner receives the leaf

The native projection MUST issue `textReflectionSkewX` only for one direct run `a:reflection` with full-span positions, one canonical `kx`, and no fade, scale, `ky`, alignment, or `rotWithShape` transform. Other or ambiguous effect graphs MUST remain source-owned or fail closed.

#### Scenario: Unsupported transforms stay opaque

- **WHEN** the reflection has a variable span, another transform, an unknown child/attribute, or a duplicate effect owner
- **THEN** no `textReflectionSkewX` capability is issued.

### Requirement: Source-bound edits splice one native token

An edit through `textReflectionSkewX` MUST replace only the proven reflection `@kx` token, preserve other effects and run topology, and write only the owning SlidePart.

#### Scenario: SkewX edit preserves the package

- **WHEN** a source-bound leaf changes from -12.5 to 20
- **THEN** only the owning SlidePart changes
- **AND** the Open XML package remains valid, all non-target parts remain byte-identical, and a second projection reports 20.
