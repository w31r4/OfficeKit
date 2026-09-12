## Purpose

Expose the direct rich-text run outer-shadow rotation flag as a typed PPJ field while preserving the owning presentation XML and rejecting unsupported effect topology.

## ADDED Requirements

### Requirement: Direct run shadow rotation is typed

An imported direct rich-text run MUST expose `run.style.shadow.rotateWithShape` as a boolean native leaf named `textShadowRotateWithShape` only when its single `a:outerShdw` effect has a canonical `rotWithShape` token (`0` or `1`), bounded blur/distance/direction geometry, one RGB or theme color child, and no `sx`, `sy`, `kx`, `ky`, sibling effect, or unknown descendant. An explicit `0` MUST remain distinguishable from a missing attribute; unsupported or non-canonical topology MUST remain source-owned and produce no typed leaf.

#### Scenario: Canonical rotation flag projects

- **WHEN** a direct rich-text run has a single bounded outer shadow with `rotWithShape="1"`
- **THEN** the PPJ projection emits `textShadowRotateWithShape` with boolean value `true` and the shadow object reports `rotateWithShape: true`

#### Scenario: Explicit false rotation flag projects

- **WHEN** a direct rich-text run has the same bounded outer shadow with `rotWithShape="0"`
- **THEN** the PPJ projection emits `textShadowRotateWithShape` with boolean value `false` and does not treat the field as missing

#### Scenario: Unsupported rotation topology stays opaque

- **WHEN** the direct outer shadow omits `rotWithShape`, uses a non-canonical token, adds a scale/skew transform, sibling effect, or unknown descendant
- **THEN** the PPJ projection emits no `textShadowRotateWithShape` leaf for that run

### Requirement: Rotation edit is source-bound

An edit targeting `textShadowRotateWithShape` MUST prove the same direct outer-shadow owner and replace only its canonical `rotWithShape` token in the owning SlidePart. The edit MUST reject a stale or changed source, preserve all unrelated ZIP entries and XML nodes, and allow a subsequent projection to observe the new boolean value.

#### Scenario: Boolean rotation edit changes one token

- **WHEN** a source-bound edit changes a proven `textShadowRotateWithShape` value from `true` to `false`
- **THEN** only the owning SlidePart is changed, only `@rotWithShape` changes from `1` to `0`, and a second projection reports `false`

#### Scenario: Stale or invalid rotation edit fails closed

- **WHEN** the source owner no longer contains the expected canonical token, the expected value is stale, or the requested value is not boolean
- **THEN** the edit fails without writing a broader effect rewrite
