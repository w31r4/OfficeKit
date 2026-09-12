## Purpose

Expose an explicit direct rich-text reflection `rotWithShape` flag in PPJ with a bounded native leaf and a source-preserving edit path for canonical DrawingML owners.

## ADDED Requirements

### Requirement: Direct run reflection rotation is typed

The PPJ v1 schema SHALL accept `run.style.reflection.rotateWithShape` as a boolean, and an imported direct rich-text run SHALL expose that value as `textReflectionRotateWithShape` only for a full-span direct reflection with one canonical `rotWithShape` token and no fade, scale, skew, or alignment attributes.

#### Scenario: Canonical rotation flag projects

- **WHEN** a run-owned `a:reflection` has `rotWithShape="1"`, full-span endpoints, and no other reflection transform
- **THEN** the run style contains `reflection.rotateWithShape = true` and one native leaf of kind `textReflectionRotateWithShape` with value `true`

#### Scenario: Explicit false remains visible

- **WHEN** a run-owned `a:reflection` has `rotWithShape="0"` and otherwise satisfies the strict profile
- **THEN** projection emits the same leaf with value `false` rather than treating the explicit token as omission

#### Scenario: Unsupported rotation topology stays opaque

- **WHEN** the run reflection also has a scale, skew, fade direction, alignment, variable endpoint, duplicate attribute, or unsupported effect sibling
- **THEN** the rotation field and native leaf are not issued for that source owner

### Requirement: Rotation flag edit is source-bound

A source-bound PPJ edit SHALL accept a changed canonical boolean token, prove the same run-owned reflection and leaf value, replace only `a:reflection/@rotWithShape` in the owning SlidePart, preserve unrelated package bytes, and make the new token visible on a subsequent projection.

#### Scenario: Canonical rotation token is replaced

- **WHEN** a `textReflectionRotateWithShape` leaf changes from `1` to `0`
- **THEN** only the owning SlidePart changes, the Open XML reflection has `rotWithShape="0"`, and the next PPJ projection reports `reflection.rotateWithShape = false`

#### Scenario: Invalid or stale edit is rejected

- **WHEN** the requested value is not a boolean, is unchanged, or the source attribute no longer matches the expected leaf value
- **THEN** compilation rejects the edit without rewriting unrelated source content
