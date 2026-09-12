## Purpose

Expose a bounded imported rich-text outer-shadow direction as a complete PPJ field so callers can change that one value while the source presentation remains structurally and byte-wise preserved elsewhere.

## ADDED Requirements

### Requirement: Direct rich-text shadow direction is projected as a bounded native field

The PPJ projection MUST expose `run.style.shadow.angle` as `textShadowDirectionDegrees` when the run owns exactly one direct outer shadow with an existing canonical non-negative `dir` token, one direct RGB or theme color child, no shadow transform attributes, no sibling effects, and no unknown descendants. The native leaf value MUST use the canonical 1/60000-degree integer and remain within 0 through 21,599,999.

#### Scenario: Canonical direct outer shadow exposes direction
- **WHEN** an imported text run contains one direct `a:effectLst/a:outerShdw` with a canonical `dir` token and the supported direct color topology
- **THEN** the projected run contains `nativeRef.leaves[]` with `kind` `textShadowDirectionDegrees` and the parsed native direction value

#### Scenario: Unsupported shadow topology stays source-owned
- **WHEN** the run contains a missing or malformed direction, a transform attribute, a sibling effect, an unknown descendant, or an out-of-range direction
- **THEN** the projection omits `textShadowDirectionDegrees` and keeps the source-owned effect graph opaque

### Requirement: Source-bound direction edits splice only the owning token

A source-bound edit of `textShadowDirectionDegrees` MUST require a changed canonical bounded integer value, verify the projected leaf's source proof, and replace only the existing `a:outerShdw/@dir` token in the owning SlidePart. It MUST preserve the run and paragraph topology, all other shadow attributes and children, and every unrelated OPC part.

#### Scenario: Direction edit preserves the package footprint
- **WHEN** a projected `textShadowDirectionDegrees` leaf is changed from one valid value to another
- **THEN** only the owning SlidePart is reported changed and its `dir` token contains the requested value while blur, distance, color, opacity, and other package parts remain unchanged

#### Scenario: Invalid or stale edits fail closed
- **WHEN** an edit supplies an out-of-range or non-canonical value, repeats the expected value, or its expected source token no longer matches
- **THEN** the operation fails without rewriting the source package

### Requirement: Edited direction reprojects to the typed PPJ field

After a successful source-bound edit, a subsequent PPJ projection MUST recover the new `run.style.shadow.angle` value and the corresponding `textShadowDirectionDegrees` native leaf without requiring host-specific rendering.

#### Scenario: Edited direction survives a second projection
- **WHEN** the edited PPTX is projected again
- **THEN** the run's typed shadow angle and native leaf both contain the edited 1/60000-degree value
