## Purpose

Expose the distance of a safely recognized imported rich-text outer shadow as a complete PPJ native field while retaining every source-owned byte outside that one scalar.

## ADDED Requirements

### Requirement: Direct rich-text shadow distance is projected as a bounded native field

The PPJ projection MUST expose `run.style.shadow.distance` as `textShadowDistanceEmu` when the run owns exactly one direct outer shadow with an existing canonical non-negative `dist` token, one direct RGB or theme color child, no shadow transform attributes, no sibling effects, and no unknown descendants. The value MUST be represented in EMU and remain within the bounded distance range of 0 through 1,270,000,000 EMU.

#### Scenario: Canonical direct outer shadow exposes distance
- **WHEN** an imported text run contains one direct `a:effectLst/a:outerShdw` with a canonical `dist` token and the supported direct color topology
- **THEN** the projected run contains `nativeRef.leaves[]` with `kind` `textShadowDistanceEmu` and the parsed EMU value

#### Scenario: Unsupported shadow topology stays source-owned
- **WHEN** the run contains a missing or malformed distance, a transform attribute, a sibling effect, an unknown descendant, or an out-of-range distance
- **THEN** the projection omits `textShadowDistanceEmu` and keeps the source-owned effect graph opaque

### Requirement: Source-bound distance edits splice only the owning token

A source-bound edit of `textShadowDistanceEmu` MUST require a changed canonical bounded integer value, verify the projected leaf's source proof, and replace only the existing `a:outerShdw/@dist` token in the owning SlidePart. It MUST preserve the run and paragraph topology, all other shadow attributes and children, and every unrelated OPC part.

#### Scenario: Distance edit preserves the package footprint
- **WHEN** a projected `textShadowDistanceEmu` leaf is changed from one valid value to another
- **THEN** only the owning SlidePart is reported changed and its `dist` token contains the requested value while blur, direction, color, opacity, and other package parts remain unchanged

#### Scenario: Invalid or stale edits fail closed
- **WHEN** an edit supplies an out-of-range/non-canonical value, repeats the expected value, or its expected source token no longer matches
- **THEN** the operation fails without rewriting the source package

### Requirement: Edited distance reprojects to the typed PPJ field

After a successful source-bound edit, a subsequent PPJ projection MUST recover the new `run.style.shadow.distance` value and the corresponding `textShadowDistanceEmu` native leaf without requiring host-specific rendering.

#### Scenario: Edited distance survives a second projection
- **WHEN** the edited PPTX is projected again
- **THEN** the run's typed shadow distance and native leaf both contain the edited EMU value
