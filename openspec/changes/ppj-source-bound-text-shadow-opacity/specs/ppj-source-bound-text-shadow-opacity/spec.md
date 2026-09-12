## Purpose

Expose a bounded imported rich-text outer-shadow alpha as a complete PPJ field
so callers can change that one value while the source presentation remains
structurally and byte-wise preserved elsewhere.

## ADDED Requirements

### Requirement: Direct rich-text shadow opacity is projected as a bounded native field

The PPJ projection MUST expose `run.style.shadow.opacity` as
`textShadowOpacityThousandthPercent` when the run owns exactly one direct outer
shadow with one direct RGB or theme color child, one direct `alpha/@val` token
containing a canonical integer from 0 through 100000, at least one supported
bounded geometry attribute, no shadow transform attributes, no sibling effects,
and no unknown descendants. The native leaf MUST retain the canonical integer
token, and its typed value MUST be the corresponding ratio from 0 through 1.

#### Scenario: Canonical direct outer shadow exposes opacity

- **WHEN** an imported text run contains one direct
  `a:effectLst/a:outerShdw` with bounded blur, distance, or direction, one
  direct RGB or theme color child, and one direct alpha token
- **THEN** the projected run contains `nativeRef.leaves[]` with `kind`
  `textShadowOpacityThousandthPercent`, the native integer value, and
  `run.style.shadow.opacity` with the corresponding ratio

#### Scenario: Unsupported shadow opacity topology stays source-owned

- **WHEN** the run contains a missing or malformed alpha, missing supported
  geometry, a transform attribute, a sibling effect, an unknown descendant, or
  an out-of-range geometry or alpha token
- **THEN** the projection omits `textShadowOpacityThousandthPercent` and keeps
  the source-owned effect graph opaque

### Requirement: Source-bound opacity edits splice only the owning alpha token

A source-bound edit of `textShadowOpacityThousandthPercent` MUST require a
changed canonical integer from 0 through 100000, verify the projected leaf's
source proof, and replace only the existing direct
`a:outerShdw/(a:srgbClr|a:schemeClr)/a:alpha/@val` token in the owning
SlidePart. It MUST preserve the color child, all other shadow attributes and
children, run and paragraph topology, and every unrelated OPC part.

#### Scenario: Opacity edit preserves the package footprint

- **WHEN** a projected opacity leaf is changed from one valid value to another
- **THEN** only the owning SlidePart is reported changed and its alpha token
  contains the requested integer while color, geometry, and other package parts
  remain unchanged

#### Scenario: Invalid or stale edits fail closed

- **WHEN** an edit supplies a malformed, out-of-range, or unchanged opacity
  value, or its expected source token no longer matches
- **THEN** the operation fails without rewriting the source package

### Requirement: Edited opacity reprojects to the typed PPJ field

After a successful source-bound edit, a subsequent PPJ projection MUST recover
the new `run.style.shadow.opacity` and corresponding
`textShadowOpacityThousandthPercent` native leaf without requiring
host-specific rendering.

#### Scenario: Edited opacity survives a second projection

- **WHEN** the edited PPTX is projected again
- **THEN** the run's typed shadow opacity and native leaf both contain the edited
  value
