## Purpose

Expose the vertical skew of a direct rich-text run outer shadow as a complete, bounded PPJ field while preserving unsupported source topology.

## ADDED Requirements

### Requirement: Direct rich-text shadow vertical skew is represented as a bounded field

The PPJ runtime MUST represent `run.style.shadow.skewY` as a number of degrees. Authored values MUST round to the native signed `a:outerShdw/@ky` integer at 1/60000-degree precision, and the rounded value MUST be strictly greater than -90 degrees and strictly less than 90 degrees. A projected native leaf MUST use kind `textShadowSkewY`, preserve the canonical signed `ky` token, and carry the same degree value.

#### Scenario: Authored shadow skewY survives native round trip

- **WHEN** an authored direct text run has a valid shadow with `skewY`
- **THEN** compilation writes `a:outerShdw/@ky`, and importing the result projects the same bounded `run.style.shadow.skewY` value and `textShadowSkewY` native leaf

#### Scenario: Invalid skewY is rejected

- **WHEN** authored `skewY` is non-numeric or rounds to a value at or beyond either strict 90-degree bound
- **THEN** compilation fails closed with a presentation-effects validation error

### Requirement: Source-bound skewY edits preserve the direct shadow owner

For a direct rich-text run whose effect list contains one direct, bounded `a:outerShdw` with the required geometry and one direct RGB or theme color child, the system MUST expose `textShadowSkewY` only when a canonical `ky` token is present and no `sx`, `sy`, `kx`, `rotWithShape`, sibling effect, or unknown descendant is present. A source-bound edit MUST replace only that `ky` token in the owning SlidePart, retain all other bytes and effect topology, and reproject the new degree value.

#### Scenario: A source-bound skewY edit changes only ky

- **WHEN** a projected direct run has `textShadowSkewY` and an edit changes its canonical signed token
- **THEN** only the owning SlidePart changes, the `a:outerShdw/@ky` token is replaced, all other shadow/XML/ZIP members remain intact, and a second projection returns the new `skewY`

#### Scenario: Unsupported shadow graphs remain opaque

- **WHEN** the run has a missing or invalid `ky`, another transform, an extra effect, an unknown descendant, or stale native proof
- **THEN** no `textShadowSkewY` edit capability is issued and the source remains unchanged
