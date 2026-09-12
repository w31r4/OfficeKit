## Purpose

This capability gives PPJ a small, source-preserving way to express and edit the alignment of an imported rich-text run outer shadow without flattening the surrounding text effects or package topology.

## ADDED Requirements

### Requirement: Direct rich-text outer-shadow alignment is a bounded native leaf

The system SHALL expose `run.style.shadow.alignment` as `textShadowAlignment` only when a text run owns exactly one direct `a:effectLst/a:outerShdw` with one existing canonical `algn` token, at least one existing bounded blur, distance, or direction token, one direct RGB or theme color child, no scale/skew/rotation transforms, no sibling effects, and no unknown descendants. The native leaf value SHALL be one of `tl`, `t`, `tr`, `l`, `ctr`, `r`, `bl`, `b`, or `br`.

#### Scenario: Project a valid direct alignment

- **WHEN** an imported run has a strict direct outer shadow with `algn="ctr"`, bounded geometry, and one RGB or theme color child
- **THEN** the fresh PPJ projection SHALL include `textShadowAlignment` with value `ctr` and the typed run shadow SHALL report alignment `ctr`

#### Scenario: Edit only the alignment token

- **WHEN** a source-bound edit changes a proven `textShadowAlignment` value from `ctr` to `br`
- **THEN** the output SHALL replace only that `outerShdw/@algn` token in the owning SlidePart, preserve the run text, geometry, color, opacity, other effects, and non-target package parts, and project `br` on the next import

#### Scenario: Preserve theme-color ownership

- **WHEN** the strict outer shadow uses one direct `schemeClr` child with a valid alignment token
- **THEN** the alignment leaf MAY be issued while the RGB color leaf remains absent, and editing alignment SHALL preserve the theme-color child and its identity

#### Scenario: Reject unsupported alignment graphs

- **WHEN** alignment is missing or invalid, geometry is missing or out of bounds, a scale/skew/rotation transform or sibling/unknown descendant is present, or the edit proof is stale or requests an invalid/no-op token
- **THEN** the projection SHALL omit the alignment leaf or the edit SHALL fail closed without rewriting the source package
