# Source-bound line opacity

## ADDED Requirements

### Requirement: expose only an existing direct line alpha

The projector MUST expose `lineOpacityThousandthPercent` only for a recognized
source-bound shape or connector whose direct outline has exactly one direct
solid RGB or theme paint and exactly one existing direct `a:alpha/@val` token
with a canonical integer from `0` through `100000`.

#### Scenario: shape and connector expose the alpha leaf

- **WHEN** a shape or connector has the bounded direct line profile and an
  existing alpha token
- **THEN** its native reference contains one `lineOpacityThousandthPercent`
  leaf with the exact canonical integer value
- **AND** the leaf identity distinguishes the owning element but does not
  expose a second semantic line owner

#### Scenario: missing or complex alpha remains source-owned

- **WHEN** the outline has no alpha token, more than one color/effect owner,
  an unsupported fill, an extension, or a malformed alpha
- **THEN** no `lineOpacityThousandthPercent` leaf is issued
- **AND** the source bytes remain available through the existing opaque or
  conservative projection boundary

### Requirement: edit only the proven alpha token

The source-bound compiler MUST accept a changed canonical opacity integer only
when the expected value still matches the hash-bound source and the requested
value is in `0..100000`.

#### Scenario: shape or connector alpha is edited and reprojected

- **WHEN** a valid native-leaf edit changes the existing line alpha
- **THEN** compilation replaces only that `a:alpha/@val` token in the owning
  SlidePart
- **AND** line paint, geometry, endpoint/style attributes, effects, and all
  non-owner parts remain unchanged
- **AND** Open XML validation and a second projection report the requested
  opacity

#### Scenario: stale, invalid, or unchanged edit is rejected

- **WHEN** the expected token is stale, the requested value is malformed or
  outside `0..100000`, or the requested value equals the current value
- **THEN** the edit fails closed without modifying the package
