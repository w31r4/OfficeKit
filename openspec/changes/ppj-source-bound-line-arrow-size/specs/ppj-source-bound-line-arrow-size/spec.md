# Source-bound line arrow size

## ADDED Requirements

### Requirement: expose existing explicit arrow size tokens

The projector MUST expose the four line arrow size leaves only when the
corresponding direct DrawingML endpoint contains one existing canonical
`sm|med|lg` width or length token and the endpoint is part of a recognized
shape or connector line profile.

#### Scenario: shape and connector expose head and tail size leaves

- **WHEN** a shape or connector has a direct `a:headEnd` or `a:tailEnd` with
  an explicit canonical `@w` or `@len`
- **THEN** the native reference contains the matching width or length leaf
- **AND** the leaf preserves the endpoint identity and does not add a second
  line owner

#### Scenario: absent or irregular endpoint size remains source-owned

- **WHEN** the endpoint has no explicit size, an invalid token, an unknown
  child, an extension, or an unsupported arrow graph
- **THEN** no corresponding arrow-size leaf is issued
- **AND** the source bytes remain available through the conservative boundary

### Requirement: edit only the selected endpoint attribute

The source-bound compiler MUST accept a changed arrow-size token only when the
expected token still matches the hash-bound source and the requested token is
one of `sm`, `med`, or `lg`.

#### Scenario: arrow size is edited and reprojected

- **WHEN** a valid native-leaf edit changes one head or tail width or length
- **THEN** compilation replaces only that endpoint attribute in the owning
  SlidePart
- **AND** the arrow type, opposite endpoint, line style, geometry, effects,
  and non-owner parts remain unchanged
- **AND** a second projection reports the requested token

#### Scenario: stale or unchanged edit is rejected

- **WHEN** the expected token is stale, the requested token is malformed or
  outside `sm|med|lg`, or it equals the current value
- **THEN** the edit fails closed without modifying the package
