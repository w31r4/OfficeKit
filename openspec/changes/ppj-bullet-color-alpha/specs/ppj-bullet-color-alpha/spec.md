# PPJ bullet color tokens and alpha

## ADDED Requirements

### Requirement: Bullet colors use the PPJ color catalog

OfficeKit SHALL compile bullet color literals and deck-local color tokens
through the same bounded RGB/alpha profile used by authored text paint.

#### Scenario: Build a token-colored translucent bullet

- **GIVEN** a valid PPJ paragraph bullet whose color references a deck token
  with alpha
- **WHEN** OfficeKit builds and projects the deck
- **THEN** the native bullet color retains the resolved RGB and alpha
- **AND** projected PPJ retains an equivalent alpha-bearing color

#### Scenario: Preserve unsupported transforms

- **GIVEN** an imported bullet color with transforms outside direct alpha
- **WHEN** OfficeKit imports the presentation
- **THEN** that bullet style remains source-owned
