# PPJ text-highlight specification

## ADDED Requirements

### Requirement: Authored text highlight

PPJ SHALL accept a bounded `highlight` color in run and default text styles and
compile it to canonical editable DrawingML text highlight state.

#### Scenario: Authored highlighted evidence

- **WHEN** a valid PPJ run declares a direct RGB highlight
- **THEN** build and re-import preserve the highlight without converting it to
  a shape, badge or background image

### Requirement: Safe imported projection

A canonical direct RGB run highlight SHALL project as typed PPJ state, while an
imported theme highlight SHALL remain available through its issued native leaf
until the source theme can be represented without semantic loss.

#### Scenario: Imported highlighted run

- **WHEN** a third-party slide contains one safe direct run highlight
- **THEN** PPJ inspection exposes the typed RGB value and its hash-bound native
  editing capability

### Requirement: Agent discoverability

The JSON Schema, generated PPJ manual and text guidance SHALL describe the same
highlight property and its editorial boundary.

#### Scenario: Maintainer check

- **WHEN** the Presentation Skill maintainer runs
- **THEN** the generated PPJ documentation and capability metadata remain in
  sync with the language schema
