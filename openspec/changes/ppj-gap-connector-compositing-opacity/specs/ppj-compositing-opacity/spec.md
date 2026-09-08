## Purpose

This capability gives authored PPJ connectors a truthful normal-opacity field
that lowers to PowerPoint's existing editable connector line alpha.

## ADDED Requirements

### Requirement: Connector compositing opacity is a bounded authored field

The PPJ compiler SHALL accept `connector.compositing.opacity` when the value
is a valid numeric value or an opacity grammar token. It SHALL apply the value
as a multiplier to the connector's existing stroke alpha, treating an absent
stroke alpha as one.

#### Scenario: Connector opacity lowers through the native line owner

- **WHEN** an authored connector has `compositing.opacity` set to `0.5` and a
  stroke opacity of `0.8`
- **THEN** the compiled connector SHALL carry an effective line alpha of `0.4`
  and remain a connector rather than being flattened to a shape or image

#### Scenario: Connector opacity accepts an opacity grammar token

- **WHEN** `compositing.opacity` references a declared grammar token of kind
  `opacity`
- **THEN** validation and authored compilation SHALL resolve that token and
  produce the same effective alpha as the equivalent numeric value

### Requirement: Connector projection has one canonical opacity owner

The PPJ projector SHALL report the effective native connector line alpha as
`connector.stroke.opacity`. It MUST NOT claim that an ordinary imported
connector has separately recoverable paint and element opacity values.

#### Scenario: Effective opacity reprojects without the authored snapshot

- **WHEN** a compiled connector is projected after its embedded PPJ snapshot
  is removed
- **THEN** the projected connector SHALL retain `type: "connector"` and report
  the effective alpha through `stroke.opacity`

### Requirement: Unsupported compositor semantics remain fail closed

Connector compositing SHALL retain the existing bounded profile: non-normal
blend modes, true isolation, and non-empty clip stacks MUST produce explicit
validation or compilation failure rather than a flattened approximation.

#### Scenario: Non-normal connector blend is rejected

- **WHEN** a connector declares a non-normal `compositing.blendMode`
- **THEN** PPJ validation or compilation SHALL reject the program with an
  explicit unsupported diagnostic
