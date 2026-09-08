## Purpose

This capability makes the existing normal-opacity PPJ field usable for icon
and placeholder elements whose authored native representation is a shape.

## ADDED Requirements

### Requirement: Shape-backed elements accept normal compositing opacity

The PPJ compiler SHALL accept a valid numeric or `opacity` grammar-token value
at `compositing.opacity` for authored `icon` and `placeholder` elements. The
value SHALL be applied to the same directly owned shape paint branches used by
ordinary shape/text opacity.

#### Scenario: Icon compositing opacity is authored

- **WHEN** an authored icon declares `compositing.opacity` as `0.5`
- **THEN** the output SHALL contain the icon's editable native shape with its
  directly owned fill alpha multiplied by `0.5`

#### Scenario: Placeholder compositing opacity is authored

- **WHEN** an authored placeholder declares `compositing.opacity` as `0.5`
- **THEN** the output SHALL contain the placeholder's shape-backed text owner
  with supported directly owned paint alpha multiplied by `0.5`

### Requirement: Shape-backed projection preserves the existing type boundary

Removing the embedded authored PPJ snapshot and projecting the generated
package SHALL preserve the existing native projection behavior. It MUST NOT
claim that an arbitrary imported shape is an icon or recover an authored
compositing declaration that has no separate native owner.

#### Scenario: Icon projection remains a shape-backed native object

- **WHEN** an authored icon package is projected without its embedded PPJ
- **THEN** the projected element SHALL remain editable shape-backed content and
  report the effective paint opacity through its existing shape fields

#### Scenario: Unsupported compositor semantics remain rejected

- **WHEN** an icon or placeholder declares non-normal blend, true isolation, or
  a non-empty clip stack
- **THEN** validation or compilation SHALL reject it with the existing explicit
  unsupported diagnostic
