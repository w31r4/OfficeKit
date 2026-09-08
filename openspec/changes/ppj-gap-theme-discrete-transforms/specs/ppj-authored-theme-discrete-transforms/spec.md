# PPJ authored discrete theme color transforms

## ADDED Requirements

### Requirement: expose direct discrete accent color operations

The PPJ schema MUST accept optional explicit `true` values for `gray`, `comp`,
`inv`, `gamma`, and `invGamma` on each `accent1` through `accent6` transform.
Authored compilation MUST lower each present operation to its matching direct
DrawingML color-transform child without changing the base accent role or
unrelated transforms.

#### Scenario: authored discrete transforms survive native output

- **WHEN** `accent5` declares `gray: true`, `comp: true`, `inv: true`,
  `gamma: true`, and `invGamma: true`
- **THEN** the authored theme contains the matching empty
  `a:gray`/`a:comp`/`a:inv`/`a:gamma`/`a:invGamma` children below
  `a:accent5Color/a:srgbClr`
- **AND** embedded PPJ projection recovers all five declarations

### Requirement: preserve discrete transform presence

When a discrete operation is omitted, the compiler MUST NOT synthesize its
DrawingML child. Other accent transform fields MUST remain independent.

#### Scenario: omitted operation remains absent

- **WHEN** `accent1` declares only `gray: true`
- **THEN** compilation emits `a:gray` but no `a:comp`, `a:inv`, `a:gamma`, or
  `a:invGamma`

### Requirement: reject unsafe discrete transforms

The compiler MUST reject false-valued discrete transform fields and empty
transform objects before native theme output.

#### Scenario: false discrete transform fails closed

- **WHEN** a role declares `inv: false`
- **THEN** PPJ validation or compilation fails
