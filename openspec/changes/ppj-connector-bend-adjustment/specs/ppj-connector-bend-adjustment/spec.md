## Purpose

Express a manual native bend adjustment in PPJ without losing its value, optional presence or independently bound connector state.

## ADDED Requirements

### Requirement: Literal bend adjustment survives authored and imported state

Elbow and curved connectors SHALL accept optional signed 32-bit integer bendAdjustment. Explicit zero and omission SHALL remain distinct through native export and fresh projection. Straight connectors with a supplied adjustment MUST be rejected. Unrecognized guide graphs or rotated bend axes not representable by endpoint normalization MUST remain opaque.

#### Scenario: Nondefault and absent bend
- **WHEN** an elbow connector supplies bendAdjustment 25000 and a curved connector omits it
- **THEN** native output and fresh PPJ SHALL retain the literal on the elbow and absence on the curve

### Requirement: Source bend edits preserve independent state

Issued setConnectorType authority SHALL permit bendAdjustment changes and removal. Source no-op SHALL preserve bytes. An adjustment edit SHALL preserve endpoints, bindings, arrows and non-target package members. Switching to straight SHALL clear an unchanged projected bend; an explicitly changed bend supplied with straight MUST reject.

#### Scenario: Zero and removal
- **WHEN** separate original-source requests set the adjustment to zero or remove it
- **THEN** fresh projection SHALL retain explicit zero or absence respectively, with only the target SlidePart changed

### Requirement: Preview reports an unpainted bend

Preview MUST NOT silently substitute a default midpoint route for a supplied nondefault bend.

#### Scenario: Internal nondefault bend
- **WHEN** the internal painter cannot map a nondefault adjustment
- **THEN** it SHALL emit a visible limitation and machine-readable diagnostic
