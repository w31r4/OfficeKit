## Purpose

Represent chart reflection alpha-gradient start and end positions without losing native omission, explicit default values or other effect state.

## ADDED Requirements

### Requirement: Optional alpha-ramp positions
Chart reflection SHALL accept optional startPosition/endPosition numbers in 0..1 and preserve their independent native presence at thousandth-percent precision. Values SHALL describe alpha-ramp positions, with no extra ordering constraint. Omitted positions SHALL remain absent in native serialization and projection; omitted whole reflection SHALL remove the effect.

#### Scenario: Direct native presence
- **WHEN** a chart reflection declares absent, equal, reversed or explicit 0/1 positions
- **THEN** its native attributes and fresh projection retain the selected state without synthesizing defaults

### Requirement: Chart lifecycle and style resolution
Positions SHALL participate in shared chart typography, named reflection-field precedence and vector defaults. Source-bound position edits SHALL preserve other reflection scalars, sibling effects, literal text and every non-target ZIP entry; no-op SHALL preserve exact source bytes.

#### Scenario: Line and combo source edits
- **WHEN** a fresh native line or combo projection changes, removes or recreates reflection position attributes
- **THEN** only the owning chart part changes and a second projection restores the new values and presence

### Requirement: Retain unrelated topology guards
The variable-position reader SHALL be used only for chart effect owners. Ordinary imported text/shape/image guards SHALL retain explicit full-span proof; extra chart reflection transforms/children and invalid positions SHALL remain source-owned or fail closed.

#### Scenario: Ordinary and invalid chart owners
- **WHEN** an ordinary imported reflection has non-full-span positions or a chart position exceeds 0..1 or contains an unknown child
- **THEN** unsupported editing remains rejected and source no-op retains original bytes
