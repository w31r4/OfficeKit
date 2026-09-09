## Purpose

Let PPJ authors and source-bound editors express arrow dimensions without losing optional native attribute presence or independent connector state.

## ADDED Requirements

### Requirement: Arrow dimensions are optional typed fields

Connectors SHALL accept startArrowWidth/startArrowLength/endArrowWidth/endArrowLength with sm, med or lg values. Native export and fresh projection SHALL preserve explicit values and omission independently. Nonempty authored dimensions without their arrow MUST be rejected.

#### Scenario: Explicit and omitted dimensions
- **WHEN** a connector supplies three dimensions and omits the fourth
- **THEN** native output and fresh PPJ SHALL retain exactly those three attributes without materializing a default for the fourth

### Requirement: Dimensions support source-bound change and removal

setConnectorArrows SHALL authorize the four dimension fields. Changing or removing one SHALL preserve paired/opposite dimensions, endpoint bindings and non-target ZIP members. Removing an arrow SHALL clear its unchanged projected dimensions. Explicit nonempty dimension edits without that arrow MUST fail closed. Source no-op SHALL preserve bytes.

#### Scenario: Remove one size attribute
- **WHEN** an original-source request omits endArrowLength while retaining endArrowWidth
- **THEN** only the length attribute SHALL disappear and fresh projection SHALL retain the width and arrow type
