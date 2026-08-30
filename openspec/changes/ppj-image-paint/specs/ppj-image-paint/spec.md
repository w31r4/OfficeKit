## ADDED Requirements

### Requirement: PPJ SHALL compile bounded native image paint
The system SHALL compile PPJ image fills for authored shapes and slide backgrounds and SHALL compile standalone image elements using stretch, deterministic cover or contain crop, explicit signed crop, direct opacity, or parameter-free tile mode without flattening the page.

#### Scenario: Authored image paint remains editable
- **WHEN** a valid source-free PPJ uses an image fill on a shape, a cropped translucent image background, and a tiled picture element
- **THEN** the compiler emits native DrawingML owners with content-addressed embedded assets and a second import recovers typed PPJ image-paint state

### Requirement: Image paint ownership SHALL remain distinct
The system SHALL preserve the distinction between native slide background paint, shape fill paint, and independently ordered picture elements.

#### Scenario: Background is not lowered to a picture layer
- **WHEN** a PPJ page declares an image background
- **THEN** the PPTX stores that asset in the page background owner and the page element z-order remains unchanged

### Requirement: Imported bounded image paint SHALL be projected and capability-bound
The system SHALL project a recognized embedded stretch or parameter-free tile blip fill with at most one signed source rectangle and one direct alpha into PPJ and SHALL issue only the operations supported by the exact source owner.

#### Scenario: Imported shape image fill changes locally
- **WHEN** an imported typed shape exposes `setFill` and the requested PPJ changes only its bounded image fill using a validated asset
- **THEN** the compiler re-proves the source and capability hashes, updates only that shape fill and necessary relationship state, and preserves unrelated parts and objects

#### Scenario: Imported page background changes locally
- **WHEN** an imported page exposes `setBackground` and the requested PPJ changes only its bounded background image paint
- **THEN** the compiler changes only the native page background and necessary relationship state without inserting a picture element

### Requirement: Unsupported image graphs SHALL remain source-owned
The system SHALL reject authoring or editing external image links, arbitrary tile transforms, multiple blip effects, artistic effects, unsupported color transforms, and custom vendor image topology.

#### Scenario: Unsupported imported paint is not normalized
- **WHEN** an imported shape or background contains an image graph outside the bounded profile
- **THEN** projection preserves it as opaque or read-only source state and an attempted semantic edit fails closed

### Requirement: Agent guidance SHALL expose executable image paint
The capability registry, generated PPJ reference, and focused Shapes and Media/Layers guidance SHALL describe the same authored and source-bound image-paint profile and its limits.

#### Scenario: Primitive is discoverable after compiler support lands
- **WHEN** the registry and generated PPJ documentation are checked
- **THEN** shape image fill, native image background crop/tile/opacity, and tiled image state are no longer listed as authored fail-closed gaps
