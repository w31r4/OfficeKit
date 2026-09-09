## Purpose

Let PPJ connectors retain signed parent-space coordinates when endpoints lie outside a page or group origin, without sacrificing native range and source-fidelity checks.

## ADDED Requirements

### Requirement: Connector endpoint coordinates are signed

The compiler SHALL accept finite negative and positive from/to coordinates and object-anchor results within the DrawingML coordinate range, with extents within the positive-coordinate range. It SHALL preserve direction and EMU quantization without clamping to zero.

#### Scenario: Rotated group has an endpoint before its child origin
- **WHEN** a connector inside a transformed group targets an object that resolves to negative child-space coordinates
- **THEN** the actual connector and captured preview scene SHALL contain those signed coordinates and export successfully

### Requirement: Signed connectors round trip and remain editable

A recognized connector with negative native offsets SHALL project its signed endpoints or original object anchor. Source-bound endpoint edits SHALL preserve the non-target ZIP members and byte-exact no-op behavior. Direct placement capabilities SHALL retain their existing topology proofs while allowing signed connector offsets.

#### Scenario: Edit a literal negative endpoint
- **WHEN** a fresh source projection changes from.x from -20 to -40
- **THEN** only its SlidePart SHALL change and a fresh projection SHALL recover -40 while preserving the other endpoint

### Requirement: Invalid coordinate arithmetic fails closed

Endpoint coordinates SHALL stay within -27273042329600..27273042316900 EMU and each native extent SHALL stay within 0..27273042316900 EMU. Native offset/extent addition and rotation MUST NOT overflow or silently clamp. Unsupported imported transforms SHALL remain opaque with original bytes preserved.

#### Scenario: Malformed native transform is outside the range
- **WHEN** a native connector offset or extent exceeds its supported bounds
- **THEN** import SHALL preserve it as opaque rather than infer a wrapped endpoint
