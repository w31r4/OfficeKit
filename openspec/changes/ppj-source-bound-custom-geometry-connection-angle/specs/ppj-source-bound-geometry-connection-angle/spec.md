## Purpose

This capability gives a recognized custom geometry one independently editable
connection-site angle while keeping the rest of its DrawingML graph opaque and
source-owned.

## ADDED Requirements

### Requirement: Expose bounded literal connection-site angles as native leaves

The system SHALL expose `customGeometryConnectionSiteAngle60000` only for an
existing, ordered, direct custom-geometry connection site whose `ang` value is
a canonical bounded signed integer in DrawingML 60,000ths-of-a-degree units.
The native index SHALL identify the connection site in its direct `cxnLst`
order, and non-literal or otherwise unsupported sites SHALL remain
source-owned.

#### Scenario: Project one literal connection-site angle without flattening geometry

- **WHEN** a source-bound custom geometry contains a direct connection site with a bounded literal `ang` token and its position is within the recognized connection-site profile
- **THEN** PPJ SHALL issue one native leaf with the ordered site index and numeric angle value while preserving the custom geometry graph and all other connection-site fields

#### Scenario: Edit only the selected connection-site angle

- **WHEN** a source-bound PPJ program changes the issued connection-site angle leaf while retaining its native identity and preconditions
- **THEN** compilation SHALL change only the selected `a:cxn/@ang` token in the owning SlidePart, preserve the position and connection-site order, and reproject the new angle

#### Scenario: Keep calculated, malformed, or stale connection-site state source-owned

- **WHEN** the selected angle is formula-backed, non-canonical, out of bounds, structurally ambiguous, or no longer matches the bound source
- **THEN** PPJ SHALL issue no editable leaf or compilation SHALL fail closed before writing any package part
