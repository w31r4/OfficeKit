## Purpose

Expose ordered custom-shape connection sites in PPJ while retaining native coordinate references and connector index identity.

## ADDED Requirements

### Requirement: Sites retain typed values and order

Custom shape geometry SHALL accept optional connectionSites with at most 1024 ordered records requiring angle, x and y. Values SHALL be numeric degrees/local points or native built-in/declared adjustment/guide references. Authored export and fresh projection SHALL preserve reference identity, numeric values and order. Empty lists SHALL project as omission. Unresolved references, out-of-frame coordinates and angles outside one signed turn MUST reject. Masks/clips MUST reject shape connection sites.

#### Scenario: Mixed site records
- **WHEN** sites combine literal values with built-in and declared references
- **THEN** native export and fresh projection SHALL preserve the ordered records

### Requirement: Source edits preserve site index identity

Issued setGeometry authority SHALL allow changing geometry.connectionSites values at existing indexes. Source list length changes MUST reject because indexes are native connector identity. Source no-op SHALL preserve bytes. Edits SHALL retain unchanged geometry, text, frame and non-target package members.

#### Scenario: Edit a site and reject removal
- **WHEN** a source request changes a site value or removes a nonempty list
- **THEN** the value change SHALL reproject correctly and list removal SHALL reject

### Requirement: Preview reports site limitations

Preview SHALL explicitly diagnose connection-site semantics it cannot render.

#### Scenario: Site preview
- **WHEN** a custom shape declares connection sites
- **THEN** unsupported site fields SHALL NOT be reported as fully rendered
