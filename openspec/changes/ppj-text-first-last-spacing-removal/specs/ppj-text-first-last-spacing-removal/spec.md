## Purpose

Allow PPJ first/last paragraph-spacing hints to be removed and restored independently of actual paragraph spacing and other text properties.

## ADDED Requirements

### Requirement: Spacing hint presence lifecycle

Supported text, shape, master/layout placeholder and table-cell owners SHALL distinguish spaceFirstLastParagraph true/false/absence. Removing the projected field SHALL remove native spcFirstLastPara while preserving paragraph spacing, other body properties and source topology. Simple removable styles SHALL support whole deletion and compact table restoration.

#### Scenario: Delete and restore booleans
- **WHEN** the hint is removed and restored as false or true
- **THEN** native state and fresh projection agree, with exact source no-op and non-target XML/ZIP preservation

### Requirement: Unambiguous deletion intent

The updated codec SHALL accept true deletion without a setter, reject false deletion or concurrent setters, and retain existing setter encoding.

#### Scenario: Conflicting edit
- **WHEN** deletion is false or supplied together with a setter
- **THEN** validation rejects the edit

### Requirement: Preview limitations remain explicit

Preview SHALL retain field-addressable layout limitations for explicit false without claiming host paragraph metrics.

#### Scenario: False hint preview
- **WHEN** spaceFirstLastParagraph false is assessed
- **THEN** unsupported layout is not declared fully supported
