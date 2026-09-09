## Purpose

Allow PPJ compatible line-spacing hints to be removed and restored independently of actual paragraph line spacing and surrounding text state.

## ADDED Requirements

### Requirement: Compatible line-spacing presence lifecycle

Supported text, shape, master/layout placeholder and table-cell owners SHALL distinguish compatibleLineSpacing true/false/absence. Removing the field SHALL remove native compatLnSpc and preserve paragraph spacing, other body properties and source topology. Simple removable styles SHALL support whole removal and compact table restoration.

#### Scenario: Delete and restore booleans
- **WHEN** the hint is deleted and restored as false or true
- **THEN** native state and fresh projection agree, with exact source no-op and non-target XML/ZIP preservation

### Requirement: Valid deletion intent

The updated codec SHALL accept true deletion without a setter, reject false deletion and concurrent setters, and preserve existing setter encoding.

#### Scenario: Invalid command
- **WHEN** a deletion command is false or accompanies a setter
- **THEN** validation rejects the edit

### Requirement: Preview boundary

Preview SHALL retain field-addressable layout limitations for explicit false and SHALL NOT claim host line metrics from this hint.

#### Scenario: False hint preview
- **WHEN** compatibleLineSpacing false is assessed
- **THEN** unsupported layout is not declared fully supported
