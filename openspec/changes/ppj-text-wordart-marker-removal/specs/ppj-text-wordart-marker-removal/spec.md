## Purpose

Allow PPJ WordArt source markers to be removed and restored independently of text warp and other source text properties.

## ADDED Requirements

### Requirement: WordArt marker presence lifecycle

Supported text, shape, master/layout placeholder and table-cell owners SHALL preserve fromWordArt true/false and remove its native attribute when the projected field is deleted. Text warp and other XML/ZIP state SHALL be preserved. Simple removable styles SHALL support whole removal and compact table restoration.

#### Scenario: Delete and restore with text warp retained
- **WHEN** the marker is removed and restored as false or true
- **THEN** native presence and fresh projection agree, source no-op is byte-identical, and text warp remains unchanged

### Requirement: Valid deletion intent

The updated codec SHALL accept true deletion without a setter, reject false deletion and concurrent setters, and retain existing setter encoding.

#### Scenario: Conflicting command
- **WHEN** deletion is false or accompanies a setter
- **THEN** validation rejects the command

### Requirement: Preview boundary

Preview SHALL retain field-addressable limitations for explicit false and SHALL NOT claim full WordArt rendering from this marker.

#### Scenario: False marker preview
- **WHEN** fromWordArt false is assessed
- **THEN** unsupported rendering is not declared fully supported
