## Purpose

Make PPJ text anti-alias hints removable without conflating absence with false or changing independent source text properties.

## ADDED Requirements

### Requirement: Force anti-alias presence lifecycle

Supported text, shape, master/layout placeholder and table-cell owners SHALL preserve forceAntiAlias true/false and remove direct forceAA when its projected field is deleted. Other properties and source topology SHALL be preserved. Simple removable styles SHALL support whole removal and compact table restoration.

#### Scenario: Delete and restore booleans
- **WHEN** forceAntiAlias is removed and restored as false or true
- **THEN** native presence and fresh projection agree, with exact no-op and non-target XML/ZIP preservation

### Requirement: Valid deletion intent

The updated codec SHALL accept a true anti-alias deletion command without a setter, reject false commands and concurrent setters, and preserve existing setter encoding.

#### Scenario: Conflicting edit
- **WHEN** a native edit contains a setter and deletion command or a false deletion command
- **THEN** it is rejected

### Requirement: Honest preview boundary

Preview SHALL retain field-addressable limitations for explicit false. The hint SHALL NOT establish host rendering equivalence.

#### Scenario: Explicit false preview
- **WHEN** forceAntiAlias false is assessed
- **THEN** unsupported rendering semantics are not declared fully supported
