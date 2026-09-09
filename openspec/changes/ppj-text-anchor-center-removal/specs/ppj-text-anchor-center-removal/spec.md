## Purpose

Allow source-bound PPJ text to distinguish explicit anchor-center true or false from deleting the direct anchor-center override.

## ADDED Requirements

### Requirement: Anchor-center presence lifecycle

Supported text, shape, master/layout placeholder and table-cell owners SHALL preserve anchorCenter true/false and remove native anchorCtr when the projected field is deleted. Vertical alignment and other source state SHALL remain unchanged. Simple removable styles SHALL support whole deletion and compact table restoration.

#### Scenario: Delete and restore booleans
- **WHEN** projected anchorCenter is deleted and restored as false or true
- **THEN** native presence and fresh projection agree, source no-op bytes and non-target XML/ZIP state remain preserved

### Requirement: Unambiguous native deletion command

The updated codec SHALL accept an explicit true anchor-center deletion command without a setter and SHALL reject a false command or concurrent setter. Existing serialized boolean values SHALL retain their encoding and semantics.

#### Scenario: Conflicting command
- **WHEN** a native edit includes both a deletion command and an anchor-center value, including false
- **THEN** validation rejects the edit

### Requirement: Honest boundaries

Unsupported owners and other-property deletion SHALL retain existing guards. Preview SHALL retain field-addressable text-layout limitations for explicit anchorCenter false.

#### Scenario: Preview assessment
- **WHEN** explicit false is assessed for preview
- **THEN** it is not advertised as fully supported host layout
