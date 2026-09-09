## Purpose

Let an Agent edit existing PPJ connector arrow fields through the same source-bound language contract used for authored and projected presentations.

## ADDED Requirements

### Requirement: Editable connectors expose arrow field authority

A recognized editable connector SHALL expose setConnectorArrows for startArrow and endArrow. A source edit lacking this authority MUST fail closed without emitting a candidate file.

#### Scenario: Missing capability
- **WHEN** a request changes an arrow but lacks its source-issued capability
- **THEN** compilation SHALL report capability rejection and return no candidate file

### Requirement: Arrow edits preserve independent state

Changing or adding an arrow SHALL preserve the other end, endpoint coordinates/bindings, and dimensions on retained arrows. Omitting an arrow field or setting none SHALL remove that arrow and its own dimensions. No-op SHALL preserve source bytes, and supported edits SHALL change only the target SlidePart and survive fresh projection.

#### Scenario: Change and then independently delete a source arrow
- **WHEN** separate requests from the same original projection change endArrow from triangle to diamond or omit it
- **THEN** the changed arrow SHALL retain its dimensions, the removed arrow SHALL have no orphan dimensions, and both candidates SHALL retain their start end and anchor bindings
