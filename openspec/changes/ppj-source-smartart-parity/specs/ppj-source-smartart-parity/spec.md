## ADDED Requirements

### Requirement: Proven SmartArt projects as typed PPJ state
An imported diagram SHALL project as `smartArt/source-bound` only when the
native codec has issued a closed editable diagram-text binding.

#### Scenario: Agent discovers editable diagram nodes
- **WHEN** the source DiagramDataPart has stable model IDs and recognized text
  topology
- **THEN** PPJ contains ordered typed nodes, exact run values, and bounded
  nativeRef authority

#### Scenario: Diagram is outside the safe profile
- **WHEN** the graph, relationships, identifiers, or text topology are not
  completely recognized
- **THEN** PPJ keeps the element opaque and issues no SmartArt text capability

### Requirement: Declarative node text editing
The complete PPJ `smartArt.nodes[]` state SHALL describe the requested text for
a capable source-bound diagram without describing its native graph.

#### Scenario: Agent changes one node run
- **WHEN** an Agent changes one existing PPJ run value and keeps node/run
  identity unchanged
- **THEN** build changes only the bound DiagramDataPart text token and second
  projection recovers the requested value

#### Scenario: Agent changes topology
- **WHEN** an Agent adds, removes, reorders, or restyles a source-bound node or
  run
- **THEN** build rejects the request before writing output

### Requirement: Native diagram graph remains source-owned
PPJ SmartArt text editing SHALL preserve layout, edges, colors, styles,
relationships, and every unrelated package part.

#### Scenario: One node changes
- **WHEN** a valid source-bound text edit succeeds
- **THEN** the slide XML, slide relationships, layout/style/color parts, and
  all other package parts remain byte-identical
