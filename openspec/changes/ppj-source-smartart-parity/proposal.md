## Why

OfficeKit already discovers a bounded SmartArt diagram-data profile, exposes
stable node model IDs, validates direct text-run topology, splices changed text
into the original DiagramDataPart, and re-proves the result after export. PPJ
nevertheless projects every imported SmartArt frame as a generic opaque
element, so the public language hides a mature source-bound editing primitive.

## What Changes

- Project an imported diagram with a proven `PresentationDiagramText` binding
  as `type: "smartArt"`, `mode: "source-bound"`.
- Represent each editable native node as one stable PPJ node whose text keeps
  the proven run boundaries.
- Issue a bounded `replaceText/smartArt.nodes[].text` capability on the diagram
  and a node-local nativeRef for discovery and audit.
- Lower changed node text back into the existing opaque wire binding; keep
  layout, edges, style, colors, relationships, and all other diagram markup
  source-owned.
- Keep diagrams outside the proven profile opaque and unchanged.

## Capabilities

### New Capabilities

- `ppj-source-smartart-parity`: Capability-issued SmartArt node text editing in
  imported PPJ programs.

### Modified Capabilities

None. The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

PPJ projection, source-bound lowering, generated Skill guidance, capability
coverage, and one existing focused SmartArt contract are affected. The native
SmartArt graph reader/writer, Edit Plan codec, PPJ schema, and Office wire are
reused.
