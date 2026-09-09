## Why

F-04 connector startArrow/endArrow already author and project, but ApplyConnectorElement rejects their source-bound edits. Agents must currently discover native leaves to change a field already present in PPJ.

## What Changes

- Issue setConnectorArrows for editable connectors, with explicit startArrow/endArrow fields.
- Support changing, adding or omitting either arrow; omission means none, preserving the opposite arrow and size attributes on retained arrows; deleting an arrow removes its own size state.
- Preserve endpoint coordinates, frame-anchor/native-site bindings and non-target ZIP parts through fresh-source edits.

## Capabilities

### New Capabilities
- `ppj-connector-arrow-edits`: Source-bound lifecycle for PPJ connector arrow fields.

### Modified Capabilities
None; no published main spec owns this capability.

## Impact

PPJ schema capability vocabulary, semantic validator, projector/compiler, a focused connector regression and discoverability docs. Reuse existing native arrow fields and codec; no wire additions or routing change.
