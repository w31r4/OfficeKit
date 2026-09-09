## Why

Ordinary PPJ connectors accept `from/to: {element, anchor}` but authored lowering ignores both fields and uses the connector's own frame. The two directed/moved-target regressions in the native preview integration expose this F-04 P0 defect in the actual compiled slide.

## What Changes

- Resolve the existing endpoint fields against final expanded object frames, including explicit sides/center, deterministic auto selection, rotation/flips and group child coordinates.
- Transform literal endpoints consistently when components expand, without applying component transforms twice to group children.
- Retain object-anchor identity through export/import and source-bound edits. Frame anchors and native geometry connection-site indexes are separate concepts; do not invent site indexes or claim host attachment without evidence.
- Add focused authored, movement, component/group and source-bound round-trip regressions. Keep the existing native preview counterexamples and their assertions.

## Capabilities

### New Capabilities

- `ppj-connector-object-anchors`: Resolve and preserve PPJ object-relative connector endpoints.

### Modified Capabilities

None. There are no published main specs in this checkout.

## Impact

The component expander, authored/source-bound presentation compiler, connector codec/projector and associated contract fields where needed. Update the F-04 backlog, capability record and focused shape reference. The preview painter continues to consume the compiler result. This change does not close full connector routing, geometry parity or the independent preview production-integration change.
