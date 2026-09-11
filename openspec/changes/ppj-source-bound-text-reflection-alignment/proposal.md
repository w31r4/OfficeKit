## Why

Direct rich-text reflection alignment is already represented by DrawingML (`a:reflection/@algn`) and by the paragraph default-text PPJ profile, but an imported run currently drops this scalar from the typed PPJ projection. A narrow source-bound field closes this adjacent P0 reflection gap while keeping ambiguous effect graphs opaque.

## What Changes

- Add `run.style.reflection.alignment` with the nine DrawingML rectangle-alignment tokens.
- Project a strict direct-run native leaf `textReflectionAlignment` for one full-span reflection with one canonical `algn`.
- Allow source-bound edits to token-splice only that run-owned `reflection/@algn`, then reproject the changed token.
- Preserve unsupported transform combinations, malformed/duplicate attributes, and non-run owners as source-owned.
- Add one focused authored/source-bound/reprojection regression and update PPJ schema, capability metadata, docs, generated presentation references, and the OpenSpec record.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-reflection-alignment`: typed direct rich-text reflection alignment and its bounded source-bound edit.

### Modified Capabilities

None.

## Impact

The PPJ v1 schema, capability registry, native PPJ projector/compiler, native-leaf source-bound edit planner, presentation references, backlog/coverage records, and one OfficeKit codec test change. The wire protocol and host-PowerPoint acceptance boundary remain unchanged.
