## Why

DrawingML already carries a direct rich-text reflection `rotWithShape` flag and PPJ can author it, but imported runs currently keep the flag outside the typed native-leaf surface. A narrow source-bound boolean field closes the next P0 text-reflection gap without widening ambiguous effect graphs.

## What Changes

- Add `run.style.reflection.rotateWithShape` as a direct-run native leaf.
- Project `textReflectionRotateWithShape` only for one full-span reflection with one explicit canonical `rotWithShape` token and no other reflection transform.
- Allow source-bound edits to replace only that run-owned `reflection/@rotWithShape` token, then reproject the changed token.
- Keep variable endpoints, other transforms, malformed/duplicate attributes, and unsupported effect topology source-owned.
- Add a focused authored/source-bound/reprojection regression and update PPJ capability evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-reflection-rotate-with-shape`: typed direct rich-text reflection rotation flag and its bounded source-bound edit.

### Modified Capabilities

None.

## Impact

The PPJ native-leaf registry and projector, direct-run reflection safety/proof and XML patching, presentation references, coverage/backlog records, generated capability matrix, and one OfficeKit codec test. The existing reflection wire model, authored compiler, and host-PowerPoint acceptance boundary remain unchanged.
