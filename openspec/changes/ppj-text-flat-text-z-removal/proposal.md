## Why

F-03/F-06 expose `flatTextZ` but removing it from a source-bound PPJ style retains the original `a:flatTx`. Zero and absence must remain distinct.

## What Changes

- Add an explicit wire deletion intent without changing setter field 39.
- Remove canonical flat-text children through text, shape, master/layout placeholder and table-cell body styles; support restoration and simple style removal.
- Preserve unrelated XML and reject duplicate or noncanonical flat-text children.
- Extend the shared lifecycle regression and public field documentation.

## Capabilities

### New Capabilities
- `ppj-text-flat-text-z-removal`: Source-bound flat-text depth presence and removal.

### Modified Capabilities

None.

## Impact

Additive protobuf field, native body codec/compiler/normalization, shared PPJ tests and generated references. Updated codec required. Full 3D scene rendering remains outside this increment.
