## Why

F-04 now exposes custom guides but still cannot express the adjustment list that precedes them in native geometry. Expose ordered custom adjustment formulas so PPJ can preserve an adjustment-to-guide-to-text-rectangle dependency chain.

## What Changes

- For kind custom, add geometry.adjustments as ordered name/formula pairs; keep kind preset's existing integer array unchanged.
- Preserve native adjustment order, formulas and guide references through authored export, fresh projection and source list edits/removal.
- Reuse final native graph validation, including cross-list duplicate names and references; retain unsupported handle/site and nonliteral path owners.

## Capabilities

### New Capabilities
- `ppj-custom-geometry-adjustments`: Editable ordered custom adjustment formulas in PPJ.

### Modified Capabilities
None; existing geometry increments remain unarchived.

## Impact

PPJ schema, authored/projected/source geometry, diagram kind-aware validation, one dependency-chain regression and Help/registry/reference/preview documentation. Existing native wire fields suffice.
