## Why

Direct imported rich-text outer shadows already expose bounded blur, distance, direction, RGB color, and explicit opacity leaves. Their existing DrawingML alignment token remains source-owned, so a PPJ consumer cannot make one small, verifiable alignment change while preserving the rest of the run effect graph.

## What Changes

- Expose `run.style.shadow.alignment` as the `textShadowAlignment` native leaf.
- Project only one direct `a:outerShdw` with a valid existing `algn`, existing bounded geometry, one direct RGB or theme color, and no transforms, sibling effects, or unknown descendants.
- Validate and source-bind edits to the canonical `tl`, `t`, `tr`, `l`, `ctr`, `r`, `bl`, `b`, or `br` token, replacing only `outerShdw/@algn` in the owning SlidePart.
- Add a focused authored/source-bound/second-projection regression plus schema, registry, generated reference, coverage, and backlog entries.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-shadow-alignment`: Bounded projection and token-only source-bound editing of direct rich-text run outer-shadow alignment.

### Modified Capabilities

None.

## Impact

The PPJ native-leaf schema and capability registry gain one string leaf. The native text projection and edit-plan compiler gain a strict direct-run alignment path. Only the owning SlidePart is writable; all other package parts and unsupported effect graphs remain source-preserved. No Office host rendering claim is added.
