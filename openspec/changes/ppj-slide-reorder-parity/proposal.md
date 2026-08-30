## Why

The native PPTX codec already reorders retained SlideParts without rebuilding
their content or relationships, and PPJ already represents presentation order
through the `pages` array. Source projection still derives page and element IDs
from the current slide position, however, so a reorder would change the public
identities after reimport. PPJ therefore hides a mature native primitive and
cannot yet treat an imported deck as stable program state across page moves.

## What Changes

- Derive imported page IDs from the stable SlidePart path instead of slide
  position.
- Derive page-local element IDs without the position-derived slide owner
  prefix, keeping them stable when the whole page moves.
- Add `pageOrder` to the existing bounded `reorder` capability.
- Issue page reorder only for source-bound pages whose section graph is absent
  or modeled; opaque section graphs remain fail-closed.
- Lower a complete retained-page permutation to the existing source-preserving
  slide reorder writer and allow modeled sections/custom shows/comments to
  retain their stable page references.
- Extend the existing comprehensive PPJ contract and generated guidance.

## Capabilities

### New Capabilities

- `ppj-slide-reorder-parity`: Stable source-projected page identity and bounded
  declarative page-order editing.

### Modified Capabilities

None. The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

PPJ schema vocabulary, projection identity, semantic validation,
source-bound lowering, generated guidance, coverage, and one existing test are
affected. No protobuf or OOXML writer change is required.
