## Context

The PPJ repeat contract already owns `items`, direction, gap, grid/flow
columns, row gap, and anchor. The authored compiler parses this layout into a
repeat model and lowers each component item to an element transform; horizontal
and vertical paths currently allocate equal slots. See `proposal.md` and the
capability scenarios for the observable contract.

## Goals / Non-Goals

**Goals:**

- Give unequal authored stack slots one unambiguous owner: `layout.weights`.
- Preserve the current omitted-weight paths byte-for-byte at the contract level.
- Keep validation deterministic and fail closed for unsupported combinations.
- Prove the result through one small authored compile/project experiment.

**Non-Goals:**

- Do not add a new `stack` direction or a general constraint/solver graph.
- Do not apply weights to source-bound imported layout or infer weights from
  native PowerPoint geometry.
- Do not teach the JavaScript reviewer to reconstruct component definitions;
  after authored lowering it reviews ordinary projected frames.

## Decisions

### Use weights as a modifier of horizontal/vertical direction

`layout.weights` is explicit and orthogonal to the existing direction field.
With `horizontal` or `vertical` it selects weighted stack allocation; with no
weights the old equal allocation remains. A new `direction: "stack"` would
duplicate the axis and make migration/validation ambiguous, so it is rejected.

### Allocate the remaining main axis

For `n` items, subtract `gap * (n - 1)` from the instance width or height,
clamp the remaining extent to the existing positive minimum, then distribute
it by `weight / sum(weights)`. Each slot fills the cross axis as the existing
stack path does. Independent X/Y scale values are allowed because the field
describes slot geometry, not preservation of the component aspect ratio.

### Bound and validate at the PPJ boundary

The schema bounds the array and numeric values. The typed parser carries the
weights, while semantic validation checks direction, count, finiteness,
positivity, and incompatible `anchor`/grid/flow combinations. This gives
callers stable diagnostics before the expander can divide by an invalid sum.

### Keep projection ordinary

The authored compiler expands the weighted component into its existing native
element tree. PPTX projection therefore emits normal frames, and no protocol
change or source-bound capability is needed.

## Risks / Trade-offs

- [Risk] Very large or tiny positive weights can produce numerically narrow
  slots. → [Mitigation] Use bounded finite schema values and the existing
  positive available-axis clamp; the declared ratio remains deterministic.
- [Risk] A caller may expect aspect-ratio-preserving tiles. → [Mitigation]
  Document that weighted stack allocation fills each slot, while grid/flow
  retain their existing uniform behavior.
- [Risk] Future solver work may need richer constraints. → [Mitigation] Keep
  this field authored-only and leave cross-object constraints explicitly open
  in the backlog.

## Migration Plan

No migration is required. Existing repeat objects omit `weights` and retain
their current behavior. Authors can add the field only to horizontal or
vertical authored repeats; removing it restores equal-slot allocation.
