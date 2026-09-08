## Why

The PPJ schema already permits `compositing.opacity` on every element base,
and the authored compiler already lowers icon and placeholder elements to
editable shapes. The semantic validator nevertheless rejects these two
shape-backed types, leaving normal opacity inconsistent for visible PPJ
elements.

## What Changes

- Accept `compositing.opacity` for authored `icon` and `placeholder` elements.
- Reuse the existing compound shape alpha lowering and keep the resulting
  native object editable as a shape-backed element.
- Preserve the existing explicit fail-closed behavior for blend modes,
  isolation, and clip stacks.
- Add one focused authored compile and snapshot-free projection experiment for
  both types, with no wire or schema version change.
- Update K-04 coverage and the PPJ reference.

## Capabilities

### New Capabilities

- `ppj-shape-backed-compositing-opacity`: Normal compositing opacity for
  authored icon and placeholder elements that own native shape paint.

### Modified Capabilities

None.

## Impact

The change affects PPJ semantic validation, one focused codec test, and the
coverage/reference records. The authored compiler already produces a shape
for both element types, so no protobuf or NativeAOT wire change is needed.
