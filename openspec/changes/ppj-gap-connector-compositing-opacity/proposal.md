## Why

K-04 already exposes `compositing.opacity` in the PPJ schema and lowers it for
text, shape, image, and line paint, but a connector is rejected even though
its native line owner already has an alpha channel. This makes the declared
field inconsistent across PPJ's visual element set and leaves a concrete P0
gap in the authoring profile.

## What Changes

- Accept `connector.compositing.opacity` for authored PPJ programs.
- Treat it as a normal-opacity multiplier over the connector's existing stroke
  alpha, preserving the connector's single native line-alpha owner.
- Project the effective result through the existing canonical
  `connector.stroke.opacity` field.
- Keep non-normal blend modes, isolation, and clip stacks explicitly
  unsupported and fail closed.
- Add one focused compile, embedded-PPJ-free projection, and semantic
  round-trip experiment, then update the gap and coverage records.
- Do not change the Office wire version or add a new native dependency.

## Capabilities

### New Capabilities

- `ppj-compositing-opacity`: Bounded authored normal opacity for connectors,
  with canonical native line-alpha lowering and projection.

### Modified Capabilities

None.

## Impact

The change affects PPJ semantic validation, the authored NativeAOT compiler,
the existing connector projection contract, one focused codec test, and the
PPJ coverage/reference documents. The schema and protobuf contract already
contain the required field and remain unchanged.
