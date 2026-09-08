## Why

Authored component repeats can currently express equal horizontal or vertical
slots, but the backlog's `stack` gap still has no explicit weight owner. Add a
bounded weight field so a repeat can state deterministic unequal slot sizes and
be exercised as ordinary PPJ frames.

## What Changes

- Add optional `repeat.layout.weights` for authored horizontal or vertical
  repeats.
- Define positive, finite, item-count-matched weights; preserve equal-repeat
  behavior when the field is omitted.
- Compile weighted repeats with the existing gap and frame projection path.
- Reject incompatible layout combinations instead of silently ignoring the
  weights.
- Keep weighted repeats authored-only; do not add solver constraints or
  source-bound layout editing.

## Capabilities

### New Capabilities

- `ppj-stack-repeat-weight`: Explicit weighted horizontal/vertical component
  repeat allocation with bounded validation and deterministic frame output.

### Modified Capabilities

None.

## Impact

- PPJ schema/reference, repeat parsing and semantic validation, and authored
  component expansion.
- One focused authored compile/project regression covers a `[1,2,1]` stack and
  invalid weight rejection.
- No protobuf or Office wire-version change; grid/flow/anchor and source-bound
  layout behavior remain unchanged.
