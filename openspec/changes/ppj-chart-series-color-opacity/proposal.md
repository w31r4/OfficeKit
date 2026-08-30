# PPJ chart-series color opacity

## Why

PPJ colors accept alpha, and chart series already own an alpha-aware solid-fill
profile, but the authored compiler rejects alpha only when the series uses the
compact `color` spelling. The equivalent structured `fill` spelling works.
This is an inconsistent language surface rather than a native limitation.

## What Changes

- Lower translucent series `color` through the existing solid chart-fill
  profile.
- Project it canonically as structured `fill` with opacity.
- Remove the obsolete compiler-boundary statement.

## Impact

- No schema, wire or package topology change.
- One existing PPJ integration contract proves the alias and canonical
  projection.

