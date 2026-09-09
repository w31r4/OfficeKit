## Context

The compiler currently merges only present columnDirection values. Native NoColumnDirection already removes rtlCol, but bounded admission excludes this marker. Tables build body styles separately.

## Goals / Non-Goals

Complete the existing field lifecycle across supported owners. Full host column layout and unrelated body-property deletions remain outside this increment.

## Decisions

Infer NoColumnDirection from old-present/new-absent PPJ styles in both merge paths and admit a valid true marker. Extend the removable-style whitelist to columnDirection/upright/rotation. Reuse existing table compact normalization and capability guards. A new wire field or a default left-to-right assignment would unnecessarily change the contract or lose absence.

## Risks / Trade-offs

Explicit left-to-right is native false, not absence: assert attribute presence after each restoration. Deleting the final property can compact table text: test reconstruction with unchanged paragraph/run topology. Compare surrounding XML and all non-target ZIP entries to catch collateral edits.
