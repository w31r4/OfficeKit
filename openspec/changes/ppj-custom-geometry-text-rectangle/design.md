## Context

See proposal.md. Native rectangle edges already carry EMU literals or reference strings. The native codec validates resolved ordering and owns numeric scaling guides. PPJ custom paths use their own viewBox units; rectangle numeric edges instead belong to the shape-local point frame.

## Goals / Non-Goals

Goals: expose the existing rectangle state for literal custom shape paths, including built-in references and mixed edges. Non-goals: arbitrary custom guide/handle graphs, image masks, compositing clips or host text layout equivalence.

## Decisions

- Add geometry.textRectangle with required left/top/right/bottom. Numbers lower from points to EMU; strings pass to the native reference resolver. Preserve reference strings rather than resolving them into stale numeric coordinates.
- Allow rectangles through otherwise unchanged literal-custom projection checks. Numeric edges divide by 12700 on projection. Custom adjustment/guide/site/handle graphs remain source-owned.
- Extend setGeometry fields and check per-field authority. Rectangle add/change/removal shares the existing recognized custom geometry rewrite while keeping path, frame and text state.
- Shared geometry lowering must explicitly permit rectangle ownership for actual shapes; image masks and clips reject supplied rectangles instead of dropping them.
- Retain explicit preview diagnostics for unsupported text-layout geometry. One fixture covers mixed/all-reference/numeric state, source lifecycle and invalid bounds/references.

## Risks / Trade-offs

- ViewBox versus frame coordinates -> document numeric rectangle units as points independent of viewBox origin/extent.
- Numeric default replaces reference -> preserve union types and native presence.
- Unknown custom graph becomes editable -> keep all existing non-rectangle topology guards.

## Migration Plan

Additive PPJ schema/capability fields only. Refresh source projections to obtain authority; no wire or runtime release change is required for the source-library regression.
