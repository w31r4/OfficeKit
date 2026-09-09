## Context

See proposal.md. PptxGlowCodec already handles a direct glow followed by a proven shadow, but compiles only in Presentation. Chart text currently parses one shadow-only effect list using a strict XElement bridge. Existing ordinary glow JSON and native validation require a radius; this increment retains that field contract.

## Goals / Non-Goals

**Goals:** reuse glow model/validation and let known effects compose and delete independently across every shared chart style owner.

**Non-Goals:** arbitrary effects, native glow without a radius, renderer and PowerPoint visual acceptance.

## Decisions

- Move PptxGlowCodec compile ownership to Shared, as already done for shadow. A chart text effects adapter owns the whole recognized effect list and delegates typed color/geometry validation to the existing codecs. Retain strict unknown-node rejection around the typed reader.
- Add PresentationGlow as chart text wire field 20 without a wire-version change, reuse the existing glow schema, and include it in meaningful-style tests, semantic comparison and global-font-only restrictions.
- Share the chart effect color/opacity resolver with shadow so declared grammar tokens, theme fallback and alpha presence have identical semantics. Keep ordinary text effect mapping unchanged.
- Write one list ordered glow then outer shadow. Parse both into independent fields; deleting one reconstructs only the supported remaining sibling.
- Reuse line/combo style-owner lifecycle fixtures, native ordering/invalid-input checks, vector defaults/overrides and JS preservation. Keep the previous shadow regression in the focused gate.

## Risks / Trade-offs

- Accepting glow changes a previous negative fixture → replace that guard with an actually unknown or malformed graph, and add positive coexistence coverage.
- Shared-code movement or effect deletion can alter existing shadows → compile both consumers and retain ordinary glow/shadow and chart shadow regressions.
- Losing alpha presence or the other effect could pass a superficial visual check → inspect native attributes and fresh projection at each source-bound lifecycle step.
