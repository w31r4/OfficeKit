## Context

See proposal.md. Native connector width/length strings already preserve sm/med/lg and absence. Authored PPJ currently sets only arrow type; projector drops sizes. setConnectorArrows already checks per-field authority and clears the dimensions of a removed end.

## Goals / Non-Goals

Goals: expose the four native size slots as optional PPJ fields with presence-aware edit/delete behavior.
Non-goals: point-based arrow contour geometry, arbitrary size values, other element owners or host exact metric claims.

## Decisions

- Use flat startArrowWidth/startArrowLength/endArrowWidth/endArrowLength to match native vocabulary and existing endpoint field names. sm/med/lg are relative native sizes, not points.
- Author only explicitly supplied sizes and let the native codec reject size state without an arrow. Project only nonempty native attributes.
- Extend setConnectorArrows fields. For each changed size, require field authority, set its string or clear on omission. A nonempty requested size needs a retained arrow.
- Process arrow removal before size edits: unchanged projected sizes are cleared with the removed end; an explicitly changed nonempty size on that removed end rejects. This keeps the existing remove-arrow operation valid while avoiding orphan state.
- Test all four authored attributes in one fixture, then independently change/remove dimensions from the same fresh original-source projection. Preserve paired dimensions, the opposite arrow and endpoint bindings; compare non-target ZIP bytes.

## Risks / Trade-offs

- Absence could become explicit medium → preserve empty native strings and omit absent PPJ properties.
- Removing arrow type can leave stale projected size fields → existing end removal clears them; only explicit conflicting edits reject.
- Native leaf and semantic size edits may conflict → existing mutation ownership checks remain in force.

## Migration Plan

Add optional schema fields and extend issued capability fields with no wire additions. Refresh source projections after upgrading. Native/JS checks and generated documentation cover this bounded increment; no full host suite is implied.
