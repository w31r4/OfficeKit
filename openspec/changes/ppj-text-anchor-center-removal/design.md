## Context

See proposal.md. anchor_center is optional bool field 32, with no deletion marker. Current native Apply writes it only when present; semantic normalization cannot infer deletion from absence.

## Goals / Non-Goals

Preserve the existing optional boolean API and wire encoding, adding a distinct source-edit command. Full inherited anchor resolution and host reflow are separate gaps.

## Decisions

Add optional no_anchor_center at unused field 40. A selected marker must be true and cannot coexist with anchor_center (even false). This avoids moving the existing field into a oneof, which would change generated callers. Validate and admit that command, remove native anchorCtr, and clear the marker during semantic normalization. PPJ infers it only from old-present/new-absent source styles. Extend simple-style deletion; use forceAntiAlias for the remaining other-field rejection fixture.

## Risks / Trade-offs

Older codecs ignore unknown fields → the new operation requires the updated codec; no compatibility claim for old binaries. Existing PPJ compilation generates this marker inside that updated codec. Conflicting setter/deleter → validation rejects both, including false setter. Marker leaking into projected state → normalize it and verify compact table restoration. Regenerate JS with the repository command and check deterministic output.
