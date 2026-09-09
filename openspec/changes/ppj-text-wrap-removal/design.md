## Context

PPJ wrap uses square/none. The native NoWrap marker deletes the attribute, whereas Wrap=none writes an explicit value. PPJ currently merges only present values; tables use a separate body builder.

## Goals / Non-Goals

Complete direct wrap presence across supported owners while preserving other content. Host automatic line-breaking, inheritance resolution and other property deletion are separate work.

## Decisions

Infer NoWrap for old-present/new-absent in both source paths and admit valid true markers. Extend the simple-style whitelist to wrap. Reuse table normalization and broaden the shared enum lifecycle experiment rather than copying test bodies. Assigning none for deletion would incorrectly materialize a default override.

## Risks / Trade-offs

The existing negative fixture includes wrap and margins; margins must continue to prevent whole-style deletion. Explicit none must survive fresh projection, unlike deletion. Removing a table's final body property may compact its text; restore through unchanged native paragraph/run topology and compare non-target ZIP bytes.
