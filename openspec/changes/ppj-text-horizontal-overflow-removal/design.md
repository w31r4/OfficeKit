## Context

Native NoHorizontalOverflowMode already removes horzOverflow and normalizes after edits. PPJ currently merges only present horizontalOverflow, and bounded admission excludes the deletion marker. Table body construction is separate.

## Goals / Non-Goals

Complete this direct field lifecycle across existing owners. Host clipping, inheritance resolution and other property deletion remain separate work.

## Decisions

Infer the existing marker from old-present/new-absent in both source paths, admit valid true markers, and extend the simple-style whitelist. Reuse the shared enum lifecycle experiment and table compact normalization. Substituting overflow for deletion would lose direct attribute absence.

## Risks / Trade-offs

Deleting a style could erase unrelated fields: retain the whitelist and capability checks. Removing the final table body property may compact its text: restore structured text with identical native topology. Assert both native values and absence, preserve verticalOverflow as an independent sibling, and compare all non-target ZIP bytes.
