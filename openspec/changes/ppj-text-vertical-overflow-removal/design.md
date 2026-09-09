## Context

Native NoVerticalOverflowMode removes vertOverflow and normalizes after edits. PPJ merge currently handles only present verticalOverflow, and bounded admission excludes the deletion marker. Tables build requested bodies separately.

## Goals / Non-Goals

Complete this direct field across existing owners. Host ellipsis/clipping, inheritance resolution and other property deletion remain separate work.

## Decisions

Infer the existing marker from old-present/new-absent in both source paths, admit valid true markers and extend the simple-style whitelist. Reuse the shared enum experiment and table compact normalization. Assigning overflow instead of deleting would lose native absence.

## Risks / Trade-offs

Whole-style deletion must retain other-field and authority guards. The experiment preserves an independent horizontalOverflow and checks overflow/ellipsis/clip against native enums. Table restoration must retain native paragraph/run topology; surrounding XML and non-target ZIP comparisons catch collateral edits.
