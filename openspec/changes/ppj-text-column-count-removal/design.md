## Context

See proposal.md. Source merge handles present columns but does not infer deletion. The native NoColumns marker already removes numCol and is normalized before semantic comparison.

## Goals / Non-Goals

Complete direct count presence across supported text owners. Full inherited text layout and preview column painting remain separate gaps.

## Decisions

Infer NoColumns from old-present/new-absent styles in text and table paths, and admit its valid true case in bounded layout. Reuse the existing marker rather than changing the wire contract. Extend the guarded simple-style removal whitelist. Parameterize the numeric columnGap lifecycle regression to cover columns 1, 3 and 16 with preserved gap and direction.

## Risks / Trade-offs

Whole-style deletion could erase unrelated properties → retain the other-field and authority guards. Single-column default could mask absence → inspect native numCol and fresh projection separately. Table deletion can compact text → restore through the existing structured text path.
