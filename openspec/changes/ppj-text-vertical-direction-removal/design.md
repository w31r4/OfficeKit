## Context

verticalText uses the existing horizontal/vertical/vertical270 enum. The native codec already applies NoVerticalTextMode and normalizes that edit marker, but PPJ merge ignores omissions and bounded admission excludes the marker. Tables construct requested bodies separately.

## Goals / Non-Goals

Complete direct verticalText presence under existing authority. Full inherited text layout and other body-property deletions remain separate work.

## Decisions

Infer the existing deletion marker from old-present/new-absent in both source paths. Admit only the valid true marker and extend the simple-style removal whitelist. Explicit horizontal remains a value; assigning horizontal instead of deleting would lose native absence. Reuse existing table compact normalization and the shared lifecycle fixture, with direct native enum checks and fresh projections.

## Risks / Trade-offs

Whole-style deletion could erase unrelated fields: keep the whitelist and capability guard. Table normalization can change PPJ text shape: restore structured text with identical native paragraph/run topology. XML comparisons must preserve every non-target value while tolerating redundant namespace declarations already produced by the table writer.
