## Context

See proposal.md. The hint uses optional field 34 and native UseParagraphSpacing/spcFirstLastPara. Optional-boolean deletion now has validation, application and normalization precedents.

## Goals / Non-Goals

Complete direct hint presence while preserving actual paragraph spacing and text topology. Host paragraph metrics remain separate.

## Decisions

Add no_space_first_last_paragraph at unused field 42, preserving setter 34. Reject selected false and concurrent setters. Apply and normalize deletion using the established optional-boolean pattern, infer it from PPJ source presence, and extend simple-style removal. Reuse shared boolean/wire tests; move the remaining other-field guard to compatibleLineSpacing.

## Risks / Trade-offs

Hint deletion could affect actual paragraph spacing → include explicit paragraph spacing in the fixture and compare non-target XML. New field ignored by older codecs → document updated codec requirement. Shared preview registry changes → stage only owned reason changes. File semantics are not host-layout acceptance.
