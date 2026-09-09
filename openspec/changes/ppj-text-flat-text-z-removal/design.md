## Context

See proposal.md. The existing setter updates a canonical flatTx child and checks duplicates. Source-bound merge currently handles only present flatTextZ values.

## Goals / Non-Goals

Goals: distinguish signed 32-bit depth, explicit zero and absence for the existing five body-style owners.
Non-goals: full 3D rendering or inferred inheritance.

## Decisions

- Add optional no_flat_text_z field 45; retain optional int64 setter 39. Require true and no coexisting setter, following other body deletion intents instead of changing wire presence.
- Route old-present/new-absent styles into this intent, including table cells and simple whole-style removal; clear the intent during semantic normalization.
- Reuse canonical child validation for removal, preserving unrelated source XML. Removing only z would leave an unmodeled empty child, so remove the canonical flatTx child.
- Extend existing numeric lifecycle tests for five owners and two simple-style owners; cover signed limits, zero, restoration, wire conflicts and malformed native children.

- The boundary regression exposed nativeLeaf schema limits of -1e9..1e9. Add a kind-specific signed 32-bit integer branch for textBodyFlatTextZ while retaining existing limits for all other leaf kinds.

- FlatText is an SDK leaf: malformed nested markup is absent from ChildElements after parsing. Check preserved OuterXml before typed attribute access. Rejection tests assert the edit throws and the child is not removed; SDK lazy parsing itself can change serialization of malformed markup.

## Risks / Trade-offs

Updated codec required for the new intent. Unknown child content must fail closed; focused native tests cover this boundary without a host rendering acceptance run.
