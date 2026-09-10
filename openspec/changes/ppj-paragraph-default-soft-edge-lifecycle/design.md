## Context

See proposal.md. The schema, authored builder and paragraph projector already carry `{ radius }`. Default-run reading isolates direct soft edges, but the source compiler has no soft-edge authority and the default-run writer still treats changes as a whole-style rewrite.

## Goals / Non-Goals

Complete the existing direct paragraph owner and preserve unknown source state. Inherited defaults, table paragraphs and host rendering retain their separate boundaries.

## Decisions

Reuse the existing radius schema, builder and isolated reader. Add independent change masking and exact field authority to the source compiler. Copy only changed paragraph soft-edge values into the edit model.

Use the shared direct-effect splice helper to create, replace and delete only `a:softEdge`; include it in modeled-state clearing. A whole-list rewrite would disturb sibling XML and cannot preserve an attribute-bearing empty list. Existing validation rejects unmodeled source owners before replacement.

The mixed-source experiment requires inserting a restored soft edge after the last known effect: the SDK's generic insertion can put it before glow when an unknown sibling exists. The invalid-radius experiment also shows the SDK reading a negative unsigned radius as zero; validate the raw radius text in the shared value reader before accepting it.

Extend the existing original-source XML/ZIP harness with small lifecycle and rejected-source fixtures. Change the old generic unsupported-edit fixture to paragraph spacing, because softEdge now becomes authorized.

## Risks / Trade-offs

Zero is confused with absence → assert native effect presence and fresh PPJ after assignment/removal/restoration.

Malformed native radius or nested content is normalized away → retain rejection and unrelated-edit preservation fixtures.

Passing structural tests is read as visual acceptance → keep explicit partial preview evidence and record unperformed NativeAOT/host checks.
