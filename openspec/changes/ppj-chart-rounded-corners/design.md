## Context

ChartSpace parsing and patching is shared by the bounded XLSX/PPTX chart codecs. Existing optional chart settings use presence-aware wire fields and direct XML token replacement.

## Goals / Non-Goals

**Goals:** Add one presence-aware boolean through the existing chart model, schema, compiler, projector, validator, capability registry, and ChartML patch path.

**Non-Goals:** 3-D chart support, extension namespaces, external workbook synchronization, or host rendering claims.

## Decisions

Represent presence separately from value so absent, explicit false, and true remain distinguishable. Place `c:roundedCorners` directly under `c:chartSpace`, validate `0|1|false|true`, and insert it before `c:chart` while preserving all unrelated XML.

## Risks / Trade-offs

[Risk] Some chart families or extension-wrapped nodes have ambiguous ownership. → Keep the existing editable-profile checks and fail closed.

[Risk] The field needs generated protobuf accessors. → Regenerate bindings with the repository's protocol check rather than hand-editing generated output.
