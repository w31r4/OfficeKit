## Context

See proposal.md. The codec already reads/writes simple a:cs nodes and validates nonblank names up to 255 characters. Source paragraph mutation currently authorizes Latin/East Asian fonts but not complex-script font presence.

## Goals / Non-Goals

Complete direct complex-script font presence for ordinary text/shape paragraphs. Preserve fixed topology, authored precedence, other script fonts and host substitution boundaries.

## Decisions

Extend the existing field-specific mask and authority checks. Copy only FontFamilyComplexScript presence into the cloned source default style. Unlike East Asian fonts, the authored builder only supplies this field when explicitly present, so no new fallback path is needed.

Extend the targeted default-style patch to call the existing complex-script font writer only when that field changes. Guard serialized leaf content before SDK typed access. Reuse the lifecycle and malformed-font fixtures for all three script-font nodes.

## Risks / Trade-offs

Whole-style writing can alter effects → patch only the changed font node. Extra native metadata/children can be flattened → reject unsupported replacement and verify unchanged input bytes. Keep the documented independent whole-default-style baseline failure exclusion.
