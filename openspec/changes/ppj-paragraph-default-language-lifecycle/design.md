## Context

See proposal.md. PptxLanguageTag accepts 2..63 characters matching a 2..8-letter initial subtag followed by hyphen-separated 1..8-character alphanumeric subtags. It preserves spelling and does not consult a language registry.

## Goals / Non-Goals

Complete direct paragraph lang presence for ordinary text/shape owners while keeping run language independent. Full language-registry validation, proofing and inherited placeholders remain separate.

## Decisions

Extend the exact field mask/authority and cloned default style. Patch lang only when its modeled value/presence changes, preserving altLang and all other default/run state. Reuse existing tag/token validation.

When a native lang value is outside the modeled grammar, reject replacing it through this new field rather than flattening source-owned state. Unrelated scalar edits still preserve it. Add that case alongside the shared valid-language lifecycle fixture.

## Risks / Trade-offs

Language can be confused with run overrides → retain explicit run language and compare native XML/ZIP state. Invalid source language is omitted from PPJ → reject replacement while permitting exact no-op and unrelated scalar edits. Keep the separately documented whole-default-style baseline test exclusion.
