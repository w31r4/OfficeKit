## Context

The existing optional wire string models five native alignment tokens. `PptxParagraphPropertiesCodec.Apply` already removes known native alignment when the complete paragraph state omits it, but assigns any supplied alignment over unknown source values. Source-bound lowering currently copies compiled alignment to all paragraphs after any alignment change. The initial fixture also reproduced projection failing with `codec_failure` on invalid native alignment; authored precedence checks must strip embedded PPJ to inspect fresh native projection.

## Goals / Non-Goals

**Goals:** Complete the existing direct field with raw-token classification and owner-local preservation.

**Non-Goals:** Additional alignment modes, inherited placeholder graph edits, NativeAOT publication and host font/layout acceptance.

## Decisions

- Read the raw unqualified `algn` attribute using a closed token mapping. Typed SDK enum access is unnecessary for deciding whether source content is modeled.
- Reject assignment over an unmodeled native attribute. Skip native assignment when the mapped value is unchanged, and scrub only known values. Preserve the existing omission-as-removal contract instead of adding a redundant wire marker.
- Compare the individual paragraph's direct alignment before lowering it. Reuse exact field authority and the full-paragraph mutation path.
- Reuse the lifecycle test helpers with text/shape owners, one neighboring paragraph and unrelated XML. Cover all five values plus unknown native tokens; keep checks focused on affected paragraph readers/writers.

## Risks / Trade-offs

- Shared paragraph properties also serve lists and other text consumers → run their focused existing regressions.
- An unsupported but valid native alignment remains read-only → retain it and reject replacement explicitly.
- Preview glyph layout is approximate → retain its partial diagnostics and avoid treating structural tests as host acceptance.
