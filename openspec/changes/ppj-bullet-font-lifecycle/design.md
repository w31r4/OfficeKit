## Context

See proposal.md. The wire already has family/follow-text oneof cases. `PptxBulletStyleCodec` reads only simple font choices; `ApplyChoice` currently replaces requested choices and leaves absent choices untouched. PPJ exposes only family. Ordinary paragraph style mutations use independent field masks and exact capability fields.

## Goals / Non-Goals

**Goals:** Complete this direct font choice using existing wire cases and the shared native style path. Exercise character/number/picture markers in one small lifecycle fixture per ordinary owner kind.

**Non-Goals:** Resolve inherited fonts, install fonts, widen unmodeled font metadata, switch marker kinds, implement host glyph metrics or change table/master editing authority.

## Decisions

- Expose `fontFollowText: true` alongside family and forbid both. A special family string would conflate a literal typeface with a native choice.
- Mask the two font properties together, then require authority separately for each property actually changed. Preserve existing marker and run topology checks.
- Copy the requested wire font choice only for changed paragraphs. Native font apply deletes a single modeled choice for absence, retains equivalent choices without reserialization, and switches only simple modeled declarations. Unknown attributes/children and duplicate choices stay opaque; don't rebuild them.
- Shared color/size apply compares modeled meaning before replacing an existing node. The lifecycle fixture reproduces lowercase RGB and signed/padded size spelling being rewritten during a font-only edit; unchanged choices must retain their original XML.
- Count valid Unicode scalars for the 255-character limit and reject invalid XML before serialization. Match schema and native validation, including supplementary characters.
- Reuse authored/source test helpers and compare ZIP bytes plus target XML with only the font choice scrubbed. Include raw native follow-text import, related list/table/master regressions and explicit preview diagnostics. Use the Skill maintainer for generated references.

## Risks / Trade-offs

- Shared native font deletion can affect table/list/master writers → run their existing focused regressions.
- Unknown metadata prevents semantic editing → preserve it through no-op/unrelated edits and reject replacement.
- Available fonts and inherited layout vary by host → retain partial/unavailable preview diagnostics; do not claim host acceptance.
