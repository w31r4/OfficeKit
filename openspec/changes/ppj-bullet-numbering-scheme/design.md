## Context

See proposal.md. `PptxBulletCodec` has a 41-token native allowlist; authored PPJ already resolves five aliases with `NumberScheme`. Projection emits the native scheme. The source compiler has independent paragraph masks, including startAt, but rejects scheme/format changes. Native application preserves a same-scheme automatic marker and rebuilds one when the scheme changes.

## Goals / Non-Goals

**Goals:** Preserve the existing automatic marker while editing either spelling of its numbering format; compose that edit with separately authorized startAt changes.

**Non-Goals:** Converting character/picture/no-bullet markers, implicit format defaults, automatic list evaluation or broader inheritance.

## Decisions

- Close the schema vocabulary over the codec's existing catalog; retain the existing oneOf between scheme and format. A missing required selection fails rather than inventing a default.
- Mask scheme and format together as two spellings of one state. Per changed paragraph, require the exact field matching the requested spelling; requests containing different spellings on different paragraphs require both capabilities. Reuse authored alias lowering instead of adding another resolver.
- Lower format before startAt and modify only each changed wire field. Native automatic-marker application updates Type and optional StartAt independently on the original element, retaining unknown XML and unchanged lexical attributes. Whole-marker replacement would lose that source content.
- Use one authored catalog fixture and the existing lifecycle helper for text/shape source edits, aliases, restoration, combined edits and negative authority. Keep unsupported preview diagnostics instead of expanding rendering work.

## Risks / Trade-offs

- Alias-only syntax changes can preserve native meaning → require the alias capability and verify equivalent aliases preserve native bytes.
- Shared native bullet application serves other owners → run the existing bullet/list/master regressions alongside the focused cases.
- A combined format/startAt edit can overwrite another field → update only changed fields and independently check both authorities and target attributes.
- Catalog discoverability does not establish list layout → retain explicit partial preview and inherited-numbering boundaries.
