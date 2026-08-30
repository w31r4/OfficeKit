## Context

`PptxLegacyCommentsCodec` already proves one relationship-free shared author
catalog and one relationship-free slide comments part. In that profile, only
`p:cm/p:text` is mutable. PPJ has the semantic comment state and an optional
nativeRef slot but does not connect them.

## Goals / Non-Goals

**Goals:** project comment binding/capability, edit text only, preserve exact
fixed topology, and recover the new text after reimport.

**Non-Goals:** comment addition/removal, metadata changes, modern comments,
replies, resolution, anchors, author changes, or review-topology synthesis.

## Decisions

### 1. Reuse `replaceText`

A comment is a semantic text-bearing object. Its nativeRef uses the same
`replaceText` operation and `text` field as other text owners, scoped to the
comment object rather than the page or an element.

### 2. Bind to comment and page evidence

The comment object hash includes the imported wire comment, including native
author/index evidence. The requested nativeRef must equal the fresh projection.
The PPTX codec then independently re-proves the owning page binding, author
catalog, comment part, fixed topology, and text-only delta.

### 3. The comment list is structurally fixed

Source-bound compile requires identical comment count, order, IDs, page,
author, date, resolved state, position, target/parent absence, and nativeRef.
Only `text` may differ.

## Risks / Trade-offs

- [Legacy PPJ looks like the richer authored comment model] -> Guidance states
  that imported capability covers text only; absent capability means read-only.
- [A caller edits metadata with text] -> Compare every JSON member except text
  before native export and fail closed.

## Migration Plan

Additive projection only. Existing PPJ remains valid.

## Open Questions

Comment-free source-bound addition remains a later explicit parity slice.
