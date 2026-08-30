## Context

The mature source-preserving writer already treats source elements as an
ordered native prefix and authored elements as a topmost suffix. It admits
only canonical textbox, rect, roundRect, ellipse, and embedded rectangular
image overlays. It preserves every retained source element as native XML,
allocates fresh drawing IDs, bounds added image relationships, and rejects an
overlay combined with another SlidePart mutation.

PPJ already represents every authored element type as state in `elements[]`.
The missing link is not another operation list or another OOXML writer; it is a
capability-aware comparison that distinguishes retained source elements from a
new typed suffix and lowers that suffix through the existing compilers.

## Goals / Non-Goals

**Goals:**

- Express imported-slide continuation by appending ordinary typed PPJ state.
- Reuse the authored element lowerer and source-preserving native writer.
- Keep all unknown source content and original z-order below the new suffix.
- Fail closed on stale authority, unsafe element kinds, or mixed mutations.

**Non-Goals:**

- Adding a procedural `add()` or mutation-command DSL.
- Inserting a new object beneath or between source-owned objects.
- Adding tables, charts, connectors, groups, media, SmartArt, or OLE in this
  bounded profile.
- Combining overlay creation with source edits, deletion, reorder, page
  metadata changes, comments, sections, or custom-show changes.
- Exposing raw OOXML, relationship IDs, native drawing IDs, or part paths.

## Decisions

### 1. Ordered PPJ state remains the only authoring surface

An Agent appends a normal typed element to the page's existing `elements`
array. Source-projected elements retain their nativeRef values and remain an
ordered prefix. New elements have fresh IDs, no nativeRef, and form a suffix.
Array order continues to mean bottom-to-top z-order.

### 2. The page grants bounded append authority

Projection adds `appendElement/elements` to the page nativeRef for ordinary
source-bound slides. Build reprojects the exact source, verifies that issued
capability and the unchanged page nativeRef, then recognizes only the new
suffix as authored state. An invented or stale capability cannot authorize a
write.

### 3. Authored and native lowering are reused

The existing authored PPJ lowerer converts the new typed suffix to the
canonical Presentation wire model. The existing PPTX writer independently
validates the bounded geometry/relationship profile and appends native XML.
PPJ does not write XML and does not duplicate the overlay validator.

### 4. One clean slide mutation per build

The PPJ compiler rejects an overlay when the same source slide also changes
metadata, source elements, order, deletion, notes, background, transition, or
animation. This gives the Agent a simple workflow: append, build, reimport,
then continue from the new reviewed revision.

## Risks / Trade-offs

- [Agent expects arbitrary z-order] -> Document that new elements are a
  topmost suffix and require build/reimport before further source-bound work.
- [Authored compiler supports more types than the overlay writer] -> Admit
  only text, four bounded shape geometries, and images before native export.
- [New image bytes are missing] -> Reuse source-bound asset hash/MIME
  validation and content-addressed native asset registration.
- [The native source changes] -> Reproject exact bytes and verify the complete
  page nativeRef before lowering.

## Migration Plan

Additive capability vocabulary only. Existing authored and source-bound PPJ
remains valid. No file migration, wire-version change, or compatibility shim.

## Open Questions

None.
