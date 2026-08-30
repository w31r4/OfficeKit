## Context

`PptxTimingCodec` distinguishes three native states: an absent graph that can
accept canonical timing, a recognized canonical graph that can be replaced,
and an opaque graph that must survive untouched. The semantic wire carries
`timing_addable`, `timing_editable`, the original semantic hash, typed
animations, and optional Morph. Export re-reads the concrete SlidePart and
rejects a stale hash or unsupported replacement.

PPJ already validates animation targets, effect/phase pairs, text/chart build
modes, per-page semantic count, and the 64-node expansion budget. The missing
link is source authority and stable-ID remapping during lowering.

## Goals / Non-Goals

**Goals:**

- Let PPJ add, replace, reorder, or remove canonical object animations on a
  proven imported page.
- Keep the complete requested animation array as the declarative source of
  truth; do not add procedural animation commands.
- Resolve top-level and group-descendant target IDs to their native wire IDs.
- Reuse the existing timing writer and native source hash checks.
- Keep opaque third-party timing byte-preserved.

**Non-Goals:**

- Editing or replacing opaque timing graphs.
- Editing Morph pairs in this slice.
- Animating a just-authored source overlay before build/reimport.
- Adding sound, media actions, trigger shapes, arbitrary motion paths, raw
  preset IDs, or raw timing XML.
- Claiming PowerPoint or Keynote playback evidence from structural round-trip.

## Decisions

### 1. `animations[]` remains the only motion state

The Agent edits the page's ordinary PPJ animation array. Adding an entry adds
canonical timing; changing an entry replaces that semantic unit; removing all
entries removes a canonical graph. Array order remains the timeline order.
There is no `animationOperations` list.

### 2. Page nativeRef grants exact timing authority

Projection issues `setAnimations/animations` only when the slide has no Morph
and its source binding reports `timing_editable` or `timing_addable`. Build
reprojects the exact source, validates the page nativeRef and capability, and
lets `PptxTimingCodec` independently compare the original semantic hash.

### 3. Stable PPJ targets are remapped, not serialized as native IDs

PPJ animation targets use stable element IDs. Source lowering walks the
freshly projected PPJ and semantic wire trees together, producing a private
ID map for the current build. The canonical animation lowerer determines the
target kind and options; only its target ID is replaced with the wire element
ID expected by the native timing codec. PPJ never exposes drawing IDs.

### 4. Overlay creation and motion are separate revisions

The native writer applies timing before allocating IDs for a newly appended
source overlay. Therefore a source-bound build that creates an overlay cannot
also animate it. Existing overlay isolation already rejects that mixed
transaction. After build/reimport the new object has a nativeRef and can be a
motion target in the next revision.

## Risks / Trade-offs

- [Agent mistakes structural success for playback proof] -> Review guidance
  labels this structural until a real host plays it.
- [Animation targets a nested group child] -> Recursive PPJ/wire mapping
  supports only the exact projected tree and fails on topology drift.
- [Page contains Morph] -> No animation-edit capability is issued; paired
  Morph requires its own source-bound slice.
- [Opaque timing disappears from projection] -> No capability is issued and
  the native requested-opaque-noop path preserves it.

## Migration Plan

Add one native capability enum value and projection vocabulary. Existing PPJ,
PPTX, tasks, and authored motion remain valid without migration.

## Open Questions

None.
