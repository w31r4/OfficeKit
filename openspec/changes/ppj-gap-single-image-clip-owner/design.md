## Context

See `proposal.md` for the motivation. The PPJ schema already validates the
shape of `compositing.clipStack`, while the semantic validator and authored
compiler currently reject every non-empty stack. `PresentationImage` already
has one editable mask owner: a preset plus ordered adjustment slots, or a
custom path. `PptxPictureCodec` validates and writes that owner and the PPJ
projector already exposes it as `image.mask`.

## Goals / Non-Goals

**Goals:**

- Make the smallest useful `clipStack` profile executable for authored PPJ.
- Preserve a single unambiguous native owner and its editable preset values.
- Prove authored compile, native XML, embedded-snapshot-free projection, and
  semantic recovery.
- Keep unsupported compositing topology visible as a diagnostic.

**Non-Goals:**

- General clip stacks, inverse clips, arbitrary custom clip paths, group/page
  clipping, masks combined with clips, blend modes, isolation, or effect-order
  semantics.
- Adding another wire field or reconstructing a multi-layer render graph.
- Treating the native mask projection as proof that an arbitrary source clip
  stack is losslessly recoverable.

## Decisions

1. **Use the picture mask as the clip owner.** A native picture has exactly one
   geometry owner in the bounded codec. Mapping the single clip to that owner
   makes it editable and re-projectable. A separate clip field would require a
   new wire contract and still could not represent multiple native geometry
   layers.

2. **Require one non-inverse preset geometry and no direct `image.mask`.**
   The validator accepts only a one-item stack whose `geometry.kind` is
   `preset`, whose preset has the existing adjustment profile, and whose
   adjustments pass the existing count/range rules. A direct mask plus a clip
   is rejected because the native picture has no second mask owner. Custom
   geometry remains reserved for the existing `image.mask` path.

3. **Canonicalize projection to `image.mask`.** After the embedded PPJ snapshot
   is removed, the existing native projector emits the one native mask as
   `image.mask`. The output need not reproduce the authored spelling
   `compositing.clipStack`; it must preserve the bounded visual/editable
   semantic and make the canonical owner obvious.

4. **Validate before lowering and retain a defensive compiler gate.** Normal
   request validation reports `ppj.compositing.clipUnsupported` with the clip
   path. The compiler repeats the narrow structural checks so direct callers
   cannot silently overwrite a direct mask or flatten an unsupported clip.

## Risks / Trade-offs

- [A single clip is narrower than the schema's eight-item array] -> Document
  the exact profile and retain fail-closed diagnostics for every broader form.
- [A preset name or adjustment list may be valid for a shape but not for the
  native picture profile] -> Reuse the same preset-profile and adjustment
  validation used by `PptxPictureCodec` before assigning the image owner.
- [Projection changes the authored field spelling] -> State the canonical
  `image.mask` representation in the reference and test after removing the
  embedded PPJ snapshot.

## Migration Plan

No migration is required. Existing `image.mask` programs and unsupported
`clipStack` programs keep their current behavior. Programs using the new
bounded form can be authored immediately; older readers continue to see the
native picture mask or their existing unsupported response.
