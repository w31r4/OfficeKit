## Context

See `proposal.md` for the motivation. The previous bounded clip profile maps
one preset clip to the picture's preset mask. Existing authored `image.mask`
support already parses custom `viewBox`/`paths` into `PresentationImage` custom
mask paths, and `PptxPictureCodec` validates and writes that graph. The
projector already emits that graph as `image.mask` when it is safe to expose.

## Goals / Non-Goals

**Goals:**

- Reuse the existing custom image-mask owner for one compositing clip.
- Apply exactly the same custom path acceptance and native validation boundary
  as the direct image-mask path.
- Keep the clip spelling canonicalized to `image.mask` on snapshot-free import.

**Non-Goals:**

- Multiple or inverse clips, custom clip stacks, clip transforms, group/page
  clipping, mask-plus-clip composition, blend modes, isolation, or effect
  ordering.
- New custom geometry syntax, a new wire field, or arbitrary third-party
  DrawingML path inference.

## Decisions

1. **Reuse `ApplyCustomGeometry` and `PptxCustomGeometryCodec.Validate`.** This
   keeps the clip and direct `image.mask` paths on one parser/validator/native
   profile. A second custom-path grammar would create drift and could accept a
   graph the picture codec cannot write.

2. **Allow only one non-inverse clip and no direct mask.** The picture still has
   one geometry owner. The validator and compiler reject any combination that
   would require composing two masks or preserving stack order.

3. **Use the existing custom-mask projector.** A native picture custom geometry
   is already projected to `image.mask.kind = "custom"`; no synthetic
   `compositing` object is emitted from one owner.

4. **Keep native validation authoritative.** Semantic validation performs the
   same conversion and custom-geometry validation used by authored lowering so
   malformed or unsupported path graphs report `ppj.compositing.clipUnsupported`
   before writing a file. The compiler repeats the gate defensively.

## Risks / Trade-offs

- [A custom path can be visually complex] -> Reuse the existing command,
  coordinate, path-count, and native custom-geometry limits; do not flatten it.
- [Projection changes the authored field spelling] -> Test after removing the
  embedded PPJ snapshot and document `image.mask` as the canonical owner.
- [Validator conversion duplicates a small amount of compiler work] -> Keep it
  limited to the existing shared parser/codec boundary and retain the compiler
  guard for direct callers.

## Migration Plan

No migration is required. Existing direct `image.mask` programs are unchanged;
unsupported clip stacks remain rejected. New programs can use the bounded
single custom clip immediately.
