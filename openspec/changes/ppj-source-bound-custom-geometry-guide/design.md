## Context

See `proposal.md` for the motivation. Custom geometry already has a source-preserving model for `gdLst`, while the native-leaf editor already supports bounded scalar values and token splicing for `avLst` adjustments.

## Goals / Non-Goals

**Goals:**

- Reuse the existing native-leaf binding and edit-plan wire shape.
- Issue leaves only for direct `gdLst` entries whose formula is a literal `val N` signed integer.
- Identify a guide by the existing ordered native-leaf index plus the source hash, and splice only its `fmla` attribute.
- Prove the same shape, guide, and value constraints when applying an edit and reopening the package.

**Non-Goals:**

- Evaluating or editing guide formulas.
- Exposing guide names as a new executable expression language.
- Editing handles, connection sites, text rectangles, paths, multiple guides as a graph, or descendant transforms.
- Adding a protocol field or changing the Office wire version.

## Decisions

- **Use a native leaf rather than a public geometry array.** The existing projection intentionally keeps formula-backed geometry opaque. A leaf gives agents one auditable scalar without implying that the rest of the graph is modeled.
- **Use the ordered `gdLst` index as the location proof.** The source element hash, leaf expected hash, and index are already part of the native-ref/edit-plan contract. Requiring the exact direct list and rechecking the guide at that index prevents a reordered or ambiguous graph from being edited.
- **Require exact `val N` syntax.** The compiler already uses this bounded literal form for geometry adjustments. Other formulas remain untouched, and values are emitted using invariant decimal formatting.
- **Token-splice the existing `fmla` attribute.** Rebuilding `custGeom` could reorder or discard unsupported XML. The edit plan therefore locates the indexed direct `a:gd` and replaces only its `fmla` attribute after validating the source precondition.

## Risks / Trade-offs

- [Risk] A source can contain a guide list that is semantically valid but not safe for this narrow proof. → Suppress the leaf unless the direct-child, attribute, uniqueness, and literal-formula checks all pass.
- [Risk] A guide value can be numerically valid but make the rendered geometry nonsensical. → This capability preserves the native guide semantics and bounds the scalar only; visual quality remains the caller's responsibility and the surrounding graph is not rewritten.
- [Risk] Native-leaf indexes are invalidated by topology changes. → Keep the full nativeRef identity and expected hashes immutable; stale or reordered leaves fail closed.

## Migration Plan

No migration is required. Existing PPJ programs and existing custom-geometry edits remain unchanged; imported sources gain the leaf only when the strict profile is proven.
