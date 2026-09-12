## Context

The committed PPJ text projection already parses a strict direct rich-text outer shadow and exposes its blur radius. The same native shadow object contains `dist`, while source-bound text edits already locate a text leaf by run and splice one outer-shadow attribute. The change must extend that narrow profile without rebuilding the effect list or touching unrelated package members.

## Goals / Non-Goals

**Goals:**

- Carry an existing bounded `outerShdw/@dist` through native projection, PPJ schema/registry, edit proof, token patching, and second projection.
- Keep the blur leaf and existing authored shadow behavior unchanged.
- Exercise the normal source-bound footprint and fail-closed boundaries with a focused codec fixture.

**Non-Goals:**

- Adding shadow color, angle, alignment, opacity, or transform leaves for direct runs.
- Creating or deleting a missing distance attribute, accepting transformed/compound effect graphs, or proving PowerPoint rendering.
- Changing the wire protocol version or implementing a general effect editor.

## Decisions

1. **Use a sibling native leaf rather than widening the existing blur leaf.** `textShadowDistanceEmu` maps directly to `run.style.shadow.distance`, so callers can edit distance independently while the existing blur proof and source token remain intact. The direct-run profile continues to require a strict one-effect list and rejects transforms and sibling effects.

2. **Keep the distance bound at 0..1,270,000,000 EMU.** This is the finite distance range already used by the typed text-shadow model (100,000 points at 12,700 EMU per point). Canonical decimal tokens are required so stale or normalized values cannot silently pass proof.

3. **Reuse the run-index/source-bound patch path.** Projection records the same text-leaf index as the blur field. Edit planning proves the current `dist` token against the native leaf, and the patcher replaces only that value. This preserves all sibling attributes, color children, text topology, and non-target OPC parts by construction.

4. **Update generated capability data from the registry.** The registry remains the source for the capability matrix; the presentation Skill reference is synchronized from the checked-in capability data. No protobuf or Office wire change is needed because native leaves are PPJ metadata.

## Risks / Trade-offs

- **[A source contains a valid distance but lacks blur]** → Keep the strict direct-run profile aligned with the existing shadow owner and only issue the new leaf when the same bounded, source-preserving shadow profile is recognized; such a graph remains opaque rather than being partially rewritten.
- **[A noncanonical or oversized token is reintroduced after projection]** → Validate both expected and requested values and require an exact decimal token match before patching.
- **[A future effect-list feature changes the run index]** → Bind and reprove the text leaf against the source hash and fail closed on a stale proof; do not infer ownership from a semantic index alone.
