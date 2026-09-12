## Context

The committed PPJ text projection parses a strict direct rich-text outer shadow and exposes its blur and distance leaves. The same native shadow object contains `dir`, while source-bound text edits already locate a text leaf by run and splice one outer-shadow attribute. The change must extend that narrow profile without rebuilding the effect list or touching unrelated package members.

## Goals / Non-Goals

**Goals:**

- Carry an existing bounded `outerShdw/@dir` through native projection, PPJ schema/registry, edit proof, token patching, and second projection.
- Keep the blur and distance leaves and existing authored shadow behavior unchanged.
- Exercise the normal source-bound footprint and fail-closed boundaries with a focused codec fixture.

**Non-Goals:**

- Adding shadow color, alignment, opacity, or transform leaves for direct runs.
- Creating or deleting a missing direction attribute, accepting transformed/compound effect graphs, or proving PowerPoint rendering.
- Changing the wire protocol version or implementing a general effect editor.

## Decisions

1. **Use a sibling native leaf rather than widening the existing distance leaf.** `textShadowDirectionDegrees` maps directly to `run.style.shadow.angle`, so callers can edit direction independently while the existing blur and distance proofs and source tokens remain intact. The direct-run profile continues to require a strict one-effect list and rejects transforms and sibling effects.

2. **Keep the native direction bound at 0..21,599,999.** DrawingML stores direction in 1/60000ths of a degree; the existing text-shadow validator accepts the canonical non-negative full-turn range. Exact decimal tokens are required so stale or normalized values cannot silently pass proof.

3. **Reuse the run-index/source-bound patch path.** Projection records the same text-leaf index as the blur and distance fields. Edit planning proves the current `dir` token against the native leaf, and the patcher replaces only that value. This preserves all sibling attributes, color children, text topology, and non-target OPC parts by construction.

4. **Update generated capability data from the registry.** The registry remains the source for the capability matrix; the presentation Skill reference is synchronized from the checked-in capability data. No protobuf or Office wire change is needed because native leaves are PPJ metadata.

## Risks / Trade-offs

- **A source contains a valid direction but lacks blur and distance:** Keep the strict direct-run profile aligned with the existing shadow owner and issue the direction leaf when the same bounded source-preserving shadow profile is recognized; a graph with no supported geometry remains opaque.
- **A noncanonical or oversized token is reintroduced after projection:** Validate both expected and requested values and require an exact decimal token match before patching.
- **A future effect-list feature changes the run index:** Bind and reprove the text leaf against the source hash and fail closed on a stale proof; do not infer ownership from a semantic index alone.
