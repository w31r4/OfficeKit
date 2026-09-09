## Context

See proposal.md. `BuildConnector` currently substitutes its frame for absent literal x/y. `MaterializeSlide` already has the complete expanded page and is the common export/preview path. The expander transforms each child frame with the component transform even inside a group, but leaves literal connector endpoints unchanged. Source-bound `ApplyConnectorElement` currently disallows from/to edits. The native connector model already distinguishes endpoint coordinates, native target IDs and connection-site indexes, while projection maps native attachments to `anchor: auto`.

## Goals / Non-Goals

**Goals:** One deterministic resolver shared by lowering and supported source-bound updates; exact endpoint identity and minimal edit footprint. Complete the requirements in the spec before claiming this change done.

**Non-Goals:** New obstacle routing, geometry-dependent native connection-site evaluation, or the independent switch of production preview to compiler scenes. Connector-to-connector and unresolved component-instance virtual ports remain unsupported and must be rejected explicitly.

## Decisions

1. Add a resolver after component expansion and before slide capture/export. Index object frames and each ancestor transform; resolve in slide space and invert the connector parent transform. This avoids a second layout implementation in the painter. Use a small double-precision affine transform, then round to EMUs once. Literal endpoints stay in their parent coordinates.
2. Enumerate auto side candidates in top/right/bottom/left order and minimize the pair distance in slide space. This avoids recursion, declaration-order dependency and unstable mutual auto selection. Reject connector targets rather than derive an endpoint from an unrelated placeholder frame.
3. On component expansion, transform literal endpoints and external group frames. Recurse through a group's children with identity placement because its childFrame already defines their local coordinates. Continue rewriting component identities. Preserve absent childFrame semantics by materializing the original effective child frame before moving the group.
4. Frame-anchor identity requires a bounded OfficeKit connector extension, separate from `a:stCxn/a:endCxn` geometry sites. Carry explicit optional anchor metadata and its native target identity through the wire and codec, validate its topology strictly, and map it back to semantic IDs during projection. Embedded PPJ alone is insufficient proof; inventing site zero or normalizing the request to literal coordinates would lose its meaning. Native host attachment is only claimed when a genuine connection-site binding is independently established.
5. For source-bound requests, resolve supported frame-anchor dependencies from the candidate's final object frames, including target moves, then apply only the affected connector and target in their SlidePart. Preserve unknown/native imported attachment topology. Unsupported fast leaf paths must route to the full candidate path or reject rather than leave stale endpoints. Validate extension target IDs before granting editing authority.

## Risks / Trade-offs

- Native geometry sites do not share frame anchor semantics → keep their model and authority separate; strict metadata and fresh-projection tests are required.
- Correcting group expansion changes previously incorrect geometry → cover nested child frames and existing component/group tests.
- A coordinate-only intermediate implementation could look complete in preview → leave persistence/editing and final integration tasks unchecked until their tests pass; do not publish a completion claim based on endpoints alone.
- Source-bound leaf shortcuts could bypass dependency updates → specifically test moving a target through its existing supported edit surface.

## Migration Plan

Keep the existing PPJ endpoint syntax. Additive native metadata must be optional; files without it retain their existing native/source-owned behavior. Develop in an isolated worktree and publish only the reviewed scope. Focused C# compiler/codec checks, JS protocol checks if the wire changes, and strict OpenSpec/documentation checks suffice for this increment; no host or full repository acceptance is inferred.
