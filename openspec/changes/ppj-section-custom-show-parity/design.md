## Context

`PptxSectionCodec` owns the canonical PowerPoint 2010 section extension and
allows a fixed number and identity of sections to change name and repartition
the ordered slide list. `PptxCustomShowCodec` similarly allows fixed custom-show
identities to change name and ordered membership, including duplicate page
references. Both re-read and compare the resulting native graph. PPJ currently
stops before those codecs.

## Goals / Non-Goals

**Goals:**

- Make imported `sections[].name/pages` and `customShows[].name/pages` editable
  when the native graph already proves that operation.
- Preserve fixed count, order, facade/native identity, and source hashes.
- Change only `ppt/presentation.xml` and recover exact state after projection.

**Non-Goals:**

- Adding or removing imported sections/custom shows.
- Changing native identity or combining membership edits with page topology.
- Repairing opaque or extension-bearing graphs.

## Decisions

### 1. State stays declarative

The Agent edits the existing arrays. `setName` authorizes `name`; new
`setPages` authorizes `pages`. PPJ does not add an operation list.

### 2. Every item carries its own source proof

Projection binds each item to its native element hash and source revision. The
compiler requires an unchanged nativeRef, fixed array position and ID, and an
expected hash matching the issued capability before mutating the native model.

### 3. Native codecs remain the semantic authority

PPJ only maps stable page IDs to the freshly imported slide IDs. The section
codec enforces a complete in-order partition; the custom-show codec enforces
bounded valid membership. Their existing source bindings and post-write re-read
remain the final proof.

## Risks / Trade-offs

- [Page membership references stale IDs] -> Resolve against the fresh requested
  page list and reject missing IDs before export.
- [A section edit silently overlaps or omits slides] -> Delegate to the native
  partition validator; do not normalize or guess.
- [A custom-show link changes identity after rename] -> Preserve facade/native
  IDs and let the native catalog retain hyperlink identity.

## Migration Plan

Additive optional `nativeRef` plus one capability vocabulary item. Existing
source-free PPJ remains valid.

## Open Questions

None.
