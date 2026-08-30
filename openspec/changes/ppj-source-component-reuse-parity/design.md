## Context

The JavaScript compatibility facade exposes `reuseSourceComponent` after an
inspect-only repeated-component candidate has selected one occurrence. Its
actual mutation is smaller than that procedural workflow: duplicate a proven
source slide, retain one direct top-level element, and issue explicit deletion
bindings for every sibling. The NativeAOT writer independently re-proves the
clone graph, retained drawing IDs, relationship ownership, connector targets,
and each deletion before changing the package.

PPJ already projects every top-level element with a stable ID and an issued
`delete/element` capability where deletion is safe. Therefore the public DSL
does not need the JavaScript candidate ID or operation sequence. It can state
the desired finite result directly.

## Goals / Non-Goals

**Goals:**

- Express one exact retained native component as declarative PPJ state.
- Reuse the existing full-slide clone and sibling-deletion writer.
- Keep opaque descendants of the retained top-level element byte-preserved.
- Require a fresh capability proof for every removed sibling.
- Reimport the new page into ordinary typed/nativeRef PPJ before editing it.

**Non-Goals:**

- Porting the JavaScript repeated-component candidate analyzer into PPJ.
- Copying a nested child independently of its owning top-level group.
- Retaining an arbitrary set of multiple elements in this first profile.
- Moving, restyling, or editing the pending component before reimport.
- Combining component reuse with overlays, page routes, source mutations, or
  another pending clone in the same build.
- Exposing drawing IDs, shape-tree indices, relationships, or raw OOXML.

## Decisions

### 1. `retainElement` extends the existing finite clone macro

```json
{
  "id": "page-source-component",
  "role": "source continuation",
  "elements": [],
  "sourceClone": {
    "page": "page-source",
    "capability": "cap_...",
    "retainElement": "element-source-group"
  }
}
```

The referenced page remains unchanged and immediately precedes the pending
clone. `retainElement` identifies one entry in that page's top-level
`elements[]`; nested IDs are not page elements and therefore cannot be used.
Omitting the field keeps the existing exact full-page clone behavior.

### 2. Authority is composed from existing capabilities

The source page must freshly issue `duplicate/pageClone`. Every sibling that
will be removed must freshly issue `delete/element`. The retained element does
not need a new invented capability because it remains byte-equivalent to its
source binding. The compiler converts the selected source wire element plus
the proven sibling deletions into the existing clone request.

### 3. Native validation remains the final oracle

The PPJ validator gives early diagnostics for unknown IDs and missing delete
capabilities. At build time the compiler uses the fresh source projection and
the codec checks the actual shape-tree order, native IDs, relationships,
connector targets, source hashes, and deletion plans. A semantic ID alone can
never authorize extraction.

### 4. Build/reimport separates reuse from editing

Like a full-page pending clone, a component clone carries no nativeRef or
explicit elements. After native build and PPJ projection it becomes an
ordinary source-bound page containing the retained object. The Agent can then
use its typed fields, native leaves, topmost overlays, or another supported
operation in a later reviewed revision.

## Risks / Trade-offs

- [The selected object is not independently useful] -> PPJ exposes mechanism,
  not aesthetic judgment; Skill guidance tells the Agent to inspect the
  retained component after reimport.
- [A sibling owns a relationship or connector identity] -> Its delete
  capability is absent or native revalidation rejects the build.
- [The component is nested] -> Retain the owning top-level group; direct nested
  extraction remains unsupported.
- [The Agent expects several objects] -> Use a real top-level group or first
  compose the desired objects on a reviewed source-derived page; multi-retain
  is deliberately outside this bounded slice.

## Migration Plan

Additive optional `sourceClone` state. Existing authored PPJ and full-page
source clones remain valid without migration. No wire or task schema change.

## Open Questions

None.
