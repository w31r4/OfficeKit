## Context

`PresentationOpaqueElement.diagram_text` exists only when the source frame owns
one closed four-part SmartArt graph and the DiagramDataPart has a recognized
plain text profile. Its node list retains immutable native model IDs, complete
text, and ordered run text values. Export already rejects identifier changes,
run-count changes, stale parts, relationship drift, and unsupported topology.

PPJ already defines `smartArt` with `mode: "source-bound"`, node text, and
nativeRef authority. The missing link is a typed projection and lowering path.

## Goals / Non-Goals

**Goals:**

- Make safe imported SmartArt text discoverable and directly editable in PPJ.
- Preserve exact native node and run identity without exposing part paths or
  relationship IDs to the Agent.
- Reuse the existing writer and postwrite verification.
- Keep unsupported SmartArt byte-preserved as opaque source content.

**Non-Goals:**

- Editing diagram layout, topology, edges, styles, colors, geometry, or assets.
- Adding or deleting nodes or runs.
- Creating native SmartArt from source-bound state; authored PPJ diagrams keep
  using OfficeKit's editable native-shape lowering.
- Exposing raw DiagramML, model IDs as public cross-file identities, XPath, or
  relationship IDs.

## Decisions

### 1. Typed state replaces the opaque presentation, not the native graph

A proven diagram projects as a PPJ `smartArt` element. Its `mode` is
`source-bound`; its element nativeRef carries the source authority. The native
wire remains the original opaque element with its complete preserved part
closure. PPJ is a semantic view, not a reconstruction of DiagramML.

### 2. Stable PPJ node IDs are derived from source model IDs

Projection derives a collision-safe PPJ node ID through the existing unique-ID
context. The node nativeRef records the same source revision and the diagram
text capability. The native model ID stays private in the fresh baseline wire
and is matched by array position plus the unchanged source binding during
lowering.

### 3. Run boundaries remain explicit

One native run projects as a simple string. Multiple native runs project as one
paragraph with one PPJ run per native text leaf. The Agent may change run text
values but cannot add, remove, reorder, or restyle runs in this slice. This
preserves formatting boundaries already owned by DiagramML.

### 4. The complete requested node array is declarative truth

The Agent edits ordinary `nodes[].text`; no procedural SmartArt operation list
is introduced. Source compilation requires identical node identity and run
topology, then copies only changed strings into
`PresentationOpaqueElement.diagram_text`. The existing codec independently
revalidates the bound part and graph before writing.

## Risks / Trade-offs

- [PPJ node ID is mistaken for native identity] -> Documentation states it is
  revision-bound and must be rediscovered after build/reimport.
- [Rich text flattening] -> Multiple native run values remain separate PPJ runs;
  topology changes fail closed.
- [Unsupported diagrams appear editable] -> Only a non-null proven diagram text
  binding projects as typed SmartArt; all others remain opaque.
- [Styles appear modeled] -> Source-bound projection omits authored layout and
  style fields and documents them as source-owned.

## Migration Plan

No migration. Existing PPJ and PPTX remain valid. A fresh projection of a
capable diagram becomes more semantic; an incapable diagram remains opaque.

## Open Questions

None.
