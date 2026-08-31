# Design

## PPJ facade

The SmartArt element keeps the instance data graph compact. Authored elements
select exactly one built-in `layout` or one `definitionAsset`. `connections`
is the canonical topology representation; it replaces node-local parent sugar.

```json
{
  "type": "smartArt",
  "mode": "authored",
  "layout": "process",
  "nodes": [
    { "id": "observe", "text": "Observe" },
    { "id": "decide", "text": "Decide" }
  ],
  "connections": [
    { "id": "observe-decide", "from": "observe", "to": "decide", "role": "sequence" }
  ]
}
```

Custom definitions use the media type
`application/vnd.officekit.smartart-definition+json` and schema
`office-kit/smartart-definition/v1`. The asset contains typed `layout`,
`style`, and `colors` documents. Editing is copy-on-write: the editor writes a
new content-addressed asset and changes only the selected element reference.

## Native representation

The presentation wire gains an additive `PresentationDiagram` content arm. It
contains the PPJ data graph, the resolved deterministic layout result, and the
definition identity needed by the PPTX writer. The writer creates one native
`p:graphicFrame`, the four standard diagram parts, and an Office 2010 cached
diagram drawing. Stable IDs derive from PPJ element and node IDs; package-local
part paths and relationship IDs are allocated only by the writer.

The first engine profile reuses the existing eight deterministic placements as
the layout oracle and serializes clean-room DiagramML programs describing those
families. This closes native identity without making PowerPoint the only render
oracle. Later operator profiles can replace individual finite placements without
changing the PPJ facade or native ownership boundary.

## Import and preservation

Import classifies data, layout, style, colors, drawing cache, relationships, and
extensions independently. A supported native graph projects as typed SmartArt
with section-scoped capabilities. Standard but not executable definition nodes
may still be represented by the definition JSON. Vendor/version extensions and
unknown OPC descendants stay in a content-addressed preservation fragment.

No-op source-bound export keeps the original parts. A local edit may change only
the sections named by its capability and must re-prove every untouched fragment
hash. The first third-party slice recognizes plain content points, the
layout-definition identity, and parent-of edges while keeping document-root and
presentation wiring private. A diagram with no provable semantic nodes remains
opaque today; projecting every such diagram as typed read-only SmartArt belongs
to the later residue-preservation slice.

## Explicit detach

`detachToShapes()` consumes only a verified cached drawing and replaces the
SmartArt element with an ordinary PPJ group. It returns a semantic-loss warning
and never runs implicitly. Missing, external, or unsupported drawing caches make
the operation unavailable.

## Determinism and visual target

Layout uses integer EMU output, stable traversal, stable model IDs, and canonical
XML ordering. Text uses explicit DrawingML properties and bounded AutoFit; V1
requires usable output without obvious overflow, not pixel equality across
PowerPoint and LibreOffice versions.
