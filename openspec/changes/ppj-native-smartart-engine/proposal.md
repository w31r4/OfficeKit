# PPJ native SmartArt engine

## Why

PPJ currently names authored semantic diagrams `smartArt`, but lowers them to
ordinary grouped shapes. That makes the content editable while discarding the
Office SmartArt identity, data graph, layout definition, and reusable style and
color programs. Imported SmartArt has the opposite limitation: OfficeKit can
preserve the native graph and edit a proven text leaf, but it cannot project a
portable definition or rebuild a supported diagram from semantic PPJ state.

## What changes

- Treat one SmartArt object as the ownership and transaction boundary while
  keeping nodes, connections, layout, style, colors, and assets independently
  editable when their capability is proven.
- Compile authored PPJ SmartArt to native DiagramData, Layout, Style, Colors,
  and cached Drawing parts rather than an ordinary Presentation group.
- Add a content-addressed JSON definition asset for custom standard DiagramML
  definitions. Built-in layouts remain compact PPJ tokens backed by clean-room
  definitions.
- Project every imported diagram as typed PPJ SmartArt. Standard supported
  regions become semantic state; unknown extensions remain hash-bound native
  residue and block only edits that touch them.
- Provide an explicit detach-to-shapes conversion for callers that knowingly
  trade SmartArt semantics for ordinary editable DrawingML.

## Boundaries

- V1 owns the deterministic operator and constraint profiles exercised by the
  eight OfficeKit layouts; it does not claim arbitrary PowerPoint layout parity.
- PPJ never exposes part paths, relationship IDs, XPath, or raw SmartArt XML.
- Unknown or unsupported definitions are preserved and fail closed on affected
  mutation. No automatic flattening or raster fallback is permitted.
- Existing authored SmartArt output intentionally changes from grouped shapes
  to native SmartArt; there is no legacy-output switch.
