# Presentation primitive surface

This is the Agent-facing index of OfficeKit presentation primitives. It is a
navigation aid, not a second API specification. Help (`src/help/index.mjs`) and
the API reference files under `artifact_tool/api/references/` remain the
authoritative contract. Search Help by intent before choosing a primitive.

## How to choose a primitive

1. Identify whether the task creates a source-free slide, continues a trusted
   imported source, or only reviews an artifact.
2. Inspect the slide and its capabilities before editing an imported object.
3. Choose the highest-level primitive that states the user's intent.
4. Keep the mutation local, render it, reopen the PPTX, and verify the declared
   footprint. Unsupported native structure stays opaque or is rejected.

Never use a raw XPath, part path, relationship ID, arbitrary XML patch, or
collection splice as an Agent-facing workaround. `presentation.editNativeLeaf`
and the capability-issued native leaf APIs are the only bounded escape hatch.

## Semantic families

| Family | Use for | Main entry points | Source-bound boundary |
| --- | --- | --- | --- |
| Deck and slide lifecycle | Create, import, order, hide, duplicate, delete, export | `Presentation.create`, `PresentationFile.importPptx`, `presentation.slides.add`, `slide.duplicate`, `slide.moveTo`, `presentation.export` | Clone/delete only after the reported ownership capability passes; never rebuild an imported package to bypass it. |
| Scene stack and surfaces | Backgrounds, images, cross-type z-order, transitions and motion | `slide.setBackground`, `slide.setNativeBackgroundImage`, `slide.images.add`, `slide.elements`, `element.moveBefore`, `slide.animations.add`, `slide.setMorph` | A native `p:bg` is behind the slide tree; a picture layer is movable. Imported underlays and opaque timing fail closed when order cannot be proven. |
| Text and typography | Titles, body copy, runs, notes and native text leaves | `shape.text.set`, `presentation.textRange`, `nativeObject.setDiagramNodeText`, `nativeObject.setDiagramNodeRunText` | Preserve run topology, facts and source styling unless the requested operation explicitly changes them. |
| Shapes, connectors and groups | Diagrams, rules, arrows, frames and nested composition | `slide.shapes.add`, `slide.shapes.connect`, `slide.groups.add`, `connector.setConnectorFrom`, `element.bringToFront` | Imported group topology, connector endpoints and unknown effects are source-bound unless their capability says otherwise. |
| Layout and theme | Master, Layout, placeholders, theme and reusable structure | `presentation.theme`, `presentation.master`, `presentation.layouts.add`, `slide.applyLayout`, `slide.placeholders.getItem` | Do not flatten inherited Master/Layout/Theme content into slide XML. New pages may materialize a resolved layout transactionally. |
| Compose and auto-layout | Build a new page from semantic content | `slide.compose`, `slide.autoLayout`, `compose.text`, `compose.chart`, `compose.image`, `compose.rule` | Helpers are mechanics, not a house style. A composition intent and design grammar choose the carrier and geometry first. |
| Tables and charts | Data relationships, comparison and evidence | `slide.tables.add`, `slide.charts.add`, `chart.delete` | Keep data and labels truthful. Unsupported chart children, external data and advanced plot graphs remain preserved or rejected. |
| Images and SVG | Pictures, crops, editable safe SVG text/style leaves | `slide.images.add`, `image.getSvgTextNodes`, `image.editSvgText`, `image.getSvgEditLeaves`, `image.editSvgLeaf` | Keep the original asset where possible. SVG edits are signed, token-scoped and reject scripts, external resources and topology changes. |
| Imported continuation and reuse | Discover, reuse and locally edit source pages/components | `presentation.inspect`, `presentation.designProfile`, `presentation.resolveComponentCandidate`, `presentation.reuseSourceSlide`, `presentation.reuseSourceComponent`, `presentation.editComponentOccurrence`, `presentation.editNativeLeaf` | Locators and hashes are bound to the source revision. Re-inspect after export/reimport and never assume JavaScript object identity survives. |
| Native packages | Recognized SmartArt, OLE Office payloads and other bounded native objects | `nativeObject.getEmbeddedWorkbook`, `nativeObject.replaceEmbeddedWorkbook`, `nativeObject.getEmbeddedOfficePackage`, `nativeObject.replaceEmbeddedOfficePackage` | Replace only the exact recognized payload while preserving its shell and relationships. Unknown OLE/native graphs stay opaque. |
| Accessibility and delivery | Alt text, decorative state, verify, review and handoff | `shape.setAccessibilityMetadata`, `image.setAccessibilityMetadata`, `presentation.auditAccessibility`, `presentation.validateLayout`, `presentation.verify`, `PresentationFile.exportPptx` | Accessibility metadata is not a claim of whole-deck WCAG conformance. Render and reopen before delivery. |

## Scene order is one shared model

`slide.elements` is the cross-type back-to-front stack. Shapes, textboxes,
images, tables, charts, connectors and groups are views over that same stack.
For image-led composition, decide the order before authoring:

```text
native slide background (optional)
→ background image layer (if movable/cropped)
→ scrim or contrast shape
→ evidence carrier
→ editable labels and data
→ decision line / foreground annotation
```

Use `slide.setNativeBackgroundImage` for a true non-reorderable backdrop. Use
`slide.setBackgroundImage` or `slide.images.add` when the image must be moved,
cropped or animated. A line, marker, label or chart must not be hidden merely
because an image or overlay was added later; inspect the final stack and
rendered result.

## Capability and failure vocabulary

- **typed-editable**: the public semantic object supports the requested intent.
- **native-leaf-editable**: `inspect` issued a bounded leaf with a revision and
  expected hash; use that leaf exactly once within the transaction.
- **source-derived-reusable**: a page/component may be reused after its source
  ownership and relationship capability passes.
- **opaque-preserved**: the source graph survives, but OfficeKit does not claim
  a safe mutation. Explain the reason instead of guessing.

Stale hashes, ambiguous ownership, cross-revision locators, unsupported
relationships, timing/extension graphs, and edits that would alter unrelated
parts must fail closed. A failed edit is evidence about the boundary, not a
reason to lower the source-preservation bar.

## Help and maintenance

The searchable Help catalog is the discovery index. Each public presentation
API should have an adoption tier, `useWhen`, `avoidWhen`, prerequisites, review
notes and a recipe. The generated `docs/api.md` is derived from Help; edit the
catalog rather than hand-editing generated docs.

When a runtime primitive, wire field, Help record, example or review invariant
changes, run the repository's `skill-update` checker. Its impact manifest lists
the owning source paths, consumer Skills, examples, focused tests and release
evidence that must be considered together.

