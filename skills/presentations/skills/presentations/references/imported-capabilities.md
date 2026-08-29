# Imported presentation capabilities

Load this index after [Edit existing](../tasks/edit-existing.md) has selected a
specific object or source-bound operation. It is a router, not another API
specification. The authoritative contract is the linked file under
`artifact_tool/api/references/`; the linked example is the shortest executable
proof.

## One transaction

1. Copy the input into the task workspace and record its SHA-256. Never edit the
   input path.
2. Import the managed copy and inspect the exact slide, object, and capability.
3. Choose the highest-level typed operation in the table below. Use a native
   leaf only when the inspection issued that exact leaf and hash.
4. Export to a new path, reimport, and inspect again. Locators are revision
   bound and never survive an export by assumption.
5. Review the declared footprint against the immutable source before commit.

Unknown topology remains opaque. A missing, stale, shared, or ambiguous
capability is a refusal, not an invitation to rebuild the deck.

## Capability router

| Intent | Inspect first | Use | Read contract | Executable proof |
| --- | --- | --- | --- | --- |
| Continue or reuse a source slide | `slide.cloneCapability`, `slide.continuationCapability` | `slide.duplicate()` or source-slide reuse | [slide](../artifact_tool/api/references/slide.spec.md) | `officekit-slide-duplicate-workflow.mjs` |
| Change one imported text/style/geometry leaf | `includeNativeLeaves: true`, issued `leaf.expectedHash` | `presentation.editNativeLeaf(...)` | [inspect](../artifact_tool/api/references/inspect.md) | `officekit-object-accessibility-edit-workflow.mjs` |
| Change a repeated bounded component | component candidate and occurrence capability | `presentation.editComponentOccurrence(...)` | [inspect](../artifact_tool/api/references/inspect.md) | component workflow in `examples/` |
| Replace or safely edit an image/SVG | `image.svgEditCapability`, issued SVG leaf | typed image replacement or `image.editSvgLeaf` / `editSvgText` | [images](../artifact_tool/api/references/images.spec.md) | `officekit-object-accessibility-edit-workflow.mjs` |
| Change one table cell | table capability and exact cell value | `table.getCell(row, column).value` | [tables](../artifact_tool/api/references/tables.spec.md) | table examples in `examples/` |
| Change chart title/data | chart capability and fixed topology | chart title/data API | [charts](../artifact_tool/api/references/charts.spec.md) | `officekit-chart-families-workflow.mjs` |
| Delete a top-level ordinary object | `element.deletionCapability` | `element.delete()` | [slide](../artifact_tool/api/references/slide.spec.md) | `officekit-slide-duplicate-workflow.mjs` |
| Edit a group or connector | group/connector capability and stable IDs | typed group/connector operation | [grouping](../artifact_tool/api/references/grouping.spec.md), [connectors](../artifact_tool/api/references/connectors.md) | `officekit-slide-duplicate-workflow.mjs` |
| Add or edit speaker notes | `slide.speakerNotes.capability` | notes API | [speaker notes](../artifact_tool/api/references/speaker-notes.spec.md) | `officekit-speaker-notes-add-workflow.mjs` |
| Add or edit review comments | `slide.comments.capability` | comment API | [comments](../artifact_tool/api/references/comments.md) | `officekit-modern-comment-workflow.mjs` |
| Rename or repartition sections | `presentation.inspect({ kind: "section" })` | section API | [sections](../artifact_tool/api/references/sections.spec.md) | section workflow examples |
| Edit a custom show/view/transition | inspect the exact native profile | corresponding typed API | [custom shows](../artifact_tool/api/references/custom-shows.spec.md), [transitions](../artifact_tool/api/references/transitions.spec.md) | matching workflow in `examples/` |
| Edit an embedded Office payload | OLE capability and unique package owner | supported embedded-package API | [OLE](../artifact_tool/api/references/ole-workbooks.spec.md) | `officekit-ole-office-package-workflow.mjs` |
| Edit canonical SmartArt text | closed `diagramText` capability | issued diagram text run | [SmartArt](../artifact_tool/api/references/smartart-clone.spec.md) | `officekit-smartart-text-edit-workflow.mjs` |
| Preserve/clone InkML or media | closed source capability | clone only; no inferred editor | [InkML](../artifact_tool/api/references/inkml-content-part-clone.spec.md), [video](../artifact_tool/api/references/embedded-video-clone.spec.md) | slide-duplicate workflow |
| Set accessibility metadata | object `accessibilityCapability` | object metadata API | [accessibility](../artifact_tool/api/references/accessibility.spec.md) | `officekit-accessibility-audit-workflow.mjs` |

Use `presentation.designProfile()` and component candidates as evidence for
selection, never as permission to edit raw XML. Reuse a source slide or
component only after its ownership graph is closed and the capability says it
is supported.

## Safety boundary

- Use the public `office-kit` package through `officekit run`. Never use
  `@oai/artifact-tool`, ZIP manipulation, XPath, arbitrary part paths,
  relationship IDs, or synthetic native leaf IDs.
- Bind every operation to the current source revision, expected text/value,
  object identity, and dependent ownership hashes. Stale or cross-object input
  fails closed before model mutation.
- Keep imported Master, Layout, Theme, OLE, SmartArt, animation, comments,
  notes, and unknown parts opaque unless the selected contract explicitly
  grants the operation. A preserved part is not automatically editable.
- Keep edits local. Do not flatten, redraw, reflow, or rebuild a whole deck to
  make one unsupported target succeed.
- Prove the mutation footprint, reimport the output, render the changed page,
  and compare non-target parts/pages with the immutable source. Structural
  validity is not visual approval.

For a source deck that is the actual starting state, use
[Source continuation](source-continuation.md). For a style reference that
should be freely recomposed, use
[Reference-deck conditioned generation](template-conditioned-generation.md)
instead; do not mix the two authorities.
