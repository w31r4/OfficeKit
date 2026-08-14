# Source-bound SmartArt clone and plain-node text update

OfficeKit imports SmartArt as a native `nativeObject` with
`nativeKind === "diagram"`. It never reconstructs a diagram as ordinary
shapes. There are two deliberately separate, source-bound contracts:

1. an unchanged closed graph may travel through `slide.duplicate()`; and
2. one narrow DiagramDataPart profile may replace text in an existing one-run
   document node through `nativeObject.setDiagramNodeText()`, or in one
   existing styled run through `nativeObject.setDiagramNodeRunText()`.

Neither contract is SmartArt authoring, layout editing, graph editing, or raw
XML access.

## Closed graph cloning

One unchanged SmartArt frame may travel through the bounded imported
`slide.duplicate()` transaction when all of these preconditions hold:

- the object is one top-level `p:graphicFrame`, never nested in `p:grpSp`;
- it has exactly one `dgm:relIds` root;
- `r:dm`, `r:lo`, `r:qs`, and `r:cs` are present once each and use four unique
  slide-local relationship IDs;
- those relationships are internal and have the standard diagram-data,
  layout, quick-style, and colors relationship types;
- each target has the matching standard OOXML content type, non-empty bytes,
  and no child, external, hyperlink, or data relationship;
- the native object and independent OPC inspection agree on all four IDs,
  target paths, content types, and SHA-256 digests.

```ts
const source = presentation.slides.getItem(0);
const diagram = source.nativeObjects.items.find(
  (object) => object.nativeKind === "diagram",
);

if (!diagram || diagram.parts.length !== 4) {
  throw new Error("Source does not expose one closed SmartArt graph.");
}

const clone = source.duplicate();
const output = await PresentationFile.exportPptx(presentation);
const rebound = await PresentationFile.importPptx(output);
```

The first export keeps the source SlidePart and its relationship part
byte-identical. It preserves the four slide-local relationship IDs but creates
four distinct typed diagram parts for the clone. Each new part is a byte copy
of its corresponding source part. After reimport, source and clone have
disjoint part paths and per-role hashes match.

Use `examples/officekit-slide-duplicate-workflow.mjs` for an Agent-facing
transaction. Its audit records the source and clone SlideParts, all four
relationship IDs, source and clone part paths, content types, hashes, exact
allowed package delta, second-import evidence, and model-render equivalence.
It writes neither output nor audit when any precondition fails.

## Canonical plain-node text profile

An imported graph that passes the closed four-part contract gains a separate
text capability only when its DiagramDataPart additionally proves all of the
following:

- its root is `dgm:dataModel` and it has exactly one direct `dgm:ptLst`;
- every exposed `dgm:pt` has `type="doc"`, a unique non-empty `modelId`, and
  one through 256 direct `dgm:t > a:p` paragraphs, each with at least one
  direct `a:r > a:t`, and one through 256 runs and fixed line breaks in total;
  optional `a:bodyPr`, `a:lstStyle`, `a:pPr`, per-run `a:rPr`, canonical
  `a:br` with optional `a:rPr`, and `a:endParaRPr` may remain;
- each run and its complete node concatenation are XML-safe and the node is at
  most 32,767 characters. Empty paragraphs, fields, attributed or otherwise
  noncanonical breaks, unknown child markup, disconnected parts, or any
  topology ambiguity withhold the capability rather than being simplified.

`nativeObject.editable` remains `false`: this is a typed exception, not general
write authority. `nativeObject.diagramText` is a defensive snapshot containing
the source data part, eligible node IDs, concatenated text, and `runs` flattened
across direct paragraphs in exact source order. Paragraph boundaries stay in
the source XML and are never reconstructed from this projection. A whole-node
replacement remains available only when that node has one run:

```ts
const diagram = presentation.slides.getItem(0).nativeObjects.items.find(
  (object) => object.nativeKind === "diagram" && object.diagramText,
);
if (!diagram) throw new Error("No canonical plain-node SmartArt target.");

const node = diagram.diagramText.nodes.find((item) => item.text === "Before");
if (!node) throw new Error("Expected source text is not unique.");
diagram.setDiagramNodeText(node.id, "After");

const output = await PresentationFile.exportPptx(presentation);
```

For a styled or multi-paragraph node, select one exact existing source-ordered
run. OfficeKit keeps every paragraph, neighboring run, `a:pPr`, `a:rPr`,
canonical `a:br`, and `a:endParaRPr` untouched. Breaks are fixed topology and
do not appear as editable entries in `runs`:

```ts
const styledNode = diagram.diagramText.nodes.find((item) => item.runs?.[1] === " approval");
if (!styledNode) throw new Error("Stale SmartArt run target.");
diagram.setDiagramNodeRunText(styledNode.id, 1, " decision");
```

Export re-proves the original graph, source digest, node IDs/order, and the
plain-node profile. It may rewrite only the bound DiagramDataPart; it preserves
the graphic frame, `dm/lo/qs/cs` relationship IDs, layout, quick-style,
colors, geometry, and every non-data package part. The output is reimported and
must expose the exact requested node list. Leading or trailing replacement
whitespace is serialized with `xml:space="preserve"`.

Use `examples/officekit-smartart-text-edit-workflow.mjs` for a no-overwrite
Agent transaction:

```sh
officekit run "$SKILL_DIR/examples/officekit-smartart-text-edit-workflow.mjs" \
  input/source.pptx output/edited.pptx output/edited.audit.json \
  "Closed SmartArt" "{B31B1833-2B65-4D6B-B3D4-9B3988427B21}" "Before" "After"
```

Append `--run-index=1` to bind `expectedText` and `replacementText` to that
zero-based run instead of the whole one-run node.

It protects the input bytes, resolves exactly one object/node/expected text,
checks that only the DiagramDataPart changed, reimports the graph, and writes a
source/output-bound audit. Its model verification is structural evidence; run
the normal LibreOffice/Poppler render review when a native rendering result is
required.

Paragraph, node, or run creation/removal/reordering, `modelId` changes,
whole-node writes across multiple style runs, presentation of arbitrary diagram
text, raw XML mutation, layout/style/color edits, geometry edits,
cross-diagram changes, clone-before-export after a pending text edit, and
arbitrary graph cloning remain unsupported. Incomplete, duplicated, mistyped,
external, nested, relationship-bearing, field-bearing, empty-paragraph,
noncanonical-break, or otherwise noncanonical SmartArt graphs fail closed.
Canonical fixed `a:br` leaves remain source-owned and are preserved unchanged.
Preserve such objects unchanged or use a separate explicit OPC operation whose
scope is independently reviewed.
