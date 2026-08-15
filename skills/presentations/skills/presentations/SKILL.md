---
name: Presentations
description: Read, create, or edit PowerPoint or Google Slides decks. Use for presentation, slide deck, PowerPoint, PPT, PPTX, or Google Slides requests.
---

# Slides Skill

Use this skill as reference material when creating or editing presentation slide decks.

If the user explicitly targets a presentation already open in desktop PowerPoint,
use the sibling `powerpoint-live-control` Skill instead. It operates the live,
possibly unsaved deck through typed Office.js operations and never switches to a
closed-file edit behind the user's back. This Skill remains the route for local
PPTX files and for creating a new deck.

## Run the deck workflow in one task

For a multi-step deck task, use `officekit repl` and the portable contract in
`../office-kit/references/repl.md`. Import the public API with
`await ctx.import("office-kit")`, keep the deck and reusable layout helpers in
`ctx.state`, and follow inspect → compose/edit → render → structural or visual
review → verify. For net-new or broadly redesigned decks, follow
[the conversational workflow](references/conversation-workflow.md), revise a
checked working draft, and call `ctx.publish` only after acceptance; narrow
edits stay direct. Register QA evidence with `ctx.recordEvidence`. Use the
image capability matrix below when generation or inspection is unavailable.
After reopening the accepted final PPTX, follow
`../office-kit/references/review.md` and report its path, SHA-256, and
`visualReview` status.

When a task asks to preserve or tune PowerPoint's shrink-to-fit state, use the
typed `shape.text.bodyProperties.normalAutoFit` profile documented in
`artifact_tool/api/references/rich-text.spec.md`. Do not estimate fitted values
from rendered pixels or patch `a:normAutofit` XML directly; inspect the imported
canonical values, apply the requested bounded percentages, export, reimport,
and render/verify. Noncanonical AutoFit markup must fail closed.

## Important Instructions

- [HARD REQUIREMENT] Use the public `office-kit` package for presentation
  artifact work. Never import or use `@oai/artifact-tool`: it is a different
  host-bundled runtime, not an OfficeKit alias or fallback, and its output must
  never be attributed to OfficeKit.
- [HARD REQUIREMENT] Content quality and storytelling: before planning the deck, read and follow [Content Quality and Narrative Rules](references/content-rules.md) and [Presentation Style Guidelines](style_guidelines.md). Ensure the deck covers everything the user requested and forms a coherent, audience-appropriate narrative rather than a collection of disconnected facts.
- [HARD REQUIREMENT] Audience-facing copy: visible slide content must be written for the intended audience, not for the person or model producing the deck. Do not expose planning notes, timing scaffolds, talk tracks, content-selection commentary, or other internal process language unless the user explicitly requests it.

- Info density: avoid cramming low-value details onto a single slide. Prefer lower-density slides with high-value content.
  - Title slide: keep the title slide minimal and simple. Avoid cramming in too much information.
- Layout: keep things clean and simple. Avoid low-quality visuals, but also avoid excessive white space. By default, use equal left and right margins on each slide.
- [HARD REQUIREMENT] Overlap: always pay attention to programmatic overlap warnings. Do not assume that overlapping elements in diagrams are intentional, and do not ignore overlap warnings without inspecting them. You MUST fix all unintended overlap errors before delivering the slides. This is critical.
- [HARD REQUIREMENT] Font size: when a template is provided, match its font sizes. When no template or style guidance is given, you MUST use at least 50pt for deck titles, 35pt for slide titles, 24pt for mid-level text such as subheadings, callout headers, and text-box titles, and 16pt for body text.
- [HARD REQUIREMENT] Object accessibility: give meaningful ordinary shapes, connectors, groups, images, tables, and charts a concise non-visible `accessibility.title` and, when needed, `accessibility.description`; classify purely decorative objects with `{ accessibility: { decorative: true } }` and no title/description. Presence matters: `decorative: false` is an explicit meaningful-object classification, while omission is unclassified. `image.alt` is only a compatibility alias for `image.accessibility.description`, not a second metadata field. For imported objects, check `object.accessibilityCapability.editable` before calling `object.setAccessibilityMetadata(...)`; change classification and clear/add alternative text in one transaction, and preserve or reject irregular native metadata instead of rebuilding it. Chart metadata is distinct from its visible chart title. This does not establish reading order or make the whole deck accessible by itself.
- [HARD REQUIREMENT] Before delivering a deck, run `presentation.auditAccessibility()`. Resolve every machine issue, then review each returned `manualChecks` record. Reading order and opaque native objects require native-host or source review; never reinterpret `machineCheckPassed: true` as whole-deck conformance. For a read-only imported-file report with immutable-source/hash/no-overwrite evidence, use `examples/officekit-accessibility-audit-workflow.mjs`. For one imported shape, connector, group, image, table, or chart, pass the complete audit locator and prior metadata to `examples/officekit-object-accessibility-edit-workflow.mjs`; it fails closed unless exactly the selected SlidePart changes and the result survives reimport, verify, and visual-stability checks. Read `artifact_tool/api/references/accessibility.spec.md` for both contracts.
- Text layout: when there is too much text, shorten it before shrinking the font size. Inspect visually for unexpected text wrapping. NEVER allow a title/banner text box intended for one line to wrap to two lines.
- Narrative copy must fit the chosen layout: shorten it or change layouts rather than adding density or shrinking type.
- Visual assets:
  - Read `../office-kit/references/capabilities.md` and determine whether the
    active agent has `image_view` and/or `image_generate`.
  - Use user and template assets first, then native PowerPoint shapes,
    connectors, charts, tables, and typography. Generated images are optional
    and must be marked for human review when they cannot be inspected.
  - Minimize diagrams unless they clarify the request. Use native shapes for
    simple diagrams and Graphviz for complex relational diagrams. Create
    connectors before entity nodes when connector layering matters.
  - If no visual input is available, run structural QA for dimensions, image
    placement/crop, text overflow, overlap, contrast, and slide geometry, and
    report `visualReview: "unavailable"` rather than an aesthetic judgement.
  - By default, DO NOT reuse the same image more than once (unless it's a background).
  - Prepare visuals for both the main concept and decorative support.
- Default styling: use one composition instead of a collection of UI panels. UI-like styling typically includes card grids, pills, badges, button-like text boxes, tab or navigation patterns, repeated modular panels, dense dashboard-style layouts, and other component-library aesthetics that imply interactivity. Use stylized text boxes sparingly, favoring a flat structure on the canvas.

## Skill Folder Contents

Contents of the `slides/` skill folder:

- `container_tools/`: Standalone python scripts for slides and relevant asset manipulation.
- `references/`: Additional workflow references for specialized presentation tasks.
- `style_guidelines.md`: Narrative, copy, layout, typography, and visual-consistency defaults.
- `routing/`: Native Google Slides routing guidance.
- `template_following_scripts/`: Helper scripts for exact source-deck/template following.
- `artifact_tool/`: API documentation and coding examples for the artifact tool library.
- `builtin_templates_support/`: Checked-in guidance, manifests, prompts, and reusable scripts for built-in templates. Each template owns its `ARTIFACT.md`; shared runners live once under `builtin_templates_support/scripts/`.
- `assets/builtin_templates/grid-layout-library/`: Blob-managed static assets for the built-in Grid Layout template, including 26 rendered previews, a model-facing registry, structured content tokens, and 26 exact plain-JavaScript artifact-tool Compose reconstructions with no JSX. This directory contains no Markdown, prompts, or reusable runners.

## Container Tools

The following helper scripts are located in the `container_tools/` directory:

- `ensure_raster_image.py`: Ensure images are rasterized; convert to PNG if needed; quick usage `--input_files <img_path1> ...`. PDF input uses the runtime-resolved `pdftoppm` directly, with no undeclared Python wrapper.
- `render_slides.py`: Render a PowerPoint file into a folder of PNG slides using default sizing; quick usage: `<input.pptx>`. Output files are named `slide-1.png`, `slide-2.png`, ... in a directory with the same name as the input file. PDF input uses runtime-resolved `pdfinfo`/`pdftoppm` directly.
- `create_montage.py`: Build a tiled montage from images in a directory (for viewing multiple image assets or rendered slides at once); quick usage: `--input_dir <imgs_dir> --output_file <montage.png>`. It supports most image formats with auto conversion under the hood.
- `slides_test.py`: Detect content overflowing the original slide canvas; usage: `<input.pptx>`.

## Grid Layout Artifact-Tool Compose Layout Reference

This skill variant does not include the Office template file. Use the distilled layout library as initial design and composition guidance when the user has not supplied a stronger template or brand system.

Before planning slides:

1. Read `builtin_templates_support/grid-layout-library/ARTIFACT.md`, `assets/builtin_templates/grid-layout-library/design_tokens.json`, and `assets/builtin_templates/grid-layout-library/artifact-tool-compose/template-registry.json`.
2. Inspect `assets/builtin_templates/grid-layout-library/assets/previews/layout-library.png`, then shortlist layouts by `templateUse`, `layoutFamily`, `slots`, `densityBudget`, and `typographyBudget`. Do not open all 26 implementation modules by default.
3. For each selected layout, inspect its generated preview and exact `assets/builtin_templates/grid-layout-library/artifact-tool-compose/slide-XX.mjs` reconstruction.
4. Use the selected module's `layers(...)`, `text(...)`, `shape(...)`, `image(...)`, and `table(...)` helper calls as the implementation reference. Keep the output as plain `.mjs` and use `slide.compose(...)`; do not introduce JSX or a transpilation step.
5. Preserve the selected layout's content ownership, spacing, hierarchy, and media frames while replacing instructional sample text with the user's content. Vary silhouettes across the deck instead of repeating one pattern.

The shared `builtin_templates_support/scripts/create-presentation.mjs` runner can materialize any compatible built-in template for validation when passed that template's static asset root. It is not a request to emit every layout in the user's deck. User-provided templates, explicit brand guidance, and exact source evidence always override this default template.

## Workspace

Read `../office-kit/references/workspace.md` before authoring. Use the shared
`workspaceRoot`, `taskRoot`, `inputRoot`, `assetRoot`, `outputRoot`, and
`evidenceRoot` names. An explicit user destination always wins; otherwise put
the final deck under `workspaceRoot/outputs` and keep temporary slide sources,
previews, layout notes, and QA under `taskRoot` or `evidenceRoot`. Use
`SKILL_DIR` only to locate bundled files. Keep `sessionId` local to OfficeKit,
never derived from a chat or thread.

## Route the Request Before Authoring

Choose the output path first:

1. **Existing native Google Slides deck**: obtain a local export or reference
   from the user/host, then preserve it as an input. OfficeKit does not operate
   a cloud deck directly.
2. **Net-new native Google Slides deck**: build and verify a local PPTX with
   this skill; a user or host may import it after delivery.
3. **PowerPoint or local deck**: build or edit the PPTX with this skill.

For every deck built with this skill, choose exactly one visual route. The first
matching route wins:

1. **User reference or template skill**: if the user supplies a reference deck,
   asks to follow an existing deck, or invokes a template skill, use only that
   file as the visual source. An existing PPTX being edited also counts as the
   reference. Do not mix in Grid Layout or another template.
2. **Explicit custom formatting**: if there is no reference and the user asks
   for a theme, brand treatment, visual style, mood, or custom formatting,
   create the deck from scratch. Do not use Grid Layout.
3. **No visual direction**: use the bundled Grid Layout Artifact.md layout
   library as the composition reference. Select and adapt layouts using the
   Grid Layout instructions above; do not run PPTX template-following mode.

User-provided references and explicit visual direction always take precedence
over Grid Layout.

For net-new or broad redesign work, ask at most three questions only when
uncertainty materially changes narrative, evidence, or visuals; otherwise
create the guided draft. Skip the loop for read-only, narrow-edit, and explicit
one-pass-final requests.

## Google Slides-Targeted Output

For a Google Slides-targeted request, first read `routing/google_slides.md`,
then create and verify a local `.pptx`. Return the verified local artifact and
state that any cloud import is a separate host step; do not claim a cloud link.

## Implementation

You MUST use `office-kit` from JavaScript ES modules to implement the slide deck.

Read the local docs before coding:

- `artifact_tool/API_QUICK_START.md`
- `artifact_tool/api/API_DOCS.md`

For parameterized DrawingML custom shapes, read
`artifact_tool/api/references/shapes.spec.md`. Use the ordered
`customAdjustments`/`customGuides` graph and DrawingML built-in or declared
guide references rather than embedding formulas in path fields. Omit custom
path `width`/`height` to use the shape-coordinate default; provide positive
literal extents only for an explicit viewport. Use ordered
`customConnectionSites` when an Agent-authored diagram needs exact native
connector anchors; bind those
shapes with explicit site indexes. Use `customAdjustmentHandles` only for the
documented bounded `xy`/`polar` handle profile, with declared adjustment names,
paired built-in/literal/guide ranges, and shape-local positions. Keep
unsupported imported handle topology opaque and fail closed. A custom shape
`textRectangle` may use pixel
edges or DrawingML built-in/declared guide names; do not invent missing guides
or bypass a failed resolved-bounds check.

Use the shape, connector, grouping, image, table, and chart references for
non-visible PowerPoint title/description/decorative metadata. Keep it distinct
from visible text, visible chart titles, and inspectable object names; preflight
imported objects with their `accessibilityCapability`. Never combine
`decorative: true` with title/description or split that classification change
across multiple calls. For an existing PPTX, prefer the source-bound
`examples/officekit-object-accessibility-edit-workflow.mjs` transaction over an
ad hoc import/mutate/export script: it consumes the audit locator directly and
proves package locality, source immutability, second import, and visual
stability for all six modeled object kinds.

For a narrow edit to a complex imported deck, inspect and resolve the exact
target before mutation. If the ordinary typed facade would rebuild or reject
unrelated native structure, request `includeNativeLeaves: true` and use only an
issued revision-bound leaf through `presentation.editNativeLeaf(...)`.
Supported leaves are existing text runs, shape RGB/local geometry, picture
local geometry, direct rich-title runs from a uniquely bound internal
ChartPart, and direct numeric bar-chart cache points proven against one exact
cell in a uniquely bound embedded XLSX. This route can edit one text run inside
a native group, a shape whose outer style remains source-owned, one issued
`chartTitleText` run without rebuilding the chart, or one issued
`chartDataValue` while changing both its ChartPart cache and workbook cell. It
does not authorize chart identity, relationships, formulas, series topology,
arbitrary workbook cells, or graphic-frame changes. The issued operation binds
the complete ownership-tree snapshot and every dependent part hash, so any
concurrent unissued change must reject. A coordinated move/resize uses one
issued call per geometry leaf and one export; the compiler sorts them into a
deterministic Edit Plan.
The resulting plan must identify one source revision and every mutation
footprint; reimport the output, prove every non-target OPC part is byte-identical,
mask only the declared XML token changes, and render the affected slide plus
unchanged comparison pages. Never replace a multi-run phrase as one
text leaf: select an exact run, or use a whole-text setter only when replacing
the complete text topology is intended.

For native charts, read `artifact_tool/api/references/charts.spec.md` before
authoring or editing. Canonical OfficeKit output covers literal bar, line,
pie, standard area, fixed 50%-hole doughnut, marker-only scatter, bounded 2D
bubble, the documented clustered bar+line combo, and bounded trendlines on
bar/line series plus one bounded native error-bar projection per bar/line
series. Use `examples/officekit-chart-families-workflow.mjs` as the
Agent-facing chart-family pattern and
`examples/officekit-chart-trendline-workflow.mjs` for trendline/error-bar
author/import/edit/reimport/render/audit. Inspect an imported ChartPart
before mutation, keep its supported topology fixed, render the final slide, and
let formula-backed custom error bars without an explicit embedded-workbook
route, external-workbook, connected, or advanced chart graphs fail closed
instead of rebuilding them from visible caches.

For slide backgrounds, use the typed `slide.setBackground(...)` and
`slide.clearBackground()` primitives documented in
`artifact_tool/api/references/slide.spec.md`. Direct solid/style-reference
backgrounds cross the canonical OfficeKit PPTX path. Never flatten an
inherited Layout/Master background or silently replace an advanced imported
background graph; preserve it unchanged or let the export fail closed.

For compound objects that must retain one ownership tree and local coordinate
space, use native `slide.groups.add(...)` and read
`artifact_tool/api/references/grouping.spec.md`. Canonical recursive groups
cross the OfficeKit PPTX path; imported topology is fixed, and complex group
shells remain one opaque read-only object rather than being flattened.

For imported deck order, `slide.moveTo(existingZeroBasedIndex)` changes only
the retained source `SlidePart` order in `p:sldIdLst`; it does not copy or
reconstruct slide graphs. Before deleting an imported slide, inspect
`slide.deletionCapability`: JS refuses an unsupported delete before changing
the model, and export independently re-proves the hash-bound source. A supported
delete removes the real SlidePart plus every exclusively owned OPC descendant,
including closed notes/comments/chart/OLE/SmartArt/InkML/media or unknown
leaves, while preserving shared layout/master/theme/image/media resources.
Inbound slide links and custom-show/section/extension identity fail closed.
Top-level imported ordinary shapes, embedded pictures, canonical connectors,
bounded tables, charts, and canonical recursive groups have a narrower element-level contract. Inspect the
facade's `deletionCapability`, then call its `.delete()` only when `supported`
is true. Shapes, connectors, and tables require a relationship-free subtree.
Picture/chart deletion removes one exact SlidePart relationship and only the
ImagePart/ChartPart descendant closure without another package parent, so
shared media and ChartParts remain. A group proves every descendant native ID,
relationship reference, and owned OPC root together; its recursive deletion
retains any part still owned outside the group. Every operation rejects
comment/connector/timing/extension consumers, requires an explicit typed
deletion intent, and is re-proved by the Codec; array splicing is never
deletion authority.
`slide.duplicate()` is a separate source-bound operation. Inspect `slide.cloneCapability` before invoking it. A supported slide is copied as an OPC ownership graph: the OfficeKit Codec recursively copies the SlidePart plus every uniquely owned OpenXmlPart and DataPart, retaining exact part bytes, content types, local relationship IDs, external relationships, and shared-node topology. It rebinds only proven shared layouts, NotesMaster, images, and retained slide-jump targets. Unknown or relationship-bearing descendants are not rejected merely because OfficeKit lacks a semantic editor for their type.

The pending JavaScript clone receives fresh slide and element identities, and connector endpoints resolve to clone-local targets. The slide, modeled elements, native-object snapshots, notes, and comments must remain unchanged until export and reimport. Custom-show membership is never extended implicitly. One pending clone per origin is allowed; origin deletion in the same transaction fails closed. Sections, modern comments, a descendant with a parent outside the owned closure, a jump to a removed slide, unresolved semantic elements or connector targets, pending native payload replacements, and graph-budget overflow fail before partial model mutation. Open XML SDK chooses collision-free part URIs, so never assume names such as `slide2.xml`; reimport and use object IDs or inspect/resolve.

After reimport, independently copied chart, embedded Office package, SmartArt, InkML, media, notes, and legacy-comment parts use their own feature-specific edit capabilities. A copied opaque part does not become semantically editable merely because its graph was preserved. Read the native-object references for those later edit boundaries.

The shipped `officekit-slide-duplicate-workflow.mjs` remains a stricter high-assurance transaction for its locked chart/OLE/SmartArt/InkML/MP4/notes/comments corpus. It performs independent type-specific oracles and render checks; it is not the limit of the public clone API and must not be used to deny a broader slide whose `cloneCapability` is supported.

For the bare agent-facing clone profile, use the shipped transaction rather
than copying ZIP parts or rebuilding the slide:

```sh
officekit run "$SKILL_DIR/examples/officekit-slide-duplicate-workflow.mjs" \
  input/source.pptx output/source-with-copy.pptx output/clone-audit.json \
  "Unique source slide name"
```

It requires exactly one explicitly named original imported slide and accepts
the closed canonical profile with no NotesSlide or legacy-comments leaf.
Recognized closed ChartParts are included without an opt-in: the workflow
proves one unique frame relationship per chart, no ChartPart child graph, a
distinct clone-local target, and byte-identical chart payload. It proves the
same independent-copy contract for every eligible embedded-XLSX OLE workbook,
including one unique inbound package edge, empty child graph, exact content
type/hash, same slide-local `r:id`, distinct clone package, and shared preview
ImagePart. It also proves every accepted SmartArt frame's exact four
`dm/lo/qs/cs` roles, relationship IDs/types, standard content types, empty
child graphs, distinct clone-local targets, and byte-identical XML. For every
accepted InkML content part it independently proves one exact `customXml`
relationship, the standard content type and root namespace, an empty child
graph, a distinct clone-local `CustomXmlPart`, and byte-identical XML. It proves the
same video/media relationship pairing, unique inbound ownership, exact
`video/mp4` bytes, distinct clone-local `MediaDataPart`, and shared immutable
poster for every accepted embedded video. It proves the
source part order, inserts one adjacent clone, keeps every retained source part
byte-identical except the required package topology records, allows only the
new SlidePart, its relationship part, and the exact cloned ChartParts, XLSX
packages, SmartArt, InkML, and MP4 parts, then checks exact source/clone
external and internal run-link relationship IDs and targets
with no orphan edge, then reimports and compares the source/clone semantics and
model render. Model
SVG comparison ignores fresh `data-*-id` locator attributes only; it is not a
claim that the clone XML is lexically byte-identical. Missing/duplicate names,
notes/comments, unresolved connector endpoints, unsupported link markup,
nonliteral or connected charts, other graph leaves, or any unexpected package
part fail closed without promoting output or audit. In particular, a nested,
incomplete, duplicated-relationship binding, mistyped, external, or connected SmartArt
graph, or a nested, extension-bearing, mistyped, non-InkML-root, ambiguous, or
connected content part, or a nested, linked, shared, non-MP4, multi-binding, or
connected media graph, is rejected before semantic import, `slide.duplicate()`,
or publication.
Read `artifact_tool/api/references/inkml-content-part-clone.spec.md` and
`artifact_tool/api/references/embedded-video-clone.spec.md` before relying on
those two high-assurance typed oracles.

The default is intentionally bare. To copy only the separately supported,
already-closed relationship leaves, opt in explicitly rather than relying on a
fallback or ZIP manipulation:

```sh
officekit run "$SKILL_DIR/examples/officekit-slide-duplicate-workflow.mjs" \
  input/source.pptx output/source-with-copy.pptx output/clone-audit.json \
  "Unique source slide name" --allow-closed-leaves
```

This opt-in accepts at most one canonical `NotesSlide` with exactly its
`NotesMaster` and back-to-source-slide relationships, and at most one canonical
legacy `SlideCommentsPart` with no child relationship graph plus the immutable
presentation-wide `CommentAuthorsPart`. The audit lists every new notes/comment
part, proves NotesSlide and comments XML are verbatim copies, proves the notes
back-reference points to the clone, and proves the immutable master/catalog are
shared. Rich/modern comments, any extra relationship, and any graph outside
that exact profile still fail closed.

For one imported canonical SmartArt document node, do not patch the ZIP or
rebuild the diagram. First inspect `nativeObject.diagramText`; it is present
only when the closed four-part graph has one through 256 direct paragraphs and
between one and 256 total plain `a:r > a:t` runs per document node. Empty
direct paragraphs remain fixed source topology and are not projected; a wholly
empty node is not editable. The returned `runs` are flattened in exact source
order; paragraph boundaries remain owned by the source package. Then run the public
transaction below, which changes only the bound DiagramDataPart and writes a
no-overwrite audit:

```sh
officekit run "$SKILL_DIR/examples/officekit-smartart-text-edit-workflow.mjs" \
  input/source.pptx output/edited.pptx output/edited.audit.json \
  "Closed SmartArt" "{B31B1833-2B65-4D6B-B3D4-9B3988427B21}" "Before" "After"
```

For a styled or multi-paragraph node, inspect `diagramText.nodes[].runs` and append
`--run-index=<zero-based-index>`. The expected and replacement text then bind
that one existing run; whole-node replacement deliberately rejects rather than
guessing a formatting boundary.

The workflow resolves exactly one object/node/optional-run/expected-text tuple,
preserves the source, verifies that no non-data package part changed, reimports
the requested node/run list, and preserves empty paragraphs plus canonical
fixed `a:br` line breaks. It fails closed for wholly empty nodes, fields,
noncanonical breaks, connected, nested, or ambiguous SmartArt. It does not add/reorder paragraphs or breaks,
nodes, or runs, change existing paragraph/run formatting, change
layout/style/colors or geometry, or claim model SVG verification is a
native-host rendering check.
Read `artifact_tool/api/references/smartart-clone.spec.md` before using either
the clone or text-edit profile.

For an original imported slide, `slide.name = "Decision review"` is a narrow
in-place metadata edit: OfficeKit changes only that SlidePart's
`p:cSld/@name`, preserves its relationship graph and all other parts, and
requires reimport for a fresh binding. It is not available for a pending clone,
which must remain an exact source copy until its export/reimport boundary.

When an imported top-level OLE object contains one uniquely bound Office
package, read `artifact_tool/api/references/ole-workbooks.spec.md` before
changing it. XLSX retains the specialized `getEmbeddedWorkbook()` and
`replaceEmbeddedWorkbook(...)` contract. The generic
`getEmbeddedOfficePackage()` and `replaceEmbeddedOfficePackage(...)` route is
compatible with that XLSX profile and currently adds exactly one source-bound
DOCX profile. It replaces validated package bytes while preserving the OLE
shell, relationship topology, preview image, and every unrelated native part.
Do not patch an embedding part directly, treat the API as arbitrary OLE access,
or present a reconstructed OLE object as equivalent. DOCX OLE frames are not
cloneable in the current bounded profile; ambiguous, shared, malformed, or
source-tampered graphs must fail closed.

For review annotations, read `artifact_tool/api/references/comments.md` before
calling `slide.comments.addThread(...)`. Canonical PPTX export supports only
bounded legacy slide-level comments with `undefined` targets, one author and
text item per annotation, and explicit coordinates. A completely comment-free
imported presentation may advertise `slide.comments.capability.addable`; that
permits creation of a canonical shared author catalog and closed slide-local
comment leaves. A closed existing legacy leaf may instead advertise
`slide.comments.capability.editable`; that permits only one existing root text
replacement while author, timestamp, coordinate, native author/index identity,
order, count, relationships, and thread topology remain fixed. Modern threads,
replies, reactions, resolved state, and element/text anchors must stay in their
native family or fail closed; never flatten them into a legacy comment.

Before running any generated presentation module, initialize its workspace so
Node.js can resolve the bundled `office-kit` package:

```bash
officekit run "$SKILL_DIR/container_tools/setup_artifact_tool_workspace.mjs" \
  --workspace "$TMP_DIR"
```

Create the ES module source file (`.mjs`) under `$TMP_DIR` and export the final
PowerPoint deck (`.pptx`) to `$FINAL_PPTX`. The generated source must be plain
JavaScript that runs directly with `node`; do not require a transpiler or build
step.

You MUST NOT use `python-pptx` or the old Python `artifact_tool` API.

### Bounded Imported Slide Name Edit

For one uniquely named original imported slide, use the shipped public
OfficeKit workflow rather than patching `ppt/slides/slide*.xml` directly:

```bash
officekit run examples/officekit-slide-name-edit-workflow.mjs \
  input.pptx output.pptx audit.json \
  "Go-no-go decision" "Go decision: controlled rollout"
```

It checks the exact source name, maps the source presentation relationship list
to the target SlidePart, changes only `slide.name`, and then proves the saved
package has the same part topology, byte-identical non-target parts, and the
requested target `p:cSld/@name`. Open XML SDK may canonicalize the target
SlidePart's XML serialization; the workflow therefore reimports, preserves the
rest of the target slide's semantics, requires a byte-identical model SVG, and
writes a source/output-bound audit. Duplicate/missing names, fallback-only
native names, unexpected package changes, pending clones, and any other
ambiguous edit fail closed. This is not a generic template metadata editor.

### Slide Show Visibility

Treat a hidden slide as playback metadata, not deletion or a custom-show edit.
Inspect `kind: "slide"`, require `visibilityCapability.known` and
`visibilityCapability.editable` for an imported PPTX, then call the typed
primitive:

```js
const slide = presentation.resolve(slideIdFromInspect);
slide.hide();       // equivalent to slide.setHidden(true)
slide.show();       // equivalent to slide.setHidden(false)
```

OfficeKit exposes `slide.hidden` but owns the native inversion: hidden writes
only `p:sld/@show="0"`, while visible clears the attribute to the schema
default. Export and reimport, verify the requested boolean, and render every
slide to prove static pixels did not change. This operation does not alter
slide order, content, custom-show membership, sections, transitions, notes,
comments, or relationships. An unknown/invalid native `@show` lexical value
reports `known: false` and fails closed rather than being guessed.

### Bounded Imported View-Properties Edit

Do not patch `ppt/viewProps.xml` or use local guide visibility as a substitute
for a file edit. For an imported deck whose `presentation.view.capability`
reports `editable: true`, use the shipped transaction:

```bash
officekit run examples/officekit-view-properties-edit-workflow.mjs \
  input.pptx output.pptx audit.json \
  '{"gridSpacingCxEmu":72000,"gridSpacingCyEmu":91440,"slideViewSnapToGrid":true,"slideViewSnapToObjects":false,"slideGuides":[{"orientation":"horizontal","position":2160},{"orientation":"vertical","position":2880}]}'
```

It keeps the original bytes immutable, requires an existing relationship-free
fixed-topology view-properties part, changes only existing grid/snap values and
guide positions, permits only `ppt/viewProps.xml` to differ, reimports and
verifies the requested semantics, proves all slide renders remain unchanged,
and writes a source/output-bound audit. It never creates view properties,
changes guide count/order/orientation, writes `showGuides`, or reconstructs an
extension/relationship graph. A missing/irregular/extended profile fails
closed. Read `artifact_tool/api/references/presentation.spec.md` before using
this path.

### Native Custom Shows

For source-free decks, create all slides and then use
`presentation.customShows.add(nameOrConfig, slides)` to author real
`p:custShowLst` playback routes. For a canonical imported list, only an
existing show's name and ordered retained-slide membership are editable; show
count/order, facade identity, and native ID remain fixed. Read
`artifact_tool/api/references/custom-shows.spec.md` before changing one.
Canonical text runs may target an existing show by exact name and may set
`returnToSlide: true|false`. OfficeKit binds that run to the show's stable
facade/native identity, so renaming the show keeps the native action and
SlidePart bytes unchanged while the next import exposes the new public name.

For one exact imported show, use the shipped transaction instead of patching
`ppt/presentation.xml`:

```bash
officekit run examples/officekit-custom-show-workflow.mjs \
  input.pptx output.pptx audit.json \
  "Board route" "Executive route" "Appendix,Overview,Appendix"
```

The workflow resolves every supplied slide name uniquely, preserves the source,
proves that only `ppt/presentation.xml` changed, retains native show identity
and all non-target shows, counts any run links bound to that fixed identity,
reimports, compares normalized visual SVG content, and writes a
source/output-bound audit. Lists with extensions, unknown children, unresolved
relationships, duplicate identities, or another noncanonical graph remain
opaque and fail closed. Missing targets and malformed, relationship-bearing,
or dangling custom-show actions fail closed. The bounded clone workflow accepts
only the canonical relationship-free run action and proves that show membership
did not change; slide deletion and custom-show topology mutation remain separate
fail-closed operations. Run LibreOffice/Poppler review after delivery when
available.

### Native PowerPoint Sections

PowerPoint sections are not custom shows: sections form the complete ordered
partition of a deck, whereas custom shows are optional playback subsets. For a
new deck, add every slide first and then define the entire partition through
`presentation.sections`:

```js
const opening = presentation.slides.add({ name: "Opening" });
const evidence = presentation.slides.add({ name: "Evidence" });
const decision = presentation.slides.add({ name: "Decision" });

presentation.sections.add("Context", [opening, evidence]);
presentation.sections.add("Decision", [decision]);
```

The export writes the native Office 2010 `p14:sectionLst` extension in
`ppt/presentation.xml`. Each section must have a unique name and at least one
slide; flattening all memberships must reproduce the current slide order
exactly, with no duplicates or omissions. Inspect an imported deck with
`presentation.inspect({ kind: "section" })`, then resolve or look up an
existing section and change only its name or boundary with `setSlides(...)`.
Canonical imports keep section count, order, public identity, and native GUIDs
fixed. Do not patch `ppt/presentation.xml` directly, add/delete/reorder an
imported section, or combine sections with slide insertion/deletion/duplicate:
those operations fail closed. Duplicate, extension-bearing, unresolved, or
otherwise irregular native section graphs are opaque-preserved and cannot be
semantically replaced. Read `artifact_tool/api/references/sections.spec.md`
before editing an imported deck, then reimport and inspect sections after
export; run native render review when available.

For one exact source-bound section-name correction, use the shipped transaction
rather than patching `ppt/presentation.xml`:

```bash
officekit run examples/officekit-section-rename-workflow.mjs \
  input.pptx output.pptx audit.json \
  "Context" "Background"
```

It requires one canonical imported section list and exactly one exact existing
name. It changes the name only: section count/order, facade ID, native GUID, and
ordered slide membership stay fixed. The source stays immutable; only
`ppt/presentation.xml` may differ; second import proves the complete section
snapshot plus non-section semantics; static model renders and `verify()` must
remain stable before no-overwrite output/audit promotion. Missing/duplicate or
case-insensitively conflicting names, section-free or opaque inputs, and any
attempt to move a boundary, add/remove/reorder a section, or pair the rename
with slide topology work fail closed. Static/native page renders prove visible
slide stability only, not PowerPoint's navigation-pane behavior.

For an imported boundary move, do not call `setSlides(...)` on one section and
infer the neighbor. Use the separate complete-partition transaction instead:

```bash
officekit run examples/officekit-section-boundary-edit-workflow.mjs \
  input.pptx output.pptx audit.json \
  @expected-sections.json \
  @replacement-sections.json
```

Both arrays must list every imported section in source order with exactly
`id`, `name`, `nativeId`, and ordered `slideIds`. The first is the exact current
snapshot. The second retains every fixed ID/name/GUID and changes membership
only as one complete ordered deck partition; a partial list, a label change,
empty group, duplicate/omitted/reordered slide, stale source, no-op, opaque
graph, or slide topology change fails closed. The transaction protects the
source, permits only `ppt/presentation.xml` to change, reimports the exact
target partition, verifies non-section semantics/static SVGs/`verify()`, and
publishes a no-overwrite audit. It makes no native navigation-pane claim.
Use `@path/to/file.json` for a large array; inline JSON remains supported and
the file form has a 32 MiB input budget.

### Bounded Slide Transitions

Use direct `p:transition` metadata only for an intentional between-slide
movement. The public profile covers the complete ECMA-376 base transition vocabulary,
plus `slow`/`medium`/`fast`, an optional bounded playback duration, click
advancement, and an optional bounded advance timer. It is not an
animation/timing/sound authoring surface:

```js
slide.setTransition({
  effect: "split",
  orientation: "horizontal",
  direction: "in",
  speed: "fast",
  durationMs: 750,
  advanceOnClick: false,
  advanceAfterMs: 4_000,
});
```

For an imported deck, inspect `slide,transition`, resolve
`${slide.id}/transition`, and read `transition.capability` before calling
`set(...)` or `clear()`. Only one existing canonical direct base-transition profile
is editable. A source-bound slide with no transition may be set only when
`addable: true`, which proves the Slide root contains only `p:cSld` plus
optional `p:clrMapOvr` and no transition, timing, or extension leaf. Unknown
or extension effects, timing trees, sound actions, non-integer-unit `p14:dur`,
or extension graphs stay opaque-preserved and fail closed on mutation. A
canonical integer-millisecond `p14:dur` is exposed as `durationMs`, distinct
from `advanceAfterMs`. The strict slide
clone profile may carry one unchanged canonical direct transition, but never a
timing or sound graph.

For one existing editable source-bound transition, use the shipped transaction
rather than a raw SlidePart patch:

```bash
officekit run examples/officekit-transition-edit-workflow.mjs \
  input.pptx output.pptx audit.json \
  "Decision slide" \
  '{"effect":"fade","throughBlack":true,"speed":"medium","durationMs":700,"advanceOnClick":true,"advanceAfterMs":1200}' \
  '{"effect":"split","orientation":"horizontal","direction":"in","speed":"slow","durationMs":1100,"advanceOnClick":false}'
```

It binds a unique imported slide name and the complete expected transition
state, requires `partPresent: true` plus `editable: true`, keeps the original
bytes immutable, permits exactly the selected SlidePart to differ, reimports,
and verifies non-transition semantics and static model renders remain stable
before it publishes a source/output-bound audit. It does not add or clear a
transition, treat an unconfigured slide as a replacement target, widen the
canonical base-transition profile, or certify native slideshow playback.

Always export, reimport, and inspect the transition again. Static
LibreOffice/Poppler review can prove the visible slide content is stable, not
slideshow playback; use a native PowerPoint playback QA lane when timing or
host effect behavior matters. Read
`artifact_tool/api/references/transitions.spec.md` before modifying imported
transition metadata.

### Rich Speaker Notes

For a source-free deck, speaker notes may use the same ordinary paragraph/run
data as slide text. Use this for an Agent's talk track, not for notes-page
design:

```js
slide.addNotes([
  {
    bulletCharacter: "•",
    runs: [
      { text: "Lead with ", style: { bold: true, fontSize: 18 } },
      { text: "the customer outcome.", style: { italic: true, fontSize: 18 } },
    ],
  },
  { bulletNone: true, runs: [{ text: "Close with the requested decision." }] },
]);
```

After importing a recognized rich NotesSlide, inspect `slide,notes`, resolve
`${slide.id}/notes`, and edit `slide.speakerNotes.textFrame.paragraphs` without
changing paragraph count, run count, or text/break kind. Assigning a new
`notes.text` string to a multi-run body would flatten it and therefore fails
closed. The legacy one-run-per-paragraph text profile still accepts full-text
replacement. Notes-local hyperlinks, fields, picture bullets, list styles,
body/layout properties, NotesMaster styling, and arbitrary notes shapes are
opaque-preserved; do not try to patch them through this API.

### Bounded Imported Rich-Notes Run Edit

For the narrow case of one known imported slide, one known title, and one known
ordinary rich-notes run, use the shipped transaction rather than replacing
`notes.text`, calling `textFrame.setText()`, or editing XML. It protects the
source, checks the exact target text and direct style, changes one fixed
paragraph/run location, then requires reimported paragraph/run topology and
every non-target run to remain exact:

```bash
officekit run examples/officekit-rich-speaker-notes-edit-workflow.mjs \
  input.pptx output/edited.pptx output/audit.json
```

The default fixture contract edits paragraph `0`, run `1` of a uniquely named
slide. Optional arguments may replace the slide/title identities and expected
title/run texts, but not widen the topology or turn this into a general
reflowing editor. Its audit records source/output hashes, target IDs,
paragraph/run indices, expected and replacement direct styles, source-bound
capability, fixed topology, second import, semantic verification, and a model
SVG check. Missing/duplicate targets, absent or irregular notes, a changed
source run/style, a topology change, or any slide/title/notes identity,
geometry, background, order, or name drift fail closed without promotion.

Use LibreOffice/Poppler after delivery: speaker notes themselves are nonvisual,
so the expected visible-slide change must come from an explicitly requested
visible edit, not an inferred notes reflow. See
`artifact_tool/api/references/speaker-notes.spec.md`.

### Bounded Imported Speaker-Notes Add

An imported slide whose source SlidePart has no NotesSlide may add plain-text
speaker notes only when `slide.speakerNotes.capability.addable` is true. Inspect
`slide,notes` first, resolve `${slide.id}/notes`, and prefer the shipped
transaction over direct OOXML relationship edits:

```bash
officekit run examples/officekit-speaker-notes-add-workflow.mjs \
  input.pptx output/with-notes.pptx output/with-notes.audit.json \
  "Unique target slide name" "Lead with the evidence.\nClose with the decision."
```

The workflow requires exactly one named imported slide with a notes-absent,
explicitly addable capability. It protects the source bytes, writes to a
temporary path, reimports, checks exact notes plus stable visible slide
semantics/order/name, compares model SVG, and audits the OPC graph. An existing
single NotesMaster is reused byte-for-byte; otherwise OfficeKit creates one
canonical NotesMaster sharing the first ordered SlideMaster's existing
ThemePart. The new NotesSlide must have exactly one NotesMaster relationship and
one back-reference to its owning SlidePart. Export independently re-proves the
source graph, so changing capability data cannot grant write authority.
Inconsistent/multiple NotesMaster graphs, unusable themes, existing/rich notes,
ambiguous slide names, and any unexpected relationship fail closed with no
output promotion. Run native LibreOffice/Poppler source-vs-output comparison
after delivery; speaker notes must not change the visible slides. See
`artifact_tool/api/references/speaker-notes.spec.md`.

### Bounded Imported Legacy Review-Comment Add

For an ordinary imported deck with no legacy or Office 2021 comments anywhere,
inspect `slide.comments.capability` before adding a review annotation. Prefer
the shipped transaction over editing `.rels`, `commentAuthors.xml`, or
`comments/comment*.xml` yourself:

```bash
officekit run examples/officekit-legacy-comment-add-workflow.mjs \
  input.pptx output/with-review.pptx output/with-review.audit.json \
  "Unique target slide name" "Confirm the imported evidence." \
  "Review Owner" "2026-07-20T03:04:05Z" 360 240
```

The workflow requires exactly one named source-bound target whose capability is
`{ format: "legacy", partPresent: false, editable: false, addable: true }`. It protects the
source, adds one slide-level annotation, exports through OfficeKit, and then
independently proves that only a canonical `CommentAuthorsPart`, one numbered
closed `SlideCommentsPart`, their two collision-free relationships, content
types, and corresponding relationship Parts changed. Slide XML, slide order,
names, and visible semantics remain unchanged. It reimports the exact author
and text, compares model SVG, emits a byte-bound audit, and uses exclusive
output publication. Native LibreOffice/Poppler source-vs-output pages must be
pixel-identical because legacy review comments are nonvisual in slideshow
rendering.

The capability is defensive preflight evidence only. OfficeKit re-proves the
complete source package and rejects a forged flag, an existing author catalog,
any legacy or modern comments part on any slide, mixed/connected comment graphs,
or a second add after reimport. Existing imported legacy comments do not become
general thread models; a separately re-proven closed leaf may use only the
text-edit workflow below. This vertical slice is canonical creation from a
comment-free source, not topology editing. See
`artifact_tool/api/references/comments.md`.

### Bounded Imported Legacy Review-Comment Text Edit

For one existing canonical legacy review annotation, first inspect the target
slide and require `comments.capability` to report
`{ sourceBound: true, format: "legacy", partPresent: true, editable: true,
addable: false }`. Select the exact stable comment ID and declare the expected
old text; never locate the target with a broad text replacement or patch
`comments/comment*.xml` directly:

```bash
officekit run examples/officekit-legacy-comment-edit-workflow.mjs \
  input-with-review.pptx output/review-text-updated.pptx output/review-text-updated.audit.json \
  "Unique target slide name" "presentation/slide/1/legacy-comment/1" \
  "Confirm the imported evidence before delivery." \
  "Confirm the imported evidence and record the delivery owner."
```

The workflow protects the input, demands a canonical closed legacy author and
comments graph, and allows only the selected root text to differ. It independently
requires the author catalog, SlidePart XML, all relationships, content types,
comment identity/index/time/coordinate, thread count/order, and every other OPC
part to stay byte-identical. Reimport checks the complete comment snapshot,
model SVG must remain identical, and the audit must report exactly one changed
existing `ppt/comments/commentN.xml`. Rich/connected or mixed-family comments,
replies, author/time/position/native-anchor edits, stale expected text, ambiguous
slide/ID selection, forged capability data, and any broader package change fail
closed with no output promotion. Run LibreOffice/Poppler source-vs-output QA
when available; review annotations must not change visible slideshow pages.

### Bounded Imported Title And Speaker-Notes Edit

For one known slide with one known text shape and a canonical plain-text Notes
part, prefer the shipped public-API/OfficeKit workflow over an ad-hoc package
patch. The title may be an ordinary editable shape or a concrete imported
SlidePart placeholder with a recognized local text body. The latter grants only
fixed-topology character replacement: native placeholder identity, geometry,
formatting, and layout binding remain source-bound. The workflow imports,
checks the exact source title and notes, changes only those two text values,
exports to a distinct path, reimports, verifies the retained slide/title/notes
identities, produces a model SVG check, and writes a byte-bound audit.

```bash
officekit run examples/officekit-title-notes-edit-workflow.mjs \
  input.pptx output.pptx audit.json
```

The optional remaining arguments are, in order: slide name, title-shape name,
expected title, replacement title, expected notes, and replacement notes. The
workflow deliberately fails closed for duplicate/missing slide or shape names,
changed expected source text, absent/rich notes, slide-name/order changes, or
any identity/geometry/direct-background change after reimport. A recognized
placeholder title must also retain its native newline/inline topology; complex
multi-run replacements and unrecognized local text graphs fail closed. It does
not claim universal template editing: SmartArt, irregular modern comment
graphs, rich notes outside the explicit fixed-topology run workflow, animations,
and other connected PresentationML graphs stay source-bound.

Run the native render/QA route after delivery when LibreOffice/Poppler is
available; the workflow's SVG check is model evidence, not a substitute for a
native-host review.

For native Office 2021 comment threads, read
`artifact_tool/api/references/comments.md` before authoring or editing. Use the
shipped workflow for a complete root/direct-reply create → import → fixed-
topology text/status edit → second import → inspect/render/audit loop:

```bash
officekit run examples/officekit-modern-comment-workflow.mjs \
  output/decision-review.pptx output/modern-comments-audit.json
```

This uses `Presentation.create({ commentFormat: "modern" })`, a top-level
element or shape-text-range anchor, independent person/GUID/time metadata, and
`thread.resolve()`/`thread.reopen()`. On imported threads only existing text and
status may change. Author/date identity, anchor and range, position, root/reply
topology, relationships, and source hashes remain fixed. Reactions/task fields,
nested replies, unknown/nested anchors, connected comment parts, and mixed
legacy/modern graphs remain opaque/source-bound and fail closed.

When one request combines an imported SmartArt node-text change, speaker-notes
editing, and a new reply in an imported modern comment thread, treat the request
as one atomic transaction. The current bounded contract cannot author that
combination: SmartArt is editable only for the separately documented canonical
plain-node profile, notes are editable only in the fixed-topology workflow, and
an imported modern thread cannot gain a new reply. Inspect before refusing:

```js
const presentation = await PresentationFile.importPptx(inputPath);
const evidence = presentation.inspectPptx();
```

Then write only `audit.json` with `status: "failed_closed"`,
`provider: { actual: "office-kit", version, silentFallback: false }`,
`savePolicy: { strategy: "none", sourceOverwriteAllowed: false,
modifiedArtifactPublished: false }`, the source path and SHA-256, an explicit
unexecuted operation for each requested edit, and
`validation: { sourceUnchanged: true, noArtifact: true }`. Include a top-level
`diagnostic` string (not only nested operation reasons) that names all four
parts of the decision: `SmartArt`, `speaker notes`, `comment reply`, and the
`source-bound`/`fail-closed` boundary. The audit must say which
SmartArt/notes/comment-reply capability caused the atomic refusal. Do not use
a custom save-policy label such as `fail-closed-no-artifact`; `none` is the
canonical strategy for a refusal. The command trace must contain both the
OfficeKit import and inspection calls plus the SmartArt, speaker-notes, and
modern-comment-reply decision. At least one actually executed shell command
must contain the literal typed calls `PresentationFile.importPptx` and
`PresentationFile.inspectPptx` (for example, an
`node --input-type=module -e '...'` command that imports the source and prints
the inspection result); mentioning those calls only in `audit.json` is not
evidence. Never flatten SmartArt, rebuild XML, edit only the supported subset,
overwrite the input, or publish a partial presentation.

Before delivery, run a local JSON assertion and repair the audit until it
passes: `status` must be `failed_closed`; `provider.actual` must be
`office-kit`; `provider.silentFallback` must be the boolean `false` (do not
substitute `fallbackUsed`); `savePolicy.strategy` must be the string `none`;
`source.sha256` must match the input; `validation.sourceUnchanged` and
`validation.noArtifact` must both be boolean `true`; and `operations` must
contain exactly three entries with `executed: false`. Missing any of these
fields is an invalid refusal, even when the prose explanation is correct.

## Template Following

Use template-following mode only when a user-provided source PPTX supplies the
layout, style, or template. Read `references/template-following.md`, use
`$TMP_DIR` from the Workspace section, and set
`TEMPLATE_PPTX="<absolute path to the user-provided PPTX>"`.

The checked-in starter command executes a validated multi-slide frame map. It
duplicates exactly one supported source slide per export/reimport boundary, so
the same source slide may be reused safely, then preflights and removes all
original slides as one source-bound ownership transaction. It renders and
verifies the result before publishing a no-overwrite starter PPTX, preview/layout
evidence, and a manifest that translates every inspected source element ID to
the corresponding final starter ID. If any clone, deletion, locator translation,
render, or verification boundary is unsupported, the command fails closed and
publishes none of those artifacts. Do not rebuild slides or share mutable parts
to bypass that refusal.

Preserve the source deck's typography, palette, spacing, layout, placeholders,
footers, page markers, and brand chrome unless the user explicitly asks to
restyle. Do not use template-following mode for a deck created from scratch.

Create:

- `$TMP_DIR/template-audit.txt`
- `$TMP_DIR/template-frame-map.json`
- `$TMP_DIR/deviation-log.txt`

Keep `$TMP_DIR/source-notes.txt` for content and asset provenance.

Inspect the complete source deck:

```bash
officekit run "$SKILL_DIR/template_following_scripts/inspect_template_deck.mjs" \
  --workspace "$TMP_DIR" \
  --pptx "$TEMPLATE_PPTX"
```

Map each output slide to an inherited source slide and identify element-level
`editTargets`. Then validate the map and build the immutable starter:

```bash
officekit run "$SKILL_DIR/template_following_scripts/validate_template_plan.mjs" \
  --workspace "$TMP_DIR" \
  --map "$TMP_DIR/template-frame-map.json"

officekit run "$SKILL_DIR/template_following_scripts/prepare_template_starter_deck.mjs" \
  --workspace "$TMP_DIR" \
  --pptx "$TEMPLATE_PPTX" \
  --map "$TMP_DIR/template-frame-map.json" \
  --out "$TMP_DIR/template-starter.pptx" \
  --preview-dir "$TMP_DIR/template-starter-preview" \
  --layout-dir "$TMP_DIR/template-starter-layout" \
  --contact-sheet "$TMP_DIR/template-starter-contact-sheet.png"
```

Create `$TMP_DIR/template-edit-plan.json` against the starter manifest. It must
bind the exact starter and manifest SHA-256 values and cover every mapped target
exactly once. Use only the typed operations documented in
`references/template-following.md`; each operation carries its old-value or
asset-hash precondition. Then apply the plan as one transaction:

```bash
officekit run "$SKILL_DIR/template_following_scripts/apply_template_edit_plan.mjs" \
  --workspace "$TMP_DIR" \
  --starter "$TMP_DIR/template-starter.pptx" \
  --manifest "$TMP_DIR/template-starter.manifest.json" \
  --plan "$TMP_DIR/template-edit-plan.json" \
  --out "$FINAL_PPTX" \
  --audit "$TMP_DIR/template-final.audit.json" \
  --preview-dir "$TMP_DIR/template-final-preview" \
  --layout-dir "$TMP_DIR/template-final-layout" \
  --contact-sheet "$TMP_DIR/template-final-contact-sheet.png"
```

The command resolves only the manifest's final `starterElementIds`, applies
bounded text/frame/table/chart/image edits plus capability-proven top-level
shape/picture/connector/table/chart/group deletion, exports and imports again, verifies, renders, rechecks
every input hash, and publishes with no overwrite. Source
inspection IDs remain provenance, not persistent identities. Unsupported
deletion/add/topology work, stale values, ambiguous targets,
unsupported run boundaries, or any render/round-trip failure publishes nothing.
If a source
slide cannot support the requested content or cannot be removed after cloning,
report the blocker and the closest viable source-slide options.

## Draft and final QA

Before showing a working draft, reopen it, inspect semantics/structure, render
every slide, and fix deterministic layout failures. Return its path, SHA-256,
and honest visual status without claiming delivery or calling `ctx.publish`.

Before delivery, rerun the complete review against the accepted draft. Inspect
each rendered slide at full size when visual input is available. Use a contact
sheet only for deck-level flow. Without visual input, run structured checks for
every slide and mark image-led or design-sensitive results
`visualReview: "requires-human"`. Request `contentView: "anydoc"` only when it
can close an identified slide-text or table content-coverage gap. The text
reading view cannot resolve layout, image, formula, metadata-provenance, or
visual-design gaps. Fix
unintended overlap, clipping, wrapping, broken connectors, unresolved
placeholders, inconsistent footers/page markers, and chart/data mismatches
before exporting. Verify that researched claims and sourced assets are
traceable, and cite sources if research was used.

## Final Response

For a working draft, return one-screen goal/structure/confirmation guidance,
suggested natural-language revisions, and its path. Do not dump slides or
expose Skill, CLI, parser, object-ID, or QA-log internals.

After explicit acceptance, return a short user-visible summary of the completed
deck. Mention the sources cited or used if research informed the deck. Do not
attach scratch plans, previews, layout JSON, or temporary assets unless the
user asks for them.

## Result and evidence

Return the final PPTX as an absolute path with `kind: "presentation"` and its
SHA-256. Include slide numbers and inspected object IDs only when they are
stable, plus render/inspect/verify evidence paths when available. Report
`visualReview: "complete"` only after the rendered slides were understood;
use `"unavailable"` or `"requires-human"` when the capability matrix requires
it. Do not emit a host-specific citation directive or cite temporary previews,
builders, or QA files as the deliverable.
