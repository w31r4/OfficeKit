# Imported PPTX and nativeRef

Use this route for a third-party PPTX whose unknown package content must survive
an edit.

```bash
officekit ppj import input.pptx -o deck.ppj
officekit ppj inspect deck.ppj --query "visible title" --json
officekit ppj check deck.ppj --json
officekit ppj build deck.ppj -o edited.pptx --json
```

Import copies the source into a content-addressed read-only asset and binds its
SHA-256. Every visible object must appear as a typed element or an `opaque`
element with `nativeRef`, location, summary, and issued capabilities. Unknown
OOXML remains in the source package and does not enter the model context.

## Edit boundary

- Edit ordinary PPJ fields only when the typed object has the corresponding
  issued capability, such as `replaceText`, `setFrame`, `replaceImage`,
  `setChartTitle`, or `setChartData`.
- For a precise imported scalar, use an issued `nativeRef.leaves[]` entry. It
  contains an opaque leaf ID, closed `kind`, source-value `expectedHash`, and
  human-facing `value`. Change only `value`.
- Never invent a leaf, copy one to another object, alter its ID/kind/hash, or
  convert a typed operation such as image replacement into a scalar leaf.
- Keep source revision, target hash, object identity, and page identity intact.
- Re-import after build and locate the same stable IDs again.

Example imported run-size edit:

```json
{
  "id": "nl_8b1f…",
  "kind": "fontSizePoints",
  "expectedHash": "<sha256 of the exact native old value>",
  "value": 20
}
```

The value uses a human unit where PPJ defines one: points for font size,
degrees for rotation, booleans for flips, and `#RRGGBB` for direct RGB. EMU
leaves remain explicitly named `*Emu`; they exist for exact source surgery,
not for ordinary authored layout.

No-op build must return the source bytes exactly. A supported edit may change
only the target part and necessary dependencies. Unrelated parts,
relationships, master/layout/theme state, unknown timing, OLE, source-owned
SmartArt topology, and other opaque content must remain stable.

Some imported shapes with a strict direct embedded `a:blipFill` and a custom
geometry that is not yet semantically decoded may expose a source-bound
`imageFill` capability. This is a bounded frame-edit path: the existing image
relationship, crop, and custom geometry remain source-owned while the shape's
position or size changes. Image replacement, fill conversion, custom-path
rewriting, and any other image-fill graph stay opaque and fail closed.

An imported shape may also expose `textEditable` while its fill or effect graph
is intentionally opaque (for example, a native gradient banner). A plain
text-only replacement is safe when the original paragraph/run topology is
kept; the native fill, effects, geometry, and relationships remain byte-owned
by the source. Style, frame, name, and topology changes still fail closed.

Source-bound connectors may expose the same bounded placement surface when
their direct frame is a legal horizontal or vertical line with one zero
extent. Moving the frame leaves endpoint bindings, line geometry, and unknown
extension children untouched. Zero-by-zero, negative, missing, or ambiguous
connector frames remain read-only; this is not permission to rewrite connector
topology.

Source-bound opaque pictures may likewise retain a bounded negative left/top
offset when the picture frame and unique image relationship are proven safe.
This supports intentional edge bleed and crop layouts; the image payload,
effects, crop, and relationship remain source-owned.

Stale hash, ambiguous target, unsupported field, unsafe relationship change,
cross-object mutation, or topology rewrite must fail. Do not patch raw OOXML,
replace the whole slide with an image, flatten the deck, or rebuild it through
an authored route to make the request appear successful.

## Edit source-bound SmartArt text

A proven imported SmartArt frame appears as `type: "smartArt"` with
`mode: "source-bound"` instead of a generic opaque object. Edit only an
existing node's `text` after confirming that both the element and node
nativeRefs advertise `setSmartArtText` for `smartArt.text`:

```json
{
  "id": "page-brief-node-1",
  "text": {
    "paragraphs": [{
      "id": "paragraph-1",
      "runs": [
        { "id": "run-1", "text": "Revised" },
        { "id": "run-2", "text": " evidence" }
      ]
    }]
  },
  "nativeRef": { "...": "keep the complete issued value unchanged" }
}
```

One native run projects as a string; multiple formatted runs project as one
ordered PPJ run list. Change only the string values. Never add, remove,
reorder, restyle, or reparent nodes or runs. Layout, connectors, geometry,
colors, quick styles, relationship identity, and all non-data parts remain
source-owned. If the graph is not fully proven, it stays opaque and no
SmartArt text capability exists. Build to a new PPTX, then re-import before a
second edit because node IDs and capabilities are revision-bound.

## Source continuation

A complete source page may be reused only when its `nativeRef` issues
`duplicate` for `pageClone`. Insert exactly one fresh page immediately after
that source page:

```json
{
  "id": "page-source-copy",
  "role": "source continuation",
  "elements": [],
  "sourceClone": {
    "page": "page-source",
    "capability": "cap-duplicate-…"
  }
}
```

The pending page is a finite source macro, not an editable copy. Do not attach
a nativeRef, elements, layout, background, notes, visibility, transition, or
animation. Do not clone the same source twice or combine this build with page
delete/reorder or section/custom-show changes. Build creates a distinct native
SlidePart through the proven source graph copier. Re-import that PPTX before
editing the new page; only then does its full typed/opaque content and fresh
capability set exist.

To reuse one exact top-level component rather than the full page, add its
source element ID to the same finite macro:

```json
"sourceClone": {
  "page": "page-source",
  "capability": "cap-duplicate-…",
  "retainElement": "element-source-group"
}
```

Use only an ID that appears directly in the source page's `elements[]`. If the
desired object is nested, retain its owning top-level group. Every other direct
element must advertise `delete/element`; otherwise the build fails closed.
OfficeKit clones the complete native slide graph, keeps the selected object
unchanged, and removes only independently proven siblings. Re-import before
editing or adding content. The result then contains one ordinary typed or
opaque source-bound element with fresh capabilities. PPJ does not require the
retired JavaScript candidate ID or expose shape-tree indices, relationships, or
raw XML.

## Add a typed overlay to an imported page

For an ordinary imported page, the page nativeRef may issue
`appendElement/elements`. Keep every existing element unchanged and in its
original order, then append fresh typed objects to the end of `elements[]`:

```json
{
  "id": "review-label",
  "type": "text",
  "frame": { "x": 560, "y": 440, "width": 300, "height": 48 },
  "text": "Reviewed 31 Aug 2026"
}
```

The new suffix is the topmost z-order and may contain only textboxes,
`rect`/`roundRect`/`ellipse` shapes, or embedded rectangular images. New
elements use fresh IDs and no nativeRef. Do not insert them below source-owned
content, add a paired SVG fallback, or combine the append with a native edit,
deletion, reorder, page metadata change, comment, section, or custom-show
transaction. Build, render/review, and re-import before the next edit; the new
objects then appear as ordinary typed source-bound elements with fresh
nativeRefs.

OfficeKit-authored PPTX is different: if a valid embedded program exists,
import restores it exactly. If an external application changed the native file
but left the program, the embedded PPJ remains authoritative. Build a new file;
never overwrite or silently merge native drift.

Use `--task` only when immutable revisions and resume evidence are useful. Task
state does not weaken source checks or restore process memory.
