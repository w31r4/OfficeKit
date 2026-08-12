# Presentation Facade

## Create And Load

```ts
const presentation = Presentation.create({ slideSize });
const imported = Presentation.load(proto);
```

## Create Inline Type

```ts
type PresentationCreateOptions = {
  slideSize?: { width: number; height: number };
};
```

## Imported Canvas Resize

For a trusted imported PPTX, assigning `presentation.slideSize = { width,
height }` is an intentionally narrow canvas operation. OfficeKit changes only
`ppt/presentation.xml` `p:sldSz` and removes a stale preset `type`; it preserves
all existing slide, layout, master, chart, and shape coordinates. It does not
reflow, scale, crop, or reposition content. After changing the canvas, inspect,
render, and make any required layout edits explicitly.

## Presentation Slide Collection

```ts
const slide = presentation.slides.add({ layout, layoutId });
const inserted = presentation.slides.insert({ after, layout, layoutId });
const byIndex = presentation.slides.getItem(slideIndex);
slide.moveTo(destinationIndex);
const duplicate = slide.duplicate();
slide.delete();
```

## Presentation Slide Collection Inline Types

```ts
type SlideAddOptions = {
  name?: string;
  layout?: string | SlideLayoutTemplate;
  layoutId?: string;
  background?: string | BackgroundConfig;
  notes?: string | PresentationParagraph[];
};

type SlideInsertOptions = SlideAddOptions & {
  after?: Slide | number | null;
};

type PresentationParagraph = {
  runs: Array<{ text?: string; break?: boolean; style?: TextStyle }>;
  level?: number;
  alignment?: "left" | "center" | "right" | "justify";
  bulletCharacter?: string;
  autoNumber?: { type: string; startAt?: number };
  bulletNone?: boolean;
  style?: TextStyle;
};
```

`slides.add(options)` appends. `slides.insert({ after, ...options })` inserts
after an existing slide facade or its 0-based index; `after: null` inserts at
the beginning and omitting `after` appends. Both paths resolve a supplied
source-free layout transactionally and materialize its direct-frame text
placeholders. Unknown targets/layouts leave the collection unchanged.

Insertion remains source-free authoring only: inserting into an imported PPTX
would change its source-bound slide topology and is rejected at export rather
than silently reconstructing the deck.

A concrete imported SlidePart `p:sp/p:ph` with a recognized local text body may
replace existing characters through `shape.text.replace(...)` or a
newline-topology-preserving `shape.text.set(...)`. This component capability
is exposed for preflight as `shape.placeholder.textEditable === true`, but is
re-proved from the source binding during export and cannot be granted by
changing the model flag. It does not make the placeholder shape editable:
type/index, name, geometry,
formatting, layout binding, Master/Layout projections, and unmodeled XML remain
source-bound, and ambiguous topology changes fail closed.

`slide.moveTo(destinationIndex)` moves one existing slide to an existing
0-based deck index. On an imported PPTX it changes only
`ppt/presentation.xml`'s `p:sldIdLst` for the retained source SlideParts; it
neither rebuilds slide parts nor copies their relationship graphs.

`slide.delete()` returns `undefined`. It removes any non-final source-free
slide. On an imported PPTX, inspect `slide.deletionCapability` first. A
supported delete removes the actual SlidePart plus every exclusively owned OPC
descendant while retaining shared descendants. JS refuses a known-unsafe
delete before mutation; export independently re-proves the source graph.
Inbound slide references and custom-show/section/extension identity fail closed.

`slide.cloneCapability` returns defensive `{ sourceBound, known, supported, blockedReason, clonedPartCount, sharedPartCount }` evidence for an imported source SlidePart. `clonedPartCount` includes the slide root and uniquely owned OpenXmlPart descendants. `sharedPartCount` counts proven resources that the clone will rebind. This is preflight evidence only; export re-analyzes the hash-bound package.

`slide.duplicate()` returns a new adjacent `Slide` only when that capability is supported and the source semantic model is unchanged. The JavaScript layer allocates fresh slide/element identities and rebinds connector endpoints inside the copied element tree. The Codec copies the SlidePart as an OPC ownership graph: every uniquely owned OpenXmlPart, DataPart, relationship ID, content type, exact payload, external relationship, and repeated owned-node edge is retained in an independent graph. Proven shared SlideLayoutPart, NotesMasterPart, ImagePart, and retained SlidePart jump targets are rebound rather than duplicated. Open XML SDK assigns collision-free physical part URIs, so callers must not infer a clone path from its slide order.

The pending clone must cross one export/reimport boundary before any edit. One pending clone per origin is allowed, and the origin cannot be removed in the same transaction. Custom-show catalogs and membership remain unchanged. Sections, Office 2021 modern comments, any would-be owned descendant with a parent outside the closure, jumps to removed slides, unresolved semantic elements or connector targets, pending native payload replacements, and part/DataPart budget overflow fail closed before partial model mutation. A copied opaque graph remains semantically read-only unless a separate feature capability permits an edit after reimport.

The packaged duplicate workflow intentionally applies stricter chart/OLE/SmartArt/InkML/media/notes/comments oracles for its locked regression corpus. That workflow is evidence for those leaves, not a public type whitelist.

## Custom Shows

```ts
const show = presentation.customShows.add("Board route", [slide1, slide3]);
const byName = presentation.customShows.getItem("Board route");
show.name = "Executive route";
show.setSlides([slide3, slide1]);

slide1.shapes.add({
  position: { left: 80, top: 80, width: 320, height: 48 },
  text: [{ runs: [{ text: "Open route", link: { customShow: "Executive route", returnToSlide: true } }] }],
});
```

Source-free export writes a native `p:custShowLst`. A canonical imported list
permits only existing-show name and ordered membership edits; count/order,
facade IDs, and native IDs remain source-bound. See
[`custom-shows.spec.md`](./custom-shows.spec.md) for budgets, opaque graphs, and
the stable-identity run-link contract, and the audited workflow.

## PowerPoint Sections

```ts
presentation.sections.add("Context", [slide1, slide2]);
presentation.sections.add("Decision", [slide3]);

const context = presentation.sections.getItem("Context");
context.name = "Background";
context.setSlides([slide1]);
presentation.sections.getItem("Decision").setSlides([slide2, slide3]);
```

Sections are presentation-wide `p14:sectionLst` groups, not custom-show
playback subsets. Their flattened membership must be every deck slide exactly
once and in current deck order. Canonical imported sections retain their
count/order, facade IDs, and native GUIDs; only an existing name or a valid
boundary may change. Add/delete/reorder topology, pending slide clone/delete,
and opaque native extension graphs fail closed. See
[`sections.spec.md`](./sections.spec.md) for the full native/opaque boundary.

## Slide Transitions

```ts
slide.setTransition({ effect: "wheel", spokes: 6, speed: "medium", advanceOnClick: true });
slide.clearTransition();
```

The direct transition profile owns the complete ECMA-376 base effect vocabulary,
explicit slow/medium/fast speed, click advancement, and an optional bounded timer.
Inspect/resolve the stable `${slide.id}/transition` facade before changing an
imported deck. Only one existing canonical direct `p:transition` is editable;
an absent transition may be added only when `transition.capability.addable`
proves the root contains only `p:cSld` plus optional `p:clrMapOvr` and no
transition, timing, or extension leaf. Timing/sound/Office-extension/irregular
graphs remain source-bound and fail closed. See
[`transitions.spec.md`](./transitions.spec.md) for the native mapping and
playback-QA boundary.

## Discover And Edit

```ts
const snapshot = await presentation.inspect({
  kind,
  search,
  maxChars,
});

const target = presentation.resolve(anchorId);
```

`inspect` returns stable anchor ids for slides, transitions, shapes, images, tables, charts,
custom shows, sections, text ranges, speaker notes, and comment threads. `resolve` maps a
returned anchor id to the matching facade. Layout records expose `layoutId` for
search and comparison; pass only model-returned IDs such as `pr/`, `sl/`,
`custom-show/`, `sh/`, `im/`, `tb/`, `ch/`, `nt/`, `th/`, and `tr/` anchors to
`resolve`.

## Inspect Inline Type

```ts
type PresentationInspectOptions = {
  target?: { id: string; beforeLines?: number; afterLines?: number };
  kind?: string; // e.g. "slide,transition,textbox,shape,image,table,chart,notes,thread,layout"
  include?: string;
  exclude?: string;
  search?: string;
  maxChars?: number;
};
```

## Help

```ts
const help = presentation.help(query, {
  search,
  include,
  maxChars,
});
```

## Help Inline Type

```ts
type PresentationHelpOptions = {
  search?: string;
  include?: string[]; // common: ["index", "examples", "notes"]
  maxChars?: number;
};
```

## Font Inventory

`presentation.fontFamilies` returns a fresh sorted, case-insensitively
deduplicated array of explicitly used text and bullet font families. Theme
tokens such as `+mj-lt` are not reported as installed font names.

```ts
const typefaces = presentation.fontFamilies;
```

## Presentation View

Use `presentation.view` to control gridlines and imported PowerPoint guides in
an editor preview. Its visibility switches are deliberately local state, not a
file-edit API.

```ts
presentation.view.showGridlines();
presentation.view.showGuides();

const gridlinesVisible = presentation.view.gridlinesVisible;
const guidesVisible = presentation.view.guidesVisible;
const horizontalGridSpacingEmu = presentation.view.gridSpacingCxEmu;
const verticalGridSpacingEmu = presentation.view.gridSpacingCyEmu;
const snapToGrid = presentation.view.slideViewSnapToGrid;
const snapToObjects = presentation.view.slideViewSnapToObjects;
const guides = presentation.view.slideGuides;

presentation.view.hideGridlines();
presentation.view.hideGuides();

const nextGridlineState = presentation.view.toggleGridlines();
const nextGuideState = presentation.view.toggleGuides();
```

For a real imported `ppt/viewProps.xml`, inspect the explicit capability before
requesting a file edit:

```ts
const capability = presentation.view.capability;
// { sourceBound, partPresent, editable,
//   gridSpacingCxEmuPresent, gridSpacingCyEmuPresent,
//   slideViewSnapToGridPresent, slideViewSnapToObjectsPresent, guideCount }

if (!capability.editable) {
  throw new Error("This imported view-properties graph is preservation-only.");
}

presentation.view.setSourceProperties({
  gridSpacingCxEmu: 72_000,
  gridSpacingCyEmu: 91_440,
  slideViewSnapToGrid: true,
  slideViewSnapToObjects: false,
  slideGuides: [
    { orientation: "horizontal", position: 2_160 },
    { orientation: "vertical", position: 2_880 },
  ],
});
```

This is an intentionally narrow imported-PPTX edit profile, not generic view
authoring. It is available only when the source has one relationship-free
`p:viewPr` topology with an existing `p:slideViewPr/p:cSldViewPr`, at most one
fully specified `p:gridSpacing`, and an optional direct ordered list of simple
horizontal/vertical guides. The patch may change only values of attributes that
already exist and the positions of existing guides. It cannot create a
`viewProps.xml` part, add/remove/reorient guides, add snap/grid attributes,
change relationships/extensions, or reconstruct an advanced view graph.

`showGridlines()`, `showGuides()`, and their hide/toggle variants remain local
editor state. They never write `p:cSldViewPr/@showGuides`; an existing value is
preserved. During export OfficeKit independently re-proves the source part,
source binding, field presence, guide topology, and a hash of every
non-editable XML leaf. The only permitted package-byte change is
`ppt/viewProps.xml`; slides remain visually unchanged. `toProto()` keeps local
guide visibility hidden and does not expose mutable source hashes as an edit
mechanism.

## Export And Serialized Data

```ts
const imageBlob = await presentation.export({ slide, format, scale });
const montageBlob = await presentation.export({
  format: "webp",
  montage: true,
  scale: 1,
});
const layoutBlob = await slide.export({ format: "layout", scale });
const proto = presentation.toProto();
```

`toProto()` returns presentation data for host adapters. File export and local
resource resolution belong to host adapter docs.

## Export Inline Type

```ts
type PresentationExportOptions = {
  slide?: Slide;
  format?: "png" | "jpeg" | "webp" | "layout";
  width?: number;
  height?: number;
  scale?: number;
  quality?: number;
  montage?:
    | boolean
    | {
        format?: "png" | "jpeg" | "webp";
        width?: number;
        slideWidth?: number;
        padding?: number;
        gap?: number;
        background?: string;
        columns?: number;
      };
};
```

## Scripts

```ts
const result = presentation.scripts.run(scriptKind, scriptOptions);
```

Scripts provide high-level authoring recipes. Use `presentation.help(...)` to discover available script keys and option shapes.

## Cookbook

```ts
// New deck skeleton: create, set theme, add slides, render checks.
const presentation = Presentation.create({
  slideSize: { width: 1280, height: 720 },
});
presentation.theme.colorScheme = {
  name: "Clean Product",
  themeColors: {
    accent1: "#2563eb",
    accent2: "#0f766e",
    accent3: "#f59e0b",
    accent4: "#dc2626",
    accent5: "#7c3aed",
    accent6: "#16a34a",
    bg1: "#ffffff",
    bg2: "#f8fafc",
    tx1: "#0f172a",
    tx2: "#475569",
    dk1: "#000000",
    dk2: "#1e293b",
    lt1: "#ffffff",
    lt2: "#e2e8f0",
    hlink: "#2563eb",
    folHlink: "#7c3aed",
  },
};

const first = presentation.slides.add();
const second = presentation.slides.add();
const third = presentation.slides.add();

await presentation.export({
  slide: first,
  format: "png",
  scale: 1,
});
const snapshot = await presentation.inspect({
  kind: "deck,slide,textbox,chart,table",
  maxChars: 6000,
});
```

```ts
// Existing deck: inspect first, then resolve exact anchors.
const before = await presentation.inspect({
  kind: "slide,textbox,shape,image,table,chart,notes,thread,layout",
  search: "Customer growth",
  maxChars: 8000,
});
const target = presentation.resolve(anchorIdFromBefore);
```
