# PowerPoint Sections

PowerPoint sections are named, ordered groups that partition a deck. They are
not custom shows: a custom show can repeat or omit slides, while a section owns
each slide exactly once in presentation order.

OpenChestnut writes the Office 2010 extension profile in
`ppt/presentation.xml`: the presentation extension URI
`{521415D9-36F7-43E2-AB2F-B90AF26B5E84}` contains one
`p14:sectionLst`; every `p14:section` has a brace-delimited GUID and a list of
native `p:sldId/@id` values. This is a package-level grouping feature, not a
SlidePart relationship graph.

## Source-Free Authoring

Create every slide first, then declare the complete partition:

```ts
const opening = presentation.slides.add({ name: "Opening" });
const evidence = presentation.slides.add({ name: "Evidence" });
const decision = presentation.slides.add({ name: "Decision" });

presentation.sections.add({
  name: "Context",
  nativeId: "{01F07B81-39E6-4BBB-9B89-66EA253FBD29}",
  slides: [opening],
});
presentation.sections.add("Decision", [evidence, decision]);
```

Names are non-empty, at most 255 characters, and unique case-insensitively.
`nativeId` is optional for source-free authoring; the codec generates a stable,
deterministic brace-delimited GUID when it is omitted. A section needs at least
one current deck slide. At export, the flattened section memberships must
partition every deck slide exactly once and in presentation order: no omission,
duplicate, gap, or out-of-order slide is accepted.

## Inspect And Resolve

Use normal inspect/resolve before changing an imported deck:

```ts
const sections = await presentation.inspect({ kind: "section" });
const context = presentation.sections.getItem("Context");
// Or pass an inspect result's stable `section/...` id to presentation.resolve.
context.name = "Background";
context.setSlides([opening, evidence]);
presentation.sections.getItem("Decision").setSlides([decision]);
```

The public facade exposes `id`, `name`, `nativeId`, `slideIds`, and resolved
`slides`. Its `nativeId` is the source-bound native section GUID, not a promise
of persistent identity across unrelated imports.

## Imported Deck Boundary

An imported graph is editable only when it contains exactly one canonical
section extension and each entry has only the recognized attributes and native
slide IDs. The imported section count, order, public facade identity, and
native GUIDs are fixed. You may rename existing sections and move boundaries by
calling `setSlides(...)`, but the resulting groups must still be the complete
ordered partition of the retained deck.

Adding, removing, or reordering imported sections fails closed. Slide
insertion, deletion, and duplicate are separate topology operations and cannot
be combined with sections. A source-bound slide reorder must be paired with a
new valid partition before export. Any duplicate extension, unknown child or
attribute, invalid GUID/name, unresolved native slide ID, empty group, repeated
membership, or non-contiguous overall partition remains opaque and is preserved
unchanged; semantic replacement is refused rather than guessed.

## Source-Bound Name Transaction

For an Agent task that changes one imported section label only, use the shipped
workflow rather than a raw package patch:

```bash
node examples/openchestnut-section-rename-workflow.mjs \
  input.pptx output.pptx audit.json \
  "Context" "Background"
```

The transaction requires exactly one exact `expectedName` in a semantic
canonical imported list. It rejects a missing or duplicate target, an empty or
opaque section list, a missing fixed facade/native identity or membership, and
a replacement that conflicts case-insensitively with another section. It does
not call `setSlides(...)`: the count, order, facade IDs, brace GUIDs, and
ordered slide membership stay fixed.

Before it promotes distinct output and audit paths without overwrite, the
workflow proves that only `ppt/presentation.xml` changed, reimports the exact
expected full section snapshot, compares all non-section semantics and static
slide SVG renders, and runs `presentation.verify({ visualQa: true })`. Its
byte-bound audit records the target section ID/GUID/ordinal/membership and
source/output hashes. Static or LibreOffice/Poppler rendering demonstrates only
visible-slide stability; it is not a claim that the native PowerPoint
navigation pane was exercised.

## Verify And Export

Run `presentation.verify()` before export, then re-import the output and
inspect `section` again. The native edit can change only
`ppt/presentation.xml`; unrelated SlideParts should remain unchanged. For a
delivered deck, render the source and result with LibreOffice/Poppler when
available. Sections are navigation/grouping metadata, so visual equality does
not by itself prove the native section graph was retained.
