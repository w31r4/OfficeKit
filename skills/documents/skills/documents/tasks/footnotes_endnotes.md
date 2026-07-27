# True footnotes and endnotes

## Goal

Create, edit, inspect, or audit **true** DOCX footnotes/endnotes and prove that
the references, note parts, semantics, and rendered pages agree. Never emulate a
footnote with footer text or a superscript character.

## Ordinary bounded workflow: public API

Use `DocumentModel` and bundled OfficeKit when each note has 1 through 16
physical plain-text paragraphs and is anchored at the end of one paragraph or
list item. The bounded profile permits at most one note per target block. The
first paragraph owns the native marker; later paragraphs each own one ordinary
text run. Use `text` for the backward-compatible one-paragraph shorthand, or
use `paragraphs` for an explicit multi-paragraph body.

```js
import { DocumentFile, DocumentModel } from "office-kit";

const document = DocumentModel.create({ name: "Research note", blocks: [] });
const claim = document.addParagraph("The pilot met its release threshold.");
const provenance = document.addParagraph("The evidence snapshot is archived.");

document.addFootnote(claim, undefined, {
  paragraphs: [
    "Pilot report, section 4.2.",
    "The independent review is retained with the release audit.",
  ],
});
document.addEndnote(provenance, "Evidence snapshot dated 2026-07-17.");

const first = await DocumentFile.exportDocx(document);
const imported = await DocumentFile.importDocx(first);
imported.notes.find((note) => note.kind === "footnote").paragraphs = [
  "Pilot report, section 4.2, independently reviewed.",
  "The independent review is retained with the release audit.",
];
const output = await DocumentFile.exportDocx(imported);
await output.save("notes.docx");
```

After import, resolve note objects again from `document.notes`, `inspect()`, or
`resolve(note.id)`. `note.paragraphs` is the physical body shape and
`note.text` is its LF-joined display projection. A recognized note may replace
the plain text in its existing paragraphs, but its count, kind, target, native
ID, reference position, formatting, and topology are source-bound. Assigning
`note.text` to an imported multi-paragraph note collapses it to one paragraph
and is rejected on export; assign the complete fixed-count `note.paragraphs`
array instead. The target paragraph/list item itself is read-only because
moving or rebuilding its reference run would change the native note graph.

## Native package contract

Footnotes and endnotes live in separate package parts:

- `word/footnotes.xml`
- `word/endnotes.xml`

The body points to them with `w:footnoteReference` or `w:endnoteReference`.
Source-free note parts include the required separator (`w:id=-1`) and
continuation-separator (`w:id=0`) entries. OfficeKit allocates positive
native IDs independently for footnotes and endnotes. A canonical body contains
1 through 16 direct `w:p` elements: the first has the appropriate marker plus
one `w:t` run; every continuation paragraph has exactly one `w:t` run.

## Inspect and audit

Use semantic inspection first:

```js
const document = await DocumentFile.importDocx(input);
console.log(document.inspect({ kind: "document,note,footnote,endnote" }).ndjson);
for (const note of document.notes) {
  console.log(document.resolve(note.id));
}
```

Then use the package reporter to inventory all references and note IDs,
including irregular graphs that remain opaque to the public model:

```bash
python scripts/footnotes_report.py input.docx
```

## Change one imported physical paragraph safely

For one recognized imported note, prefer the shipped public-model transaction
instead of an XML patch. It is intentionally narrow: it changes one existing
physical paragraph only, with a complete inspected locator and exact old text.
The input, output, and audit paths must be distinct.

1. Import and inspect the input. Record the selected note's `id`, `kind`,
   positive `nativeId`, `targetId`, `paragraphs`, and the zero-based physical
   paragraph index.
2. Run the no-overwrite workflow. The target JSON binds all source facts; use a
   trimmed, non-empty, single-line replacement. Edit a second physical
   paragraph through its own index rather than adding a line break.

```bash
officekit run examples/officekit-note-text-edit-workflow.mjs \
  input.docx output.docx audit.json \
  '{"kind":"footnote","noteId":"document/note/1","nativeId":1,"targetId":"document/block/1","paragraphIndex":0,"expectedText":"Pilot report, section 4.2."}' \
  'Pilot report, section 4.2, independently reviewed.'
```

The transaction imports through `DocumentFile`, changes only the selected
`note.paragraphs[index]`, then proves that only `word/footnotes.xml` or
`word/endnotes.xml` changed. It checks the complete native note body before
and after export, masks only the selected canonical `w:t` payload for its raw
OPC residual, reimports, verifies the complete note/body projection, model
renders, and writes source/output hashes plus its no-fallback provenance to
the audit. It fails closed for missing/mismatched IDs or text, a changed count,
rich/multi-run paragraph, unsupported note attributes, output collision, or
any package drift.

This does not move a reference, create/delete a note, alter the marker,
formatting, anchor, native ID, paragraph count, or note topology. Render all
affected pages after the transaction; note wrapping and Word-specific layout
remain host-renderer concerns.

## Explicit advanced package workflow

Use `insert_note.py` only when the requested operation is deliberately
package-level and cannot fit the bounded public model—for example, inserting at
an exact in-paragraph marker in a controlled template. It is not an automatic
fallback and must not be used to conceal an unsupported imported graph.

1. Put `[[FN]]` or `[[EN]]` at the exact controlled insertion point.
2. Patch a copy, never the source:

```bash
python scripts/insert_note.py input.docx --kind footnote --marker "[[FN]]" --text "Footnote text" --out with_fn.docx
python scripts/insert_note.py input.docx --kind endnote  --marker "[[EN]]" --text "Endnote text"  --out with_en.docx
```

3. Run `footnotes_report.py`, inspect the package, render every page, and keep
an audit record of the explicit low-level operation.

Bodies with more than 16 paragraphs, rich/multi-run paragraphs, reused
references, multiple notes on one target, custom separator/numbering/restart
graphs, anchor movement, or other irregular topologies remain opaque/source-bound
through the public codec. If the narrow helper cannot prove a safe
transformation, fail closed.

## Render and verification gate

```bash
python render_docx.py notes.docx --output_dir rendered_notes
```

Verify all of the following:

- semantic re-import contains the expected note kind, target, body text, and
  exact paragraph array;
- `document.xml` contains the expected reference IDs and the matching note
  part contains each positive ID exactly once;
- separators `-1` and `0` exist for a source-free note part;
- footnotes render at the expected page bottom and endnotes in the note section;
- numbering is unique and ordered as intended;
- long note text wraps without clipping or overlap;
- unrelated pages/content remain unchanged for an imported-file edit.

For high-stakes delivery, add Microsoft Word application validation when the
environment is available; LibreOffice rendering remains required local visual
evidence, not proof of every Word-specific behavior.
