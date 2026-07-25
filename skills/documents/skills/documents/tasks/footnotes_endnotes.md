# True footnotes and endnotes

## Goal

Create, edit, inspect, or audit **true** DOCX footnotes/endnotes and prove that
the references, note parts, semantics, and rendered pages agree. Never emulate a
footnote with footer text or a superscript character.

## Ordinary bounded workflow: public API

Use `DocumentModel` and bundled OpenChestnut when each note has 1 through 16
physical plain-text paragraphs and is anchored at the end of one paragraph or
list item. The bounded profile permits at most one note per target block. The
first paragraph owns the native marker; later paragraphs each own one ordinary
text run. Use `text` for the backward-compatible one-paragraph shorthand, or
use `paragraphs` for an explicit multi-paragraph body.

```js
import { DocumentFile, DocumentModel } from "open-office-artifact-tool";

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
continuation-separator (`w:id=0`) entries. OpenChestnut allocates positive
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
