# PDF API quick start

Use `office-kit` for greenfield modeled PDF creation, semantic editing of a trusted internal model, tagged export, inspect, resolve, render, and verification. The PDF pipeline is independent from OfficeKit and does not require Microsoft Office, LibreOffice, or a local .NET SDK. Do not use this model as a fidelity-preserving mutation path for an arbitrary imported PDF.

For an executable six-page accessible report with H1-H3, a Figure alt description, meaningful Link annotation/OBJR, running Artifact text, a constrained cross-page logical Table, CJK font embedding, Poppler rendering, and separate modeled/veraPDF/human evidence, use [`../examples/accessible-board-report.mjs`](../examples/accessible-board-report.mjs).

## Startup

Use a supported Node.js runtime and resolve the installed package through standard Node.js module resolution. Work in a writable task directory and use ES modules.

```js
import {
  FileBlob,
  PdfArtifact,
  PdfFile,
  verifyArtifact,
} from "office-kit";
```

## Create an accessible PDF

```js
const pdf = PdfArtifact.create({
  metadata: { title: "Readiness report", language: "en-US" },
  pages: [{
    text: "Readiness report\nDecision evidence for the launch gate",
    width: 612,
    height: 792,
  }],
});

const decision = pdf.addText("Decision", {
  bbox: [72, 150, 180, 24],
  fontSize: 18,
  bold: true,
  headingLevel: 2,
});
const recommendation = pdf.addText("Approve the controlled rollout.", {
  bbox: [72, 180, 420, 18],
  fontSize: 12,
});
const evidence = pdf.addTable({
  id: "readiness-evidence",
  name: "readiness-evidence",
  values: [
    ["Gate", "Owner", "Status"],
    ["Model", "Artifact Platform", "Pass"],
    ["Native render", "Release QA", "Pending"],
  ],
  bbox: [72, 230, 468, 96],
});
const trend = pdf.addChart({
  name: "readiness-trend",
  title: "Readiness by gate",
  alt: "Bar chart showing readiness increasing from 76 to 96 percent.",
  chartType: "bar",
  categories: ["Model", "Package", "Render"],
  series: [{ name: "Readiness", values: [76, 90, 96], color: "#0F766E" }],
  bbox: [72, 370, 468, 180],
});

const page = pdf.pages[0];
page.setReadingOrder([
  `${page.id}/text`,
  decision,
  recommendation,
  evidence,
  trend,
]);

const report = verifyArtifact(pdf);
if (!report.ok) throw new Error(report.ndjson || JSON.stringify(report.issues));

const output = await PdfFile.exportPdf(pdf, {
  title: "Readiness report",
  language: "en-US",
});
await output.save("readiness-report.pdf");
```

Use explicit H1-H6 `headingLevel` values, meaningful figure `alt` text or `decorative: true`, semantic table cells/spans, and a complete page reading order. `verify()` detects missing or duplicate reading-order targets, invalid heading nesting, inaccessible figures, malformed tables, geometry errors, and other modeled defects.

Current model details that matter in authoring:

- Non-empty `page.text` is painted and contributes an implicit H1. Do not add another H1 unless two top-level headings are intended.
- `${page.id}/text` exists as a reading-order target only when `page.text` is non-empty.
- `addText(...)` is positioned and does not wrap. Use `addFlowText(...)` for wrapped paragraphs and automatic pagination.

## Import, inspect, and edit

PDFs exported by this package carry a clean-room model envelope and round-trip directly:

```js
const input = await FileBlob.load("readiness-report.pdf");
const pdf = await PdfFile.importPdf(input);

const inspection = pdf.inspect({
  kind: "page,text,textItem,readingOrder,table,tableCell,image,chart",
  maxChars: 20_000,
});
console.log(inspection.ndjson);

const table = pdf.pages.flatMap((page) => page.tables)
  .find((candidate) => candidate.name === "readiness-evidence");
if (!table) throw new Error("Readiness evidence table was not found.");
table.getCell(2, 2).value = "Pass";

const edited = await PdfFile.exportPdf(pdf);
await edited.save("readiness-report-final.pdf");
```

Model IDs are locators for the current object graph, not durable identities across unrelated imports. Locate targets again by bounded text, kind, name, page, or table position after importing a different file.

## Arbitrary PDF extraction and native operations

For PDFs not created by this package, MuPDF.js is the default runtime-lazy parser:

```js
const input = await FileBlob.load("third-party.pdf");
const parsed = await PdfFile.importPdf(input);

console.log(parsed.extractText());
console.log(parsed.extractTables());

const inspection = await PdfFile.inspectPdf(input);
const page = await PdfFile.renderPdf(input, { page: 1, dpi: 144 });
await page.save("third-party-page-1.png");
```

Parser-backed import reconstructs a modeled view for extraction, inspect, and QA. It is not the edit representation and must not be exported as a faithful edit. Table reconstruction is heuristic. Direct-original mutations use `PdfFile.editPdf(input, { operations, savePolicy })`; signatures still route to pyHanko, and strict sanitize/OCR or complex forms/merge route to the documented specialist tools. Inject `createPdfjsParser()` only when an independent PDF.js read adapter is specifically required.

For a bounded imported AcroForm update, inspect first and bind the exact source
plus the returned grouped `mupdfFormField` snapshot. This direct route supports
only one non-password text widget, one non-multiselect combo with identical
display/export options, or one checkbox; radio/shared/list/complex fields route
explicitly to pypdf:

```js
const inspection = await PdfFile.inspectPdf(input);
const city = inspection.records.find((record) => record.kind === "mupdfFormField"
  && record.name === "sender.city");
if (!city?.snapshot) throw new Error("Expected one inspectable city field.");

const filled = await PdfFile.editPdf(input, {
  savePolicy: "incremental",
  operations: [{
    type: "update_form_field",
    sourceSha256: inspection.summary.sourceSha256,
    formFieldId: city.id,
    expected: city.snapshot,
    value: "Shanghai",
  }],
});
await filled.save("third-party-form-filled.pdf");
```

The output gets new bytes, so source-bound locators/snapshots cannot be reused:
re-inspect before a second edit. Incremental is available only for unsigned
input and proves the original byte prefix; it never claims signed-document
permission or content sanitization.

For a bounded visible crop, take the raw page box from native inspection and edit the original bytes directly. The operation is intentionally not redaction: it changes `CropBox`, retains off-window content, and supports only unrotated pages.

```js
const pageRecord = inspection.records.find((record) => record.kind === "mupdfPage" && record.page === 1);
if (!pageRecord?.mediaBox) throw new Error("Missing native MediaBox evidence.");

const cropped = await PdfFile.editPdf(input, {
  savePolicy: "incremental",
  operations: [{ type: "set_page_crop", page: 1, bbox: [72, 72, 468, 648] }],
});
await cropped.save("third-party-page-1-cropped.pdf");
```

The requested `[x, y, width, height]` must fit fully inside the inspected raw `MediaBox`. Reopen and render the result; use a rewrite-plus-sanitize route for any task that requires actual removal of sensitive content.

For an orientation-only edit, use the same direct-original route. `rotation` is
an absolute clockwise `/Rotate` value, not a relative turn; it must be `0`,
`90`, `180`, or `270`. It changes viewer orientation without transforming or
removing content, so unsigned byte-prefix-verified incremental save is allowed:

```js
const rotated = await PdfFile.editPdf(input, {
  savePolicy: "incremental",
  operations: [{ type: "rotate_page", page: 1, rotation: 90 }],
});
await rotated.save("third-party-page-1-rotated.pdf");
```

Inspect and render the result before delivery. Rotated-coordinate text/image
editing remains an explicit specialist-provider task.

For source-bound page removal or reordering, use a fresh inspection and a
single-operation full rewrite. Deletion binds the selected page snapshot;
reordering binds every current page snapshot in current order:

```js
const sourcePages = inspection.records
  .filter((record) => record.kind === "mupdfPage")
  .map(({ page, bbox, rotation }) => ({ page, bbox, rotation }));

const reordered = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "rearrange_pages",
    pages: [3, 1, 2],
    sourceSha256: inspection.summary.sourceSha256,
    expectedPages: sourcePages,
  }],
});
await reordered.save("third-party-reordered.pdf");
```

`delete_page` uses the same source hash with
`expectedPage: { bbox, rotation }` for its selected page. Both operations reject
Tagged PDFs, incremental save, additional operations, stale/incomplete
snapshots, and an incomplete, duplicate, or no-op permutation. Keep the input
unchanged, re-inspect the output, and compare every retained source-page to its
declared output page with Poppler before reusing any page locator.

For a bounded same-document page copy, use `duplicate_page` with the exact
inspection hash and page snapshot. The optional `insertAt` is a 1-based output
position; without it the copy is inserted directly after the source page:

```js
const sourcePage = inspection.records.find((record) => record.kind === "mupdfPage"
  && record.page === 2);
if (!sourcePage) throw new Error("Expected one inspectable source page.");

const duplicated = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "duplicate_page",
    page: sourcePage.page,
    sourceSha256: inspection.summary.sourceSha256,
    expectedPage: { bbox: sourcePage.bbox, rotation: sourcePage.rotation },
    insertAt: 4,
  }],
});
await duplicated.save("third-party-with-page-copy.pdf");
```

This must be the only operation in the rewrite. It accepts a right-angle page
only when the source document is untagged and that page has no annotations,
links, widgets/forms, page actions, associated files, article beads,
transitions, or template steps. It copies visible page content/resources but
does not synthesize outlines or named destinations. Re-inspect the output and
use Poppler to compare every retained page plus the inserted page against its
declared source-page mapping. Do not reuse old page numbers afterward.

For standard imported Document Info, bind the one inspect-returned metadata
record instead of treating a title string as identity:

```js
const metadata = inspection.records.find((record) =>
  record.kind === "mupdfDocumentMetadata"
  && record.updateCapability.supported);
if (!metadata) throw new Error("Document Info is not safely editable.");

const updatedMetadata = await PdfFile.editPdf(input, {
  savePolicy: "incremental",
  operations: [{
    type: "set_metadata",
    sourceSha256: inspection.summary.sourceSha256,
    metadataId: metadata.id,
    expected: metadata.snapshot,
    patch: { title: "Reviewed board packet", author: "Finance operations" },
  }],
});
await updatedMetadata.save("third-party-with-reviewed-metadata.pdf");
```

Re-inspect and compare the complete new record. `null` clears a field; empty
strings, unknown keys, no-op/stale/partial evidence, and legacy unbound
`values` payloads fail closed. When a catalog XMP stream is present, proceed
only when `snapshot.xmpProfile === "field-safe-v1"`, the requested keys are all
listed in `snapshot.xmpMutableFields`, their names do not appear in
`snapshot.xmpBlockedFields`, and
`updateCapability.supported === true`. The same transaction updates Document
Info plus those existing XMP properties and proves the rest of the decoded XMP
packet byte-identical. A unique `x-default` is editable without changing other
languages; common scalar description attributes are editable; a multi-author
sequence or another irregular property blocks only its own field. Missing or
blocked requested fields, CDATA/DTD/invalid entities, malformed XML, or a
packet with no mutable standard field fail closed without another provider.
This is not metadata sanitization.

For one existing PDF outline/bookmark entry, use its path-bound inspection
record rather than matching a visible title:

```js
const outline = inspection.records.find((record) =>
  record.kind === "mupdfOutline"
  && record.updateCapability.mutableFields.includes("title"));
if (!outline) throw new Error("No safely editable outline entry.");

const updatedOutline = await PdfFile.editPdf(input, {
  savePolicy: "incremental",
  operations: [{
    type: "update_outline",
    sourceSha256: inspection.summary.sourceSha256,
    outlineId: outline.id,
    expected: outline.snapshot,
    patch: { title: "Reviewed section" },
  }],
});
await updatedOutline.save("third-party-with-reviewed-outline.pdf");
```

The complete snapshot binds path, title, URI, expansion state, resolved page,
and child count. `open` may be patched only for a parent. URI/page, ordering,
nesting, and child count remain fixed, every non-target entry is re-proven,
and the output must be re-inspected because the edited record receives a new
fingerprinted ID. Missing/stale/partial evidence, leaf expansion, control
characters, no-op patches, destination changes, and topology edits fail closed
without another provider. Incremental history is not sanitization.

For a canonical catalog attachment, use its current inspection record, never a
filename:

```js
const attachment = inspection.records.find((record) =>
  record.kind === "mupdfEmbeddedFile"
  && record.name === "review"
  && record.deleteCapability.supported);
if (!attachment) throw new Error("No safely removable catalog attachment entry.");

const withoutAttachmentEntry = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "delete_embedded_file",
    sourceSha256: inspection.summary.sourceSha256,
    embeddedFileId: attachment.id,
    expected: attachment.snapshot,
  }],
});
await withoutAttachmentEntry.save("third-party-without-review-entry.pdf");
```

Re-inspect for one fewer entry, run qpdf, and compare all pages. Ambiguous or
stale graphs fail closed. This removes a catalog NameTree entry only, not its
payload or a sanitize claim; full cleanup uses pikepdf or PyMuPDF.

For a new Text review note on an imported PDF, bind the exact source hash and
target page snapshot first. This is a pin, not a rectangle API: MuPDF owns the
native icon's normalized size and returns its actual rectangle in the audit:

```js
const notePage = inspection.records.find((record) => record.kind === "mupdfPage"
  && record.page === 1);
if (!notePage) throw new Error("Expected one inspectable target page.");

const withReviewNote = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "add_text_annotation",
    page: notePage.page,
    sourceSha256: inspection.summary.sourceSha256,
    expectedPage: { bbox: notePage.bbox, rotation: notePage.rotation },
    point: [72, 128],
    contents: "Review this assumption.",
    author: "Reviewer",
  }],
});
await withReviewNote.save("third-party-with-review-note.pdf");
```

The pin is measured in the inspected page's explicit `mupdf-page-space`: an
upper-left origin with y increasing downward and the current 0/90/180/270-degree
rotation already applied. Raw `mediaBox`/`cropBox` records are separate
unrotated PDF-space facts and are not placement coordinates. The operation
must fit both its requested 20-point anchor rectangle and the provider-reported
conservative `appearanceBbox` inside `mupdfPage.bbox`; `text`, `bbox`, `rect`,
icon selection, stale evidence, clipped appearance, and incremental save are
rejected. Re-inspect the rewrite and compare the fresh annotation
`appearanceBbox` before relying on its current-source-only locator.

For visible review text, use the separate bounded FreeText primitive. It binds
the same source/page evidence but takes one `[x, y, width, height]` box in
`mupdf-page-space` and verifies that the native appearance contains all
requested text before publication:

```js
const withVisibleReview = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "add_free_text_annotation",
    page: notePage.page,
    sourceSha256: inspection.summary.sourceSha256,
    expectedPage: { bbox: notePage.bbox, rotation: notePage.rotation },
    bbox: [72, 128, 260, 56],
    contents: "Review this assumption before approval.",
    fontSize: 12,
    textColor: [0.1, 0.2, 0.8],
    alignment: "left",
    author: "Reviewer",
  }],
});
await withVisibleReview.save("third-party-with-visible-review.pdf");
```

This rewrite-only profile fixes the native font resource to Helvetica,
supports 4–72 point text and left/center/right alignment, and rejects clipped
text, off-page appearances, borders, fills, rich text, callouts, arbitrary
fonts, and stale evidence. Re-inspect and render the output; FreeText updates
are accepted only when the fresh record advertises
`updateCapability.profile === "fixed-helvetica-v1"`. Pass its complete
`snapshot` to `update_annotation` and patch only `contents`, `author`, or
`subject`:

```js
const revisedInspection = await PdfFile.inspectPdf(withVisibleReview);
const review = revisedInspection.records.find(
  (record) => record.kind === "mupdfAnnotation" && record.type === "FreeText",
);
if (!review?.updateCapability.supported) throw new Error("FreeText update is not supported.");

const revisedReview = await PdfFile.editPdf(withVisibleReview, {
  savePolicy: "rewrite",
  operations: [{
    type: "update_annotation",
    page: review.page,
    sourceSha256: revisedInspection.summary.sourceSha256,
    annotationId: review.id,
    expected: review.snapshot,
    patch: { contents: "Updated review: ready for approval." },
  }],
});
```

The update preserves the rectangle, Helvetica size/color, alignment, flags,
and locator within that source, then returns `appearanceTextVerified: true`
after re-reading the native appearance text.
Clipped or provider-unencodable contents fail closed. Font, color, alignment,
geometry, border/fill, rich-text, and callout updates remain unsupported; the
generic source-bound delete route remains available after a fresh inspection.

For an image, chart, table area, or other non-text review target, add an unfilled
native region outline instead of inventing text quadrilaterals:

```js
const areaReview = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "add_area_annotation",
    page: notePage.page,
    sourceSha256: inspection.summary.sourceSha256,
    expectedPage: { bbox: notePage.bbox, rotation: notePage.rotation },
    shape: "ellipse",
    bbox: [72, 210, 260, 96],
    strokeColor: [0.85, 0.1, 0.1],
    borderWidth: 3,
    contents: "Confirm this figure before approval.",
    author: "Reviewer",
  }],
});
```

`shape` is exactly `rectangle` or `ellipse`; the native result is Square or
Circle with a solid 0.5–12 point RGB outline and no interior color. Unknown
fill/dash/cloud/opacity/appearance fields fail closed. The requested bbox and
the provider-reported `appearanceBbox` must both fit inside the inspected
rotation-aware page, so a thick stroke at the page edge is refused. Re-inspect
and render the rewrite. Region style/geometry updates remain unsupported, but a
fresh complete annotation snapshot may use the generic delete operation.

For review markup, give the provider one requested text string instead of
trying to calculate a rectangle or character quadrilaterals. It is accepted
only if native search finds exactly one selection on the same inspected visible
page. The provider uses that page's rotation-aware `mupdf-page-space`:

```js
const highlightPage = inspection.records.find((record) => record.kind === "mupdfPage"
  && record.page === 1);
if (!highlightPage) throw new Error("Expected one inspectable target page.");

const withMarkup = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "add_text_markup",
    markup: "strikeout",
    page: highlightPage.page,
    sourceSha256: inspection.summary.sourceSha256,
    expectedPage: { bbox: highlightPage.bbox, rotation: highlightPage.rotation },
    text: "Revenue assumptions remain provisional",
    color: [0.9, 0.2, 0.2],
    contents: "Validate before approval.",
    author: "Reviewer",
  }],
});
await withMarkup.save("third-party-with-review-markup.pdf");
```

`markup` is exactly `highlight`, `underline`, `strikeout`, or `squiggly`;
`add_text_highlight` remains the compatibility form of `highlight`. `text` is
non-empty and at most 4,096 characters. The optional RGB color uses three
`[0,1]` components (yellow for highlight, red otherwise), while optional
`contents`, `author`, and `subject` carry non-empty review metadata. Caller
quads or rectangles, zero/multiple native hits, stale page evidence, a native
`appearanceBbox` beyond the visible page, and incremental save are rejected.
Right-angle page rotation itself is supported and must match `expectedPage`.
Re-inspect and render the rewrite before handoff; its `mupdfAnnotation` record
returns the native type, quadrilaterals, color, appearance, and a
current-source-only locator.

For an imported annotation, do not use its array index as identity. Inspect the
exact input bytes, retain the returned `summary.sourceSha256`, and delete only
one source-bound annotation locator with a semantic precondition. This is a
rewrite-only operation because a deletion must not leave the original object in
an incremental revision:

```js
const annotation = inspection.records.find((record) =>
  record.kind === "mupdfAnnotation"
  && record.page === 2
  && record.type === "Text"
  && record.contents === "Resolved in board review"
);
if (!annotation?.id || !inspection.summary.sourceSha256) {
  throw new Error("The target annotation was not uniquely inspectable.");
}

const withoutReviewNote = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "delete_annotation",
    page: annotation.page,
    annotationId: annotation.id,
    sourceSha256: inspection.summary.sourceSha256,
    expected: {
      type: annotation.type,
      contents: annotation.contents,
      rect: annotation.rect,
    },
  }],
});
await withoutReviewNote.save("third-party-without-review-note.pdf");
```

`mupdf-annotation-<page>-<xref>` is a locator for these exact source bytes,
not a durable annotation identity. Re-inspect after every rewrite: MuPDF may
renumber or reuse xrefs. A mismatched source hash, page, locator, or expected
snapshot fails closed before output is written.

To update one imported Text annotation instead of deleting it, use the same
inspection evidence. The bounded patch changes only non-empty `contents`,
`author`, and `subject`; it does not claim arbitrary annotation editing or
layout reflow:

```js
const revisedReviewNote = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "update_annotation",
    page: annotation.page,
    annotationId: annotation.id,
    sourceSha256: inspection.summary.sourceSha256,
    expected: {
      type: annotation.type,
      contents: annotation.contents,
      rect: annotation.rect,
    },
    patch: {
      contents: "Resolved in board review",
      author: "Reviewer",
      subject: "Resolved",
    },
  }],
});
await revisedReviewNote.save("third-party-review-note-updated.pdf");
```

For an existing native Highlight, Underline, StrikeOut, or Squiggly, require
`annotation.updateCapability.supported` and pass the complete inspect-returned
`annotation.snapshot`. Its patch may additionally replace `color` with three
RGB components in `[0,1]`:

```js
const markup = inspection.records.find((record) =>
  record.kind === "mupdfAnnotation"
  && record.type === "Underline"
  && record.contents === "Validate before approval."
);
if (!markup?.updateCapability.supported) {
  throw new Error("The target text markup is not safely mutable.");
}

const revisedMarkup = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "update_annotation",
    page: markup.page,
    annotationId: markup.id,
    sourceSha256: inspection.summary.sourceSha256,
    expected: markup.snapshot,
    patch: {
      contents: "Reviewed and accepted.",
      subject: "Resolved",
      color: [0.15, 0.55, 0.25],
    },
  }],
});
await revisedMarkup.save("third-party-markup-updated.pdf");
```

Text-markup update never changes type, quadrilaterals, rectangle, appearance
bounds, flags, page, or locator. A partial/stale snapshot, no-op or invalid RGB,
geometry patch, unsupported annotation type, or incremental save fails closed.
Re-inspect and render the rewrite before another operation.

The rectangle can guard the source snapshot but cannot appear in `patch`.
MuPDF normalizes native Text annotation geometry, so moving or resizing a note
must be an explicit delete-plus-add transaction from a fresh inspection, or a
specialist-provider task. Re-inspect this rewrite before any later annotation
mutation; the old xref locator is not a persistent identity.

For an imported link, use the same inspect → source-bound locator → snapshot
pattern. Do not select a link by mutable page-array index or URL alone: several
links may share a target URL. A duplicate semantic fingerprint fails closed
rather than choosing one arbitrarily.

```js
const link = inspection.records.find((record) =>
  record.kind === "mupdfLink"
  && record.page === 2
  && record.url === "https://example.com/obsolete-policy"
);
const linkPage = inspection.records.find((record) =>
  record.kind === "mupdfPage" && record.page === link?.page
);
if (!link?.id || !linkPage || !inspection.summary.sourceSha256) {
  throw new Error("The target link was not uniquely inspectable.");
}

const withoutObsoleteLink = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "delete_link",
    page: link.page,
    linkId: link.id,
    sourceSha256: inspection.summary.sourceSha256,
    expected: {
      url: link.url,
      bbox: link.bbox,
      external: link.external,
    },
  }],
});
await withoutObsoleteLink.save("third-party-without-obsolete-link.pdf");
```

To add a new link, bind its rectangle to the same inspected page geometry.
`add_link` uses the inspected rotation-aware `mupdf-page-space`, supports page
rotations 0/90/180/270, and accepts only internal `#...` or absolute `http`,
`https`, and `mailto` destinations:

```js
const withCurrentPolicyLink = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "add_link",
    page: linkPage.page,
    sourceSha256: inspection.summary.sourceSha256,
    expectedPage: { bbox: linkPage.bbox, rotation: linkPage.rotation },
    bbox: [72, 128, 160, 18],
    url: "https://example.com/current-policy",
  }],
});
await withCurrentPolicyLink.save("third-party-current-policy-link.pdf");
```

The rectangle must fit fully within the inspected `mupdfPage.bbox`; do not use
raw unrotated `mediaBox`/`cropBox` coordinates on a rotated page. The operation
audit reports `coordinateSpace` and `pageRotation`, and a fresh inspection must
reproduce the added bounds before delivery.

To replace that same link's target without moving its native rectangle, use the
same locator and snapshot with a URL-only patch:

```js
const updatedPolicyLink = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "update_link",
    page: link.page,
    linkId: link.id,
    sourceSha256: inspection.summary.sourceSha256,
    expected: {
      url: link.url,
      bbox: link.bbox,
      external: link.external,
    },
    patch: { url: "https://example.com/current-policy" },
  }],
});
await updatedPolicyLink.save("third-party-policy-link-updated.pdf");
```

The patch must contain exactly one non-empty safe internal `#...` or absolute
`http`, `https`, or `mailto` `url`; a link rectangle is snapshot evidence, not
mutable geometry. MuPDF's bounds setter does not provide a stable
saved/reloaded coordinate contract for this public API, so move a link through
an explicit `delete_link` + `add_link` transaction in one rewrite using the
same original link/page evidence, or use a specialist provider. Re-inspect the
rewrite before any later link operation.

`mupdf-link-<page>-<fingerprint>` is source-byte-bound, not a persistent link
identity. It has no native xref because the MuPDF link API abstracts that
object away; its fingerprint covers page, URL, rectangle, and externality. A
new output must be re-inspected before a later link operation.

## Sign and validate exact bytes

Signature mutation and trust are outside the JavaScript model. Use the shipped
pyHanko signer for one bounded local-PKCS#12 incremental revision, then use the
separate validator with a fresh output hash and explicit trust policy:

```bash
PYTHON_BIN="${OFFICE_KIT_PDF_PROVIDER_PYTHON:-python3}"
SOURCE_SHA256="$(shasum -a 256 input.pdf | awk '{print $1}')"
CREDENTIAL_SHA256="$(shasum -a 256 /secure/signer.p12 | awk '{print $1}')"
"$PYTHON_BIN" scripts/pyhanko_sign_provider.py sign \
  input.pdf tmp/pdfs/signed.pdf \
  --expected-sha256 "$SOURCE_SHA256" --trusted-input \
  --credential /secure/signer.p12 \
  --credential-sha256 "$CREDENTIAL_SHA256" --passphrase-stdin \
  --field-name Approval --field-mode create-invisible \
  --signature-kind approval --expected-signature-count 0 \
  > tmp/pdfs/signing-report.json
# Hidden terminal prompt; automation pipes stdin directly from its secret manager.
SIGNED_SHA256="$(shasum -a 256 tmp/pdfs/signed.pdf | awk '{print $1}')"
"$PYTHON_BIN" scripts/pyhanko_provider.py verify tmp/pdfs/signed.pdf \
  --expected-sha256 "$SIGNED_SHA256" \
  --trust-policy explicit-roots \
  --trust-root /trusted/root-ca.pem \
  --require-signature --require-all-integrity-valid \
  --require-all-trusted --require-docmdp-compliant \
  --require-all-bottom-line \
  > tmp/pdfs/signature-validation.json
```

The signer keeps the source prefix intact, adds one signature, consumes the
credential secret only on stdin, and validates integrity/DocMDP before
collision-safe promotion. The validation result keeps ByteRange integrity, certificate trust, timestamps,
post-signing changes, and DocMDP/FieldMDP evidence separate. It disables network
fetching and implicit system trust, never mutates the PDF, and does not claim
complete PAdES profile conformance. Re-run it after every incremental revision.
See the Skill's [sign and verify](../tasks/sign_verify.md) task for dependency,
revocation, signing, and key-handling boundaries.

## Validate PDF/A or PDF/UA machine rules on exact bytes

Conformance validation is outside the JavaScript model. Resolve the explicit
veraPDF provider first (a verified managed pack when published or a declared
`system-only` runtime), then use the shipped thin adapter with a fresh
final-file hash and one explicit built-in profile:

```bash
export OFFICE_KIT_PDF_VERAPDF="/absolute/path/to/verapdf"
PYTHON_BIN="${OFFICE_KIT_PDF_PROVIDER_PYTHON:-python3}"
SOURCE_SHA256="$(shasum -a 256 output.pdf | awk '{print $1}')"
"$PYTHON_BIN" scripts/verapdf_provider.py validate output.pdf \
  --expected-sha256 "$SOURCE_SHA256" --flavour ua1 --require-compliant \
  > tmp/pdfs/verapdf-ua1.json
```

The result distinguishes a completed validation operation from machine-rule
compliance. `--require-compliant` turns a false result into a delivery failure;
PDF/UA still requires human judgment even when the machine rules pass. See the
Skill's [accessibility task](../tasks/accessibility.md) for supported profiles
and review boundaries.

## Render and visual QA

Use the model SVG preview while authoring, then render the exported PDF with Poppler and inspect every page before delivery:

```js
import { createPopplerRenderer } from "office-kit/renderers/poppler";

const preview = await pdf.render({ pageIndex: 0 });
await preview.save("readiness-report-preview.svg");

const renderer = createPopplerRenderer({ dpi: 144, timeoutMs: 60_000 });
for (let pageIndex = 0; pageIndex < pdf.pages.length; pageIndex += 1) {
  const png = await pdf.render({
    source: "pdf",
    format: "png",
    pageIndex,
    renderer,
  });
  await png.save(`readiness-report-page-${pageIndex + 1}.png`);
}
```

`PdfFile.inspectPdf(...)` verifies byte-level evidence such as PDF version, page/object counts, EOF, tagged structure, reading-order IDs, headings, Figure alternative text, Table/TR/TH/TD roles, spans, and font evidence. It complements semantic `pdf.verify()` and visual page review; none of the three replaces the others.

Use `pdftoppm`/`pdfinfo` as independent native render and file QA tools. The surrounding PDF Skill defines the MuPDF.js, ReportLab, pdfplumber, pypdf, PyMuPDF, pyHanko, veraPDF, and OCR routing contract; its capability resolver selects a managed or `system-only` provider without generic flag passthrough or fallback. The thin adapters then provide bounded source-bound qpdf structure operations, pyHanko local-PKCS#12 signing and independent signature validation, and veraPDF conformance validation.
