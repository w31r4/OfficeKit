# Edit an existing PDF

Do not route an existing PDF through `PdfArtifact` for mutation. Pass the original file directly to the chosen provider.

## Mandatory preflight

The default MuPDF.js path probes and inspects before mutation:

```bash
officekit run scripts/mupdf.mjs probe
officekit run scripts/mupdf.mjs inspect input.pdf
```

Its typed operations are source-bound `add_text_annotation`, `add_free_text_annotation`, `add_area_annotation`, `add_text_markup` (Highlight/Underline/StrikeOut/Squiggly), and compatibility `add_text_highlight`; legacy text/choice/checkbox `fill_form`; source-bound `update_form_field`, `delete_page`, `duplicate_page`, complete `rearrange_pages`, `delete_annotation`, `update_annotation`, `set_metadata`, `update_outline`, `delete_embedded_file`, `add_link`, `delete_link`, and `update_link`; visible-only `set_page_crop`; absolute-quarter-turn `rotate_page`; `redact_text`; and `redact_rect`. Run with one explicit save policy:

```bash
officekit run scripts/mupdf.mjs edit input.pdf tmp/pdfs/edit-operations.json tmp/pdfs/edited.pdf \
  --save-policy rewrite
```

The CLI refuses source overwrite, writes atomically, and rejects incremental page-tree mutation, redaction, source-bound annotation/link creation or mutation (including every text-markup style), deletion, and signed-PDF incremental edits. Bounded source-bound single-widget form-field, metadata, outline-title/expansion, crop, and rotation updates may use unsigned incremental save; unsupported operations do not route elsewhere.

## Opaque RichMedia/3D boundary

Treat a PDF containing `/RichMedia`, `/3D`, default-view/model dictionaries, or
associated runtime scripts as a read-only preservation boundary. The current
typed routes can inventory these objects, but they cannot prove that a cover
annotation or any other mutation preserves the complete opaque closure *and*
the viewer runtime behavior. Therefore a mutation request must stop after
inspection and write only an audit record with `status: "failed_closed"` and
`savePolicy.strategy: "none"`; it must not write a reviewed PDF.

Use the configured provider interpreter and require the explicit PyMuPDF
capability probe before deciding the boundary:

```bash
PYTHON_BIN="${OFFICE_KIT_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pymupdf_edit.py probe --accept-license agpl
"$PYTHON_BIN" scripts/pdf_provider.py plan --task edit-content --provider pymupdf \
  --strategy incremental --input input.pdf --output tmp/pdfs/reviewed.pdf \
  --accept-license agpl --require-provider
```

The plan/probe is evidence for the refusal, not permission to mutate this
class of document. Record the source SHA-256, inspected RichMedia/3D closure,
actual provider/version, missing proof, `save: none`, and the unexecuted
operation in `outputs/audit.json`. Do not switch to pypdf, ReportLab, PDF.js,
content-stream patching, or a different provider as a fallback. Only a future
dedicated provider with an explicit preservation/runtime oracle may change this
contract.

## Update one imported form field

Inspect the exact input and select one `mupdfFormField` record, not a name or
array index. Copy its source hash, `id`, and full `snapshot` into the operation
file. This example is safe only when inspection returned a single-widget field:

```json
{
  "savePolicy": "incremental",
  "operations": [
    {
      "type": "update_form_field",
      "sourceSha256": "<inspect summary sourceSha256>",
      "formFieldId": "mupdf-form-field-42",
      "expected": {
        "name": "sender.city",
        "type": "text",
        "value": "",
        "readOnly": false,
        "widgets": [{
          "id": "mupdf-widget-1-42",
          "page": 1,
          "xref": 42,
          "rect": [72, 98, 180, 24]
        }]
      },
      "value": "Shanghai"
    }
  ]
}
```

The native path permits one non-password text field, one non-multiselect combo
whose inspected display and export options are identical, or one checkbox. It
checks the full snapshot before mutation and re-reads the field afterward. A
shared-widget group, radio/list/multi-select field, password field, choice
export mismatch, stale snapshot, or unknown option fails closed; use the
explicit pypdf form workflow instead. The locator is valid only for the exact
input bytes, so inspect the output before another mutation.

## Update standard Document Info and bounded field-safe XMP metadata

Treat metadata as one source object, not a loose title lookup. Select the single
`mupdfDocumentMetadata` record and copy its complete snapshot into the
operation. `patch` accepts `author`, `title`, `subject`, `keywords`, `creator`,
`producer`, `creationDate`, or `modificationDate`; use `null` to clear a field.

```json
{
  "savePolicy": "incremental",
  "operations": [{
    "type": "set_metadata",
    "sourceSha256": "<inspect summary sourceSha256>",
    "metadataId": "mupdf-document-info",
    "expected": "<complete inspect record snapshot object>",
    "patch": {
      "title": "Reviewed board packet",
      "author": "Finance operations",
      "producer": null
    }
  }]
}
```

The provider fingerprints every raw Document Info entry and, when present, the
complete decoded XMP stream. It verifies all non-target Info entries after the
mutation. A stale or partial snapshot, display-title identity, empty-string
patch, unknown key, no-op, or legacy unbound `values`/`metadata` payload fails
closed.

For XMP-bearing input, inspect `snapshot.xmpProfile`, `xmpMutableFields`,
`xmpBlockedFields`, and `updateCapability` before constructing the patch. The
built-in provider accepts `field-safe-v1`: one `x:xmpmeta`/`rdf:RDF` wrapper
with direct `rdf:Description` document properties. `dc:title` and
`dc:description` may contain multiple uniquely named languages when exactly one
`x-default` value exists. A single-item `dc:creator` sequence is mutable;
multiple creators leave `author` in `xmpBlockedFields` without blocking an
unrelated title, subject, or scalar field. `pdf:Keywords`, `pdf:Producer`, and
the supported `xmp:*` scalars may be direct text elements or ordinary
description attributes.

Every requested field must appear in `xmpMutableFields`; use the structured
reason in `xmpBlockedFields` rather than guessing around it. The transaction
updates Info and each exact XMP text or attribute slot together, then proves
the decoded packet equals the inspected bytes with only those slots replaced.
Other languages, creators, nested unknown graphs, comments, qualifiers, and
custom namespaces therefore remain byte-exact.

A duplicate or irregular standard field blocks that field. A missing requested
property also fails closed because this operation does not synthesize RDF.
CDATA, DTDs, invalid/bare entities, malformed XML, direct/non-XML streams, or a
packet with no mutable standard field make the whole metadata capability
unsupported. This is ordinary metadata editing, not metadata sanitization;
signed-file policy still applies.

## Update one existing outline title or parent expansion state

Treat bookmarks/table-of-contents entries as a source-bound tree, never as a
title lookup. Inspection flattens the native tree into ordered
`mupdfOutline` records while retaining each zero-based `path`, title, URI,
resolved 1-based page, expansion state, child count, fingerprinted `id`, and
complete `snapshot`. Select exactly one record and copy its evidence without
reconstructing it:

```json
{
  "savePolicy": "incremental",
  "operations": [{
    "type": "update_outline",
    "sourceSha256": "<inspect summary sourceSha256>",
    "outlineId": "<inspect record id>",
    "expected": "<complete inspect record snapshot object>",
    "patch": {
      "title": "Reviewed results",
      "open": false
    }
  }]
}
```

`title` must be non-empty, control-free, and no longer than 4,096 UTF-16 code
units. `open` is available only when the record has children; a leaf has no
real expansion state and fails closed. The operation preserves URI/page,
path, order, nesting, child count, and every non-target outline, then
re-inspects the complete graph. It does not add/delete/reparent entries,
change destinations, synthesize named destinations, or repair an irregular
outline graph. A stale source hash, locator, partial/tampered snapshot,
unsupported field, or no-op patch fails before save. Re-inspect the output
because the edited record receives a new fingerprinted ID. Incremental output
retains old revisions and is ordinary navigation editing, never sanitization.

## Visible page crop

Use `set_page_crop` only when the task is to change the visible page window without deleting underlying content. Inspect first and use the raw unrotated `MediaBox`/`CropBox` coordinates returned for the target page:

```json
[
  { "type": "set_page_crop", "page": 1, "bbox": [72, 72, 468, 648] }
]
```

The box must be fully inside the inspected `MediaBox`; rotated pages fail closed and need an explicitly selected specialist route. This operation writes only `CropBox`, retains content outside the crop, and may use unsigned `incremental` save. It is never a redaction, deletion, or sanitize substitute.

## Page rotation

Use `rotate_page` when the task is only to change the viewer orientation of one
existing page. It writes an absolute normalized `/Rotate` value and does not
transform, reflow, or remove content:

```json
[
  { "type": "rotate_page", "page": 2, "rotation": 90 }
]
```

`rotation` must be exactly `0`, `90`, `180`, or `270`; inspect before and after
to retain the prior value and prove the requested orientation. This bounded
unsigned operation may use `incremental` save, subject to the same source-prefix
and signature refusal rules. It is not a substitute for rotated-coordinate text
or image editing; route those tasks explicitly to the specialist provider.

## Delete or reorder imported pages

`delete_page` and `rearrange_pages` are source-bound page-tree transactions,
not loose page-number commands. Start from one fresh inspection. Deletion
requires the exact source SHA-256 and selected page bbox/rotation snapshot:

```json
{
  "savePolicy": "rewrite",
  "operations": [{
    "type": "delete_page",
    "page": 2,
    "sourceSha256": "<inspect summary sourceSha256>",
    "expectedPage": { "bbox": [0, 0, 792, 612], "rotation": 90 }
  }]
}
```

Rearrangement requires a complete permutation plus one snapshot for every
current page in its current order:

```json
{
  "savePolicy": "rewrite",
  "operations": [{
    "type": "rearrange_pages",
    "pages": [3, 1, 2],
    "sourceSha256": "<inspect summary sourceSha256>",
    "expectedPages": [
      { "page": 1, "bbox": [0, 0, 612, 792], "rotation": 0 },
      { "page": 2, "bbox": [0, 0, 792, 612], "rotation": 90 },
      { "page": 3, "bbox": [0, 0, 420, 600], "rotation": 0 }
    ]
  }]
}
```

Each operation must be the only operation in a full rewrite. A Tagged PDF, a
no-op reorder, or missing, stale, duplicated, or out-of-order evidence fails
before publication; there is no incremental or silent provider fallback.
Keep the source unchanged, re-inspect the output, prove the exact page
count/order/geometry, and render every retained source-page to output-page
mapping before delivery. Re-inspect again before a later edit because every
prior page locator is invalidated.

## Duplicate one ordinary imported page

Use `duplicate_page` only for one page in the same exact source PDF. Copy the
source hash and target `mupdfPage` bbox/rotation from a fresh inspection.
`insertAt` is a 1-based position in the resulting document; omit it to place
the duplicate immediately after the source page:

```json
{
  "savePolicy": "rewrite",
  "operations": [{
    "type": "duplicate_page",
    "page": 2,
    "sourceSha256": "<inspect summary sourceSha256>",
    "expectedPage": {
      "bbox": [0, 0, 792, 612],
      "rotation": 90
    },
    "insertAt": 4
  }]
}
```

This is deliberately a single-operation full rewrite. It accepts 0/90/180/270
degree pages only when the page has no annotations, links, widgets/form fields,
Tagged-PDF structure ownership, page actions, associated files, article beads,
transitions, or template/presentation steps. It copies the page content and
resources through MuPDF's page grafting API, but does not create a new outline,
named destination, or navigation promise. A stale hash/snapshot, interactive
page, tagged document, unsupported page-bound graph, invalid output position,
additional operation, incremental policy, or signed-document policy failure
or projected page/object budget overflow stops before publication. Re-inspect
the output and render this mapping:

- every retained original page to its shifted output page;
- the source page to its retained output page;
- the same source page to the inserted duplicate.

Require identical page geometry and Poppler pixels for each mapping. Re-inspect
again before using page numbers in a later operation because insertion changes
current-document locators.

## Remove one canonical catalog attachment entry

Use `delete_embedded_file` only for an inspected `mupdfEmbeddedFile` with
`deleteCapability.supported: true`. Copy its current-source locator and complete
snapshot; a display filename or NameTree key alone is not identity:

```json
{
  "savePolicy": "rewrite",
  "operations": [{
    "type": "delete_embedded_file",
    "sourceSha256": "<inspect summary sourceSha256>",
    "embeddedFileId": "mupdf-embedded-file-<current-source-id>",
    "expected": {
      "name": "review",
      "filename": "review.txt",
      "legacyFilename": "review.txt",
      "description": null,
      "mimeType": "text/plain",
      "declaredSize": 42,
      "fileSpecObject": 31,
      "embeddedStreamObject": 30
    }
  }]
}
```

Only a direct, unique catalog `/EmbeddedFiles` NameTree is accepted; non-target
entries are re-proved. Ambiguous or malformed graphs fail closed to pikepdf or
PyMuPDF. Rewrite, re-inspect for one fewer entry, run qpdf, and compare pages
with Poppler. `payloadErasureClaimed` and `sanitizeClaimed` stay false; full
cleanup requires sanitize.

## Add one imported Text annotation

Copy the exact inspection hash and target `mupdfPage` evidence. This is a
source-bound **pin** operation, not a rectangle/layout API: MuPDF normalizes
the native Text-note icon size. `point` is `[x, y]` in the inspected visible
`mupdfPage.bbox` coordinate space. Inspection labels it
`mupdf-page-space`: upper-left origin, y downward, with the current
0/90/180/270-degree page rotation already applied. Raw `mediaBox`/`cropBox`
remain unrotated PDF-space facts. The native footprint must fit fully inside
the visible page:

```json
[
  {
    "type": "add_text_annotation",
    "page": 2,
    "sourceSha256": "<inspect summary sourceSha256>",
    "expectedPage": {
      "bbox": [0, 0, 612, 792],
      "rotation": 0
    },
    "point": [72, 128],
    "contents": "Review this assumption.",
    "author": "Reviewer",
    "subject": "Board review"
  }
]
```

This operation is rewrite-only and supports page rotations 0/90/180/270 when
the snapshot matches. It rejects stale hashes/page geometry, out-of-window
pins, `text`/`bbox`/`rect` aliases, icon choices, empty content, clipped native
appearance, and incremental output. The operation audit contains the actual
provider-normalized annotation rectangle and a conservative `appearanceBbox`
that covers native Text-note `NoZoom`/`NoRotate` renderer differences.
Re-inspect the delivered bytes and compare that appearance before using its
fresh `mupdf-annotation-<page>-<xref>` locator for a later update or deletion.
Do not treat that xref as a persistent document identity.

## Add one visible FreeText review box

Use `add_free_text_annotation` when the review text itself must be visible on
the page rather than hidden behind a Text-note pin. Bind the same exact source
hash and `mupdfPage` snapshot, then provide `[x, y, width, height]` in the
rotation-aware `mupdf-page-space`:

```json
[
  {
    "type": "add_free_text_annotation",
    "page": 2,
    "sourceSha256": "<inspect summary sourceSha256>",
    "expectedPage": {
      "bbox": [0, 0, 612, 792],
      "rotation": 0
    },
    "bbox": [72, 128, 260, 56],
    "contents": "Review this assumption before approval.",
    "fontSize": 12,
    "textColor": [0.1, 0.2, 0.8],
    "alignment": "left",
    "author": "Reviewer",
    "subject": "Board review"
  }
]
```

The built-in profile fixes the native font resource to Helvetica, accepts a
4–72 point size, RGB text color, and `left`, `center`, or `right` alignment,
and does not expose borders, fills, rich text, callouts, or arbitrary fonts.
Before save it reads the native annotation appearance back as structured text;
if any requested text is clipped or omitted, the transaction fails and asks
for a larger box or smaller type. The full appearance must remain inside the
inspected visible page. This is rewrite-only, source-bound placement; re-open,
inspect, and render the output before delivery. The fresh FreeText xref can be
used by the generic source-bound deletion operation. A recognized
`fixed-helvetica-v1` record may also update its contents/author/subject through
the complete-snapshot route below; style and geometry remain immutable.

## Mark one visible page region

Use `add_area_annotation` when an Agent must point at an image, chart, table
region, or other content that cannot be selected as native text. Bind the exact
source and `mupdfPage` snapshot, then provide one visible bbox in
`mupdf-page-space`:

```json
[
  {
    "type": "add_area_annotation",
    "page": 2,
    "sourceSha256": "<inspect summary sourceSha256>",
    "expectedPage": { "bbox": [0, 0, 612, 792], "rotation": 0 },
    "shape": "rectangle",
    "bbox": [72, 196, 260, 96],
    "strokeColor": [0.85, 0.1, 0.1],
    "borderWidth": 3,
    "contents": "Confirm the assumptions in this region.",
    "author": "Reviewer"
  }
]
```

The bounded profile accepts exactly `rectangle` or `ellipse`, an RGB outline,
a 0.5–12 point solid border, and optional non-empty contents/author/subject. It
never fills or hides the marked content and does not expose dash, cloud, opacity,
or arbitrary appearance streams. Both the requested box and provider-reported
`appearanceBbox` must stay inside the inspected visible page, so an edge-touching
thick stroke fails closed and asks for an inset box. The operation is
rewrite-only; re-inspect and render the output. A fresh `solid-no-fill-v1`
Square/Circle may update only contents/author/subject/RGB color through the
complete-snapshot route below; geometry, border width/style, fill, and arbitrary
appearance remain fixed. Generic source-bound deletion also remains available.

## Highlight one unique imported text selection

Use a native Highlight only when the requested input text selects exactly one
native location on the inspected visible page. Selection uses the same
rotation-aware `mupdf-page-space` reported by inspection. This is deliberately
not a rectangle, quad, or generic search-and-replace API:

```json
[
  {
    "type": "add_text_highlight",
    "page": 2,
    "sourceSha256": "<inspect summary sourceSha256>",
    "expectedPage": {
      "bbox": [0, 0, 612, 792],
      "rotation": 0
    },
    "text": "Revenue assumptions remain provisional",
    "color": [1, 0.92, 0.2],
    "contents": "Validate before approval.",
    "author": "Reviewer"
  }
]
```

`text` must be non-empty and no longer than 4,096 characters. MuPDF searches
the exact page and requires exactly one hit; zero or multiple hits fail rather
than allowing an agent to guess an occurrence. The optional color is RGB in
the closed `[0,1]` interval (the default is yellow), and optional `contents`,
`author`, and `subject` must be non-empty strings. Caller-supplied quads,
rectangles, stale source/page evidence, native appearances that leave the
inspected visible bbox, and incremental output fail closed. Page rotations
0/90/180/270 are supported when the snapshot matches. The rewrite audit carries
the provider's actual quadrilateral/color/`appearanceBbox` evidence; re-inspect
and render the delivered bytes before handoff. The resulting annotation xref
is current-source-only, not a persistent document identity.

## Delete one imported annotation

First run `inspect` on the exact input. Select one `mupdfAnnotation` record by
semantic facts such as page, type, contents, author, and rectangle; never use
its array position. Copy its `id`, the inspection `summary.sourceSha256`, and a
snapshot into a rewrite operation:

```json
[
  {
    "type": "delete_annotation",
    "page": 2,
    "annotationId": "mupdf-annotation-2-42",
    "sourceSha256": "<inspect summary sourceSha256>",
    "expected": {
      "type": "Text",
      "contents": "Resolved in board review",
      "rect": [72, 128, 20, 20]
    }
  }
]
```

The source SHA-256, page encoded in the locator, xref, and every supplied
snapshot field must match before mutation. `delete_annotation` is a destructive
rewrite-only operation; incremental output is refused. Its locator is bound to
the inspected source bytes rather than a persistent document identity, so
re-inspect the output before any later annotation operation. This prevents a
rewritten PDF's xref reuse from silently targeting a different annotation.

## Update one imported Text note, fixed FreeText, area mark, or text markup

Use the same inspect-first source binding to edit the semantic text fields of
one native Text annotation without pretending to reflow or reposition it:

```json
[
  {
    "type": "update_annotation",
    "page": 2,
    "annotationId": "mupdf-annotation-2-42",
    "sourceSha256": "<inspect summary sourceSha256>",
    "expected": {
      "type": "Text",
      "contents": "Needs legal review",
      "rect": [72, 128, 20, 20]
    },
    "patch": {
      "contents": "Resolved in board review",
      "author": "Reviewer",
      "subject": "Resolved"
    }
  }
]
```

`expected` is a stale-target guard and must contain one or more supported
snapshot facts. `patch` must contain one or more non-empty `contents`,
`author`, or `subject` strings; all other patch fields fail closed. In
particular, `rect` is allowed as an expected snapshot fact but cannot be
patched: MuPDF normalizes native Text annotation geometry. For a real move or
resize, use an explicit delete-plus-add transaction with a fresh inspection, or
route to a specialist provider. `update_annotation` is rewrite-only, and the
output must be re-inspected before a subsequent annotation update or deletion.

An inspected FreeText record is editable only when `updateCapability` reports
`supported: true` and `profile: "fixed-helvetica-v1"`. Pass its complete
inspect-returned `snapshot` as `expected`, then patch only non-empty
`contents`, `author`, or `subject`. Updated contents keep the 4,096-character
and control-character limits. The provider preserves rectangle,
`appearanceBbox`, Helvetica size/color, alignment, flags, page, and locator,
and reads the rebuilt native appearance as structured text. Partial snapshots,
font/color/alignment/geometry changes, clipped or unencodable contents, no-op
patches, and incremental output fail closed.

An inspected Square/Circle is editable only when `updateCapability` reports
`supported: true` and `profile: "solid-no-fill-v1"`. Pass its complete snapshot,
then patch only non-empty `contents`/`author`/`subject` or RGB `color`. The
provider preserves type, rectangle, appearance bounds, border width/style,
no-fill state, flags, page, and locator and rechecks that the painted appearance
remains inside the visible page. Geometry, line width/style, fill/opacity,
partial snapshots, no-op patches, and incremental output fail closed.

An inspected native Highlight, Underline, StrikeOut, or Squiggly adds a complete
`snapshot` and `updateCapability`. Require that capability, pass the snapshot
unchanged as `expected`, and patch only non-empty `contents`/`author`/`subject`
or RGB `color` in `[0,1]`. The provider proves type, quadrilaterals, rectangle,
appearance bounds, flags, page, and locator unchanged before saving. Partial or
stale snapshots, no-op/color-equivalent patches, geometry fields, unsupported
annotation types, and incremental output fail closed. Re-inspect and render the
rewrite; the old locator does not identify the new byte sequence.

## Add one imported-PDF link

Copy both the input SHA-256 and the target `mupdfPage` record from the exact
inspection. `bbox` is `[x, y, width, height]` in that record's visible CropBox
coordinates, not a PDF object-array index or a viewer-relative guess:

```json
[
  {
    "type": "add_link",
    "page": 2,
    "sourceSha256": "<inspect summary sourceSha256>",
    "expectedPage": {
      "bbox": [0, 0, 612, 792],
      "rotation": 0
    },
    "bbox": [72, 128, 160, 18],
    "url": "https://example.com/current-policy"
  }
]
```

`add_link` requires the exact input hash and both the inspected page bbox and
rotation. It accepts page rotations 0/90/180/270 and a rectangle fully inside
the rotation-aware `mupdf-page-space` bbox, plus an internal `#...` destination
or absolute `http`, `https`, or `mailto` URL. Raw unrotated
`mediaBox`/`cropBox` values are not link-placement coordinates. It rejects
`javascript:`, `file:`, `data:`, a stale page snapshot, and an exact duplicate
URL/rectangle pair rather than creating an output that cannot later be selected
uniquely. It is rewrite-only and reports `coordinateSpace`/`pageRotation`.
Reopen the output and confirm the same link bounds before using its newly
generated `mupdf-link` locator.

To move a source-bound imported link, put its `delete_link` operation followed
by `add_link` in the **same** rewrite operation list. Reuse the original source
hash, use the old link's locator/snapshot for deletion, and use the original
page snapshot for addition. The deletion runs first, so the new rectangle can
reuse the old URL without creating a duplicate. This is the public replacement
for MuPDF's unstable `setBounds()` route; never mutate a link handle directly.

## Update one imported link URL

Use the same inspect-first source binding to replace one link target while
retaining its current native rectangle:

```json
[
  {
    "type": "update_link",
    "page": 2,
    "linkId": "mupdf-link-2-<inspect fingerprint>",
    "sourceSha256": "<inspect summary sourceSha256>",
    "expected": {
      "url": "https://example.com/obsolete-policy",
      "bbox": [72, 128, 160, 18],
      "external": true
    },
    "patch": {
      "url": "https://example.com/current-policy"
    }
  }
]
```

`update_link` accepts only one non-empty safe internal `#...` or absolute
`http`, `https`, or `mailto` `patch.url` field. The source
fingerprint, page, and every supplied expected fact must match before mutation.
It is rewrite-only, and the output must be re-inspected before any later link
operation. Link geometry is intentionally not patchable: the MuPDF bounds
setter's saved/reloaded coordinate semantics are not a stable public API
contract. Use the same-source-bound delete-plus-add transaction above or a
specialist provider when the rectangle needs to move.

## Delete one imported link

Select one `mupdfLink` record from the exact source inspection by page, URL,
rectangle, and externality. Never pass a mutable link-array index or URL by
itself. Copy its locator, source hash, and snapshot into a rewrite operation:

```json
[
  {
    "type": "delete_link",
    "page": 2,
    "linkId": "mupdf-link-2-<inspect fingerprint>",
    "sourceSha256": "<inspect summary sourceSha256>",
    "expected": {
      "url": "https://example.com/obsolete-policy",
      "bbox": [72, 128, 160, 18],
      "external": true
    }
  }
]
```

`delete_link` verifies every supplied fact and deletes only one unique source
link. It is rewrite-only. The link fingerprint is derived from source-visible
page/URL/rectangle/external facts, not a durable PDF object ID; a duplicate
fingerprint or any later output requires a fresh inspection and fails closed
instead of selecting by order.

## Optional PyMuPDF specialist path

For a capability outside the JavaScript contract, run the exact specialist adapter probe and route plan before any mutation. Do not defer either command until audit generation.

```bash
PYTHON_BIN="${OFFICE_KIT_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pymupdf_edit.py probe --accept-license agpl
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task edit-content \
  --provider pymupdf \
  --strategy rewrite \
  --input input.pdf \
  --output tmp/pdfs/edited.pdf \
  --accept-license agpl \
  --require-provider
```

For `replace_text` with the required `sanitize` policy, plan `--task redact --strategy sanitize --invalidate-signatures` instead. The adapter probe proves that `replace_text` is in the installed operation surface; the plan binds provider, save policy, source, destination, license, availability, and signature-invalidating acknowledgement. If either fails, stop before `pymupdf_edit.py edit`.

## Specialist operations

Prepare an operation list:

```json
[
  {
    "type": "insert_textbox",
    "page": 1,
    "rect": [72, 640, 540, 700],
    "text": "Reviewed and approved",
    "font_size": 12,
    "font_name": "helv",
    "color": [0.06, 0.36, 0.42]
  },
  {
    "type": "insert_image",
    "page": 1,
    "rect": [450, 620, 540, 710],
    "path": "tmp/pdfs/approved-mark.png",
    "keep_proportion": true
  }
]
```

Then run one explicit save policy:

```bash
"$PYTHON_BIN" scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/edited.pdf \
  --strategy rewrite \
  --operations tmp/pdfs/edit-operations.json \
  --accept-license agpl
```

Use `--strategy incremental` only for bounded changes where retaining old revisions is intended. The script copies the original bytes to the destination, requires `Document.can_save_incrementally()`, appends the update, and verifies that the original prefix is byte-identical.

`replace_image` requires an xref observed on the selected page and replaces every use of that image object. Confirm shared-object effects before delivery.

Text replacement in an ordinary PDF is not Word-style reflow. For a short replacement that fits the original geometry, use `replace_text` under `sanitize`; it requires each match to resolve to one horizontal source span, preserves its baseline and default font/size/color, performs real redaction and a same-box overlay, then runs the full residue gate. The fit check allows only a fixed sub-millipoint numerical tolerance for provider/search-box float quantization and reports the source/output style, measured width, overflow, baseline, and tolerance in `operations[].fitChecks`; it is not user-configurable layout slack. Cross-span/rotated text or replacement beyond that bound fails closed. For paragraph/page reflow, use a trusted source model or explicitly create a reconstructed new document.

Signed input requires prior signature/DocMDP inspection. Use `--allow-signed` only after the requested operation has been reviewed against the signature policy; validate before and after with pyHanko. Rewrite requires explicit `--invalidate-signatures`.

After editing, compare intended deltas, reopen independently, render every page, and preserve the source file and operation manifest.

Write the canonical [`office-kit.pdf-audit.v1`](../references/AUDIT_SCHEMA.md) envelope and validate it against the exact delivered bytes:

```bash
python3 scripts/pdf_audit.py validate outputs/audit.json \
  --source input.pdf --artifact outputs/edited.pdf \
  --require-operation replace_text
```
