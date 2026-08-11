# Edit an existing PDF

Do not route an existing PDF through `PdfArtifact` for mutation. Pass the original file directly to the chosen provider.

## Mandatory preflight

The default MuPDF.js path probes and inspects before mutation:

```bash
officekit run scripts/mupdf.mjs probe
officekit run scripts/mupdf.mjs inspect input.pdf
```

Its typed operations are source-bound `add_text_annotation` and `add_text_highlight`, legacy text/choice/checkbox `fill_form`, source-bound `update_form_field`, source-bound `delete_page`, `duplicate_page`, and complete `rearrange_pages`, source-bound `delete_annotation` and `update_annotation`, visible-only `set_page_crop`, absolute-quarter-turn `rotate_page`, source-bound Document Info `set_metadata`, `delete_embedded_file`, source-bound `add_link`, `delete_link`, and `update_link`, `redact_text`, and `redact_rect`. Run with one explicit save policy:

```bash
officekit run scripts/mupdf.mjs edit input.pdf tmp/pdfs/edit-operations.json tmp/pdfs/edited.pdf \
  --save-policy rewrite
```

The CLI refuses source overwrite, writes atomically, and rejects incremental page-tree mutation, redaction, source-bound annotation/link creation or mutation (including text highlights), deletion, and signed-PDF incremental edits. A bounded source-bound single-widget form-field update may use unsigned incremental save; unsupported operations do not route elsewhere.

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

## Update standard Document Info and bounded canonical XMP metadata

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

For XMP-bearing input, inspect `snapshot.xmpProfile`, `xmpMutableFields`, and
`updateCapability` before constructing the patch. The built-in provider accepts
only `canonical-simple-v1`: one `x:xmpmeta`/`rdf:RDF` graph with direct
`rdf:Description` properties, a single `x-default` item for `dc:title` or
`dc:description`, a single sequence item for `dc:creator`, and direct text for
the supported `pdf:*`/`xmp:*` properties. A requested field must already appear
in `xmpMutableFields`. The transaction updates Info and that exact XMP text slot
together, then proves the decoded packet equals the inspected bytes with only
the requested slots replaced. Unknown properties therefore remain byte-exact.

Multilingual titles/descriptions, multiple authors, property attributes,
CDATA/DTD/entities outside the XML built-ins, nested descriptions, duplicate
properties, a missing requested property, direct/non-XML streams, or malformed
packets report an unsupported capability and fail closed. This is ordinary
metadata editing, not metadata sanitization; signed-file policy still applies.

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

## Update one imported text annotation

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
