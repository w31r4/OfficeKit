# Forms and annotations

Use MuPDF.js for bounded source-bound single-widget text/combo/checkbox updates, Text-note pins, and unique native text markup (Highlight, Underline, StrikeOut, or Squiggly). Use pypdf when radio export values, shared widgets, choice display/export mappings, appearance-state validation, flattening, or more complex AcroForm handling is required. Always open the original PDF directly.

## Inspect first

```bash
python3 scripts/pypdf_edit.py inspect input.pdf \
  --output tmp/pdfs/pypdf-inspect.json
```

Check field hierarchy, widget pages, current values, annotations, encryption, signatures, and DocMDP before mutation.

For a supported MuPDF.js field or text note:

```bash
officekit run scripts/mupdf.mjs probe
officekit run scripts/mupdf.mjs inspect input.pdf
officekit run scripts/mupdf.mjs edit input.pdf tmp/pdfs/form-operations.json tmp/pdfs/filled.pdf \
  --save-policy rewrite
```

Native inspection emits individual `mupdfWidget` records and groups them into
`mupdfFormField` records. For an agent-safe direct field update, select one
field record by semantic name/type/value, then copy **both** the inspection
`summary.sourceSha256` and that record's `id`/`snapshot`. Do not select by
array position or field name alone:

```js
const inspection = await PdfFile.inspectPdf(input);
const field = inspection.records.find((record) => record.kind === "mupdfFormField"
  && record.name === "sender.city");
if (!field?.snapshot) throw new Error("Expected one inspectable city field.");

const edited = await PdfFile.editPdf(input, {
  savePolicy: "incremental",
  operations: [{
    type: "update_form_field",
    sourceSha256: inspection.summary.sourceSha256,
    formFieldId: field.id,
    expected: field.snapshot,
    value: "Shanghai",
  }],
});
```

`update_form_field` accepts exactly one non-password text widget, one
non-multiselect combo whose display and export options are identical, or one
checkbox. The complete snapshot protects name/type/current value/read-only
state/options/visible widget geometry. It verifies the field state before save,
but it is not a durable field identity: re-inspect the output before any second
mutation. It may use unsigned `incremental` save and proves the exact source
prefix; it still does not authorize signed changes.

Radio buttons, shared-widget fields, list or multi-select choices, password
fields, mismatched export values, stale snapshots, and unsupported options fail
closed in this path. Route them to the explicit pypdf workflow below. Signed
PDF incremental edits are also rejected.

## Update one native review mark

Inspection returns a `snapshot` and `updateCapability` for native Text notes,
fixed FreeText, solid/no-fill Square/Circle marks, and Highlight, Underline,
StrikeOut, and Squiggly annotations. Keep the Text-note compatibility patch to
non-empty `contents`, `author`, or `subject`. For every profiled record, require
`updateCapability.supported`; FreeText permits those three fields, while an
area or text-markup profile also permits RGB `color` in `[0,1]`:

```js
const mark = inspection.records.find((record) => record.kind === "mupdfAnnotation"
  && record.type === "Highlight" && record.contents === "Check this value");
if (!mark?.updateCapability.supported) throw new Error("Review mark is read-only.");

const edited = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "update_annotation",
    page: mark.page,
    annotationId: mark.id,
    sourceSha256: inspection.summary.sourceSha256,
    expected: mark.snapshot,
    patch: { contents: "Value confirmed", subject: "Resolved", color: [0.2, 0.7, 0.3] },
  }],
});
```

The provider must preserve all type-specific invariants. For a Square/Circle,
that includes rectangle, appearance bounds, border width/style, no-fill state,
flags, page, and locator, followed by another visible-page appearance check.
Partial/stale snapshots, no-op, geometry/border/fill patches, unsupported
profiles, invalid colors, and incremental save fail closed. Re-inspect and
render the distinct output before delivery.

## Add one source-bound Text note

Select the target `mupdfPage` record from the same inspection. Use its `bbox`
and `rotation` as an exact coordinate precondition, then place one Text-note
pin with non-empty `contents`:

```js
const page = inspection.records.find((record) => record.kind === "mupdfPage"
  && record.page === 1);
if (!page) throw new Error("Expected an inspectable first page.");

const annotated = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "add_text_annotation",
    page: page.page,
    sourceSha256: inspection.summary.sourceSha256,
    expectedPage: { bbox: page.bbox, rotation: page.rotation },
    point: [72, 128],
    contents: "Review this assumption.",
    author: "Reviewer",
  }],
});
```

`point` is in the inspected page's explicit `mupdf-page-space`: upper-left
origin, y downward, with the current 0/90/180/270-degree rotation already
applied to `mupdfPage.bbox`. Raw `mediaBox`/`cropBox` facts remain unrotated PDF
coordinates and must not be substituted. This is not a request for a specific
note rectangle: the provider normalizes the native icon geometry, verifies
exactly one new Text annotation, and records both the actual rectangle and a
conservative `appearanceBbox`. The latter covers renderer differences caused
by native Text-note `NoZoom`/`NoRotate` flags. A `text` alias, `bbox`/`rect`,
icon selection, stale hash/page snapshot, clipped native appearance, or
incremental save fails closed. Re-inspect the rewrite and compare the fresh
appearance before a later annotation update/deletion; the returned xref is
current-source-only.

## Add one visible FreeText review box

Use FreeText when the annotation contents must appear directly on the page.
The operation binds the exact source and page snapshot just like a Text note,
but accepts one visible bounding box in `mupdf-page-space`:

```js
const page = inspection.records.find((record) => record.kind === "mupdfPage"
  && record.page === 1);

const annotated = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "add_free_text_annotation",
    page: page.page,
    sourceSha256: inspection.summary.sourceSha256,
    expectedPage: { bbox: page.bbox, rotation: page.rotation },
    bbox: [72, 128, 260, 56],
    contents: "Review this assumption before approval.",
    fontSize: 12,
    textColor: [0.1, 0.2, 0.8],
    alignment: "left",
    author: "Reviewer",
  }],
});
```

The bounded native appearance uses Helvetica at 4–72 points, an optional RGB
text color, and left/center/right alignment. OfficeKit extracts the generated
appearance text before publication and fails closed if the box clips or omits
requested content. It does not expose a background, border, arbitrary font,
rich-text payload, or callout. Re-inspect and render the rewrite; imported
FreeText is editable only when the fresh record advertises the
`fixed-helvetica-v1` update profile. Its complete snapshot permits
contents/author/subject changes while preserving style and geometry, and the
rebuilt appearance must still contain all requested text. Other FreeText
profiles fail closed; the generic source-bound delete route remains available.

## Mark a non-text region

Use `add_area_annotation` for a visible outline around an image, chart, or page
region where native text selection is unavailable. It requires the exact source
hash and inspected page bbox/rotation, a `rectangle` or `ellipse`, and one
`mupdf-page-space` bbox. The optional stroke is RGB, the solid border is bounded
to 0.5–12 points, and optional contents/author/subject remain review metadata.
No interior fill is created. The provider verifies one native Square/Circle,
its style, and its complete `appearanceBbox`; requested or painted bounds that
leave the visible page fail closed. Re-inspect and render the rewrite. A fresh
`solid-no-fill-v1` record may update only contents/author/subject/RGB color with
its complete snapshot; geometry, border width/style, fill, and arbitrary
appearance remain fixed. Complete-snapshot source-bound deletion is available.

## Mark one unique text selection

For an Agent review mark, select the requested page text from the same native
inspection. Do not infer character boxes or pass a viewer rectangle: MuPDF
must find one and only one native selection on the target page.

```js
const page = inspection.records.find((record) => record.kind === "mupdfPage"
  && record.page === 1);
if (!page) throw new Error("Expected an inspectable first page.");

const marked = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "add_text_markup",
    markup: "underline",
    page: page.page,
    sourceSha256: inspection.summary.sourceSha256,
    expectedPage: { bbox: page.bbox, rotation: page.rotation },
    text: "Revenue assumptions remain provisional",
    color: [0.2, 0.35, 0.9],
    contents: "Validate before approval.",
    author: "Reviewer",
  }],
});
```

`markup` is exactly `highlight`, `underline`, `strikeout`, or `squiggly`.
`add_text_highlight` remains a compatibility alias for `highlight` and does not
accept `markup`. Text is limited to 4,096 characters. Color is optional RGB in
`[0,1]` (yellow for highlight, red otherwise); `contents`, `author`, and
`subject` are optional non-empty review metadata. A zero/multiple hit, caller
quad/rectangle, stale source or page snapshot, off-page native
`appearanceBbox`, or incremental save fails closed. Right-angle rotation itself
is supported when it matches the inspected page snapshot. The audit and fresh
`mupdfAnnotation` record expose the verified native type, quadrilaterals, color,
and appearance. Render the delivered output as part of review; the resulting
xref is valid only for those exact output bytes.

Before a pypdf mutation, probe and bind the exact route. Change `--task` to `annotate` for notes:

```bash
python3 scripts/pdf_provider.py check --provider pypdf --require
python3 scripts/pdf_provider.py plan \
  --task fill-form --provider pypdf --strategy incremental \
  --input input.pdf --output tmp/pdfs/filled.pdf --require-provider
```

## Fill form with pypdf

This interactive fill is an incremental-only operation. Do not switch to
`rewrite` or flattening: those are different delivery contracts and invalidate
the editable-form and original-prefix guarantees below.

```bash
python3 scripts/pypdf_edit.py fill-form input.pdf tmp/pdfs/filled.pdf \
  --strategy incremental \
  --field 'sender.city=Shanghai' \
  --field 'approved=Yes'
```

After the fill, render and validate the exact output before handing it off:

```bash
mkdir -p tmp/pdfs/form-render
pdftoppm -png -r 144 tmp/pdfs/filled.pdf tmp/pdfs/form-render/page
python3 scripts/pdf_audit.py validate tmp/pdfs/audit.json \
  --source input.pdf --artifact tmp/pdfs/filled.pdf \
  --require-operation fill-form
```

The `pdftoppm` output is evidence for every final page, not a preview-only
step. For this pypdf form route, a MuPDF render alone is insufficient: keep the
Poppler render and validation commands after the typed `fill-form` invocation in
the same audit trace; a form output without both checks is not a deliverable.

The script sets `auto_regenerate=False` so the output carries explicit field state rather than asking the viewer to regenerate it. Use `--flatten` only with `rewrite`, after confirming that interactivity should be removed.

`--flatten` is a whole-document static-delivery operation, not a visual hint.
It paints both selected values and every unchanged canonical field value before
it removes every `/Widget` annotation and the root `/AcroForm` tree. The
adapter rejects orphan/unmodeled Widgets, fields without a page Widget, and
unsupported field types instead of silently dropping their contents. It then
reopens the output and requires
`formValidation.mode === "static"`, `allWidgetsRemoved === true`, and
`fieldTreeRemoved === true`. It retains non-Widget annotations. A viewer PNG or
the pypdf `flatten=True` flag alone is not evidence that an interactive form
was removed. Keep a separate interactive copy when the recipient might need to
edit the form later.

The adapter resolves each field type before mutation. Text and choice values remain strings; radio buttons and checkboxes are matched against their real `/AP /N` appearance-state names and written as PDF Names. Unknown button states, read-only fields, signature fields, push buttons, unsupported field types, missing appearances, or a post-write `/V`/`/AS` mismatch fail closed and remove the transactional output. This prevents a radio value from looking filled in field metadata while every widget still renders `/Off`. The adapter never calls `reattach_fields()` automatically: an orphan/ambiguous widget is not enough authority to manufacture or duplicate a canonical field tree.

## Add annotation with pypdf

```bash
python3 scripts/pypdf_edit.py add-note input.pdf tmp/pdfs/annotated.pdf \
  --strategy incremental \
  --page 1 --rect 72,640,96,664 \
  --text 'Review this assumption.'
```

The optional PyMuPDF specialist script also exposes `add_text_annotation` and `fill_form`, but it is selected explicitly rather than used as a fallback.

## Signed forms

An incremental update can retain signed byte ranges, but it can still violate DocMDP or a field lock. The script refuses signed inputs unless `--allow-signed` is explicit. Run pyHanko validation before and after and compare the reported modifications.

For one pre-verified DocMDP P=2 certification, there is a separate
[`pyHanko controlled form finalisation`](sign_verify.md#finalise-one-allowed-docmdp-p2-field)
route. It accepts one flat empty visible `/Tx` field, one explicit locked field
and value, one expected certification field, and one caller-provided trust root.
It finalises the target as a visible static/read-only decimal in one incremental
revision only after explicit-root validation proves that exactly the target field
changed. It is not a route for arbitrary signed form edits, hierarchical/shared
fields, radio/choice fields, changing a lock set, or preserving interactivity.

Record the canonical [`office-kit.pdf-audit.v1`](../references/AUDIT_SCHEMA.md) envelope and run `scripts/pdf_audit.py validate` against the exact source and delivered artifact before handoff.
