# Accessibility (A11y) Audit + Quick Fixes

## Goal
Given a `.docx`, produce an **accessibility audit report** and (optionally) apply **safe, mechanical fixes** that reduce common A11y failures.

This is **not** a full WCAG compliance engine. It targets the highest-ROI checks you can do reliably in OOXML:
- Heading hierarchy (no skipping levels)
- Images missing alt text (`descr`)
- Tables missing a header row flag
- Hyperlink text that is non-descriptive ("click here", raw URLs)

## Audit

Use the OfficeKit audit first. It reports stable block locators, separates
machine issues from manual author-intent checks, binds an existing source by
SHA-256, and never edits the DOCX:

```bash
officekit run examples/officekit-accessibility-audit-workflow.mjs \
  input.docx a11y_report.json
```

The machine check covers explicit heading-level skips, empty image alternative
text, tables without a declared repeating-header prefix, and empty hyperlink
text. Missing table purpose/description and generic or raw-URL link labels are
manual checks. Opaque package content and complete heading/table/link intent
also require source or native-host review. Read
`../artifact_tool/ACCESSIBILITY_AUDIT.md` for the result contract.

The Python helper below remains an explicit package-level audit/fix route for
the reference-compatible quick-fix policy; it is not a fallback from the
OfficeKit model:

```bash
python scripts/a11y_audit.py input.docx
```

This prints a JSON-ish report to stdout and exits non-zero if **high severity** issues exist.

To write the report to a file instead:
```bash
python scripts/a11y_audit.py input.docx --out_json a11y_report.json
```

## Apply quick fixes (optional)
### 1) Update one identified image with reviewed alternative text
When an agent or reviewer can identify one canonical imported body image and
write its actual description, prefer the source-bound OfficeKit transaction:

```bash
officekit run examples/officekit-image-alt-text-edit-workflow.mjs \
  input.docx reviewed.docx image-alt.audit.json 1 \
  "Existing architecture overview" \
  "Architecture overview showing the API gateway and three service lanes"
```

It binds the inspected block index and exact prior text, changes only paired
native description leaves, preserves the media and layout, and emits a
no-overwrite audit. It cannot decide that a picture is decorative or invent a
missing description. See `tasks/images_figures.md` for the complete boundary.

### 2) Fill missing image alt text using filenames
This is a pragmatic baseline that is better than empty alt text.
```bash
python scripts/a11y_audit.py input.docx --fix_image_alt from_filename --out a11y_fixed.docx
```

### 3) Mark repeating table headers
Only do this when leading rows really are headers. For a new document, declare
the native repeat-header semantics separately from the visual first-row fill:

```js
document.addTable({
  values: [["Metric", "Value"], ["Readiness", "92%"]],
  headerFill: "DCEAF3",
  headerRowCount: 1,
});
```

For one imported flat rectangular table, bind its inspected block index and the
complete current/replacement count:

```bash
officekit run examples/officekit-table-header-rows-edit-workflow.mjs \
  input.docx a11y_fixed.docx table-header.audit.json 1 1 2
```

It changes only native `w:tblHeader` markers and fails closed for non-prefix,
duplicate, explicit-value, extension-bearing, merged, nested, or irregular
inputs. It does not infer header semantics from bold text or fill. The Python
audit helper remains explicit when you deliberately want its report/fix policy.

### 4) Repair one imported hyperlink label
When an audit finds an empty label, or a reviewer replaces generic/raw-URL text,
bind the imported block index, complete current text, and exact destination:

```bash
officekit run examples/officekit-hyperlink-text-edit-workflow.mjs \
  input.docx reviewed.docx hyperlink-text.audit.json 4 \
  "https://example.com/accessibility" "" \
  "Read the accessibility guide"
```

The workflow changes one canonical whole-paragraph hyperlink's native `w:t`
leaf only, preserves its http(s) URL or bookmark destination, relationship ID,
tooltip, history flag, paragraph/run formatting, and every non-document package
part, then reimports and reruns the accessibility audit. Empty replacement
text, stale source facts, rich/multi-run/nested/table/textbox hyperlinks,
destination edits, output collisions, and package drift fail closed. A person
must still judge whether the new label communicates the destination's purpose.

### 5) Review non-visible table alternative text
Table alternative text is not a visible caption. When a reviewer supplies the
actual title and description for one imported canonical table, bind both the
inspected block index and the complete current metadata, then use the
source-bound OfficeKit transaction:

```bash
officekit run examples/officekit-table-accessibility-edit-workflow.mjs \
  input.docx reviewed.docx table-a11y.audit.json 1 \
  '{"title":"Quarterly delivery readiness","description":"A two-column release-readiness matrix."}' \
  '{"title":"Release-readiness decision matrix"}'
```

The replacement object is deliberate: omitting `description` clears that
non-visible native `w:tblDescription` value, while the table remains visually
unchanged. The workflow changes only canonical `w:tblCaption/@w:val` and
`w:tblDescription/@w:val` leaves in `word/document.xml`, preserves the input,
reimports the complete table projection, and writes a no-overwrite audit. It
does not create a caption paragraph, infer author intent, edit cell text or
formatting, or accept duplicate, empty, child-bearing, extension-bearing, or
otherwise irregular alternative-text leaves.

## Verification loop
1) Apply fixes (if any)
2) **Render → inspect PNGs** to confirm nothing drifted visually:
```bash
python render_docx.py a11y_fixed.docx --output_dir out_a11y
```

## Pitfalls
- "Fixing" headings is rarely mechanical; it usually requires editorial judgement. This tool **reports** heading issues but does not rewrite styles.
- Setting table header flags can change repeated header rendering across page breaks. Always re-render and review.
- Alt text generated from filenames is a baseline; replace it with meaningful descriptions for real accessibility.
- A table title and description help assistive technology but do not make an arbitrary table accessible by themselves; header semantics, reading order, and real author intent still need review.
