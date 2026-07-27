# Accessibility (A11y) Audit + Quick Fixes

## Goal
Given a `.docx`, produce an **accessibility audit report** and (optionally) apply **safe, mechanical fixes** that reduce common A11y failures.

This is **not** a full WCAG compliance engine. It targets the highest-ROI checks you can do reliably in OOXML:
- Heading hierarchy (no skipping levels)
- Images missing alt text (`descr`)
- Tables missing a header row flag
- Hyperlink text that is non-descriptive ("click here", raw URLs)

## Audit
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
