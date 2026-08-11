# Workbook accessibility audit

`workbook.auditAccessibility()` is a host-neutral, read-only preflight for
modeled worksheet images and charts. It requires each drawing to be explicitly
meaningful or decorative and keeps machine findings separate from author and
native-host review.

```js
const report = workbook.auditAccessibility({ maxChars: 200_000 });
for (const issue of report.issues) {
  console.error(issue.type, issue.sheet, issue.id, issue.message);
}
```

Machine issues are `unclassifiedDrawing` and
`meaningfulDrawingTextMissing`. A visible chart title is not silently reused as
non-visible alternative text, and an image name or filename never becomes
`image.alt`. Set `{ title?, description?, decorative? }` through
`setAccessibilityMetadata(...)`; `decorative: true` cannot carry text.

The report always sets `conformanceClaimed: false`. Drawing keyboard/assistive
order, worksheet names and purpose, table header intent, merged-cell
navigation, color-only meaning, opaque native objects, and Excel Accessibility
Checker results remain manual review.

For an existing file, run:

```bash
officekit run examples/officekit-accessibility-audit-workflow.mjs \
  input.xlsx accessibility-report.json
```

The workflow binds and rechecks the source SHA-256, imports only through the
OfficeKit Codec, runs `workbook.verify()`, and publishes one private JSON report
without overwrite. Its save policy is `none`; it does not emit or mutate an
XLSX and has no alternate authoring fallback.
