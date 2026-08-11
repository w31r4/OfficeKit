# Document accessibility audit

`document.auditAccessibility()` is the host-neutral, read-only accessibility
preflight for modeled Word content. It returns stable block locators and keeps
machine-checkable defects separate from author-intent review.

```js
const report = document.auditAccessibility({ maxChars: 200_000 });
if (!report.machineCheckPassed) {
  for (const issue of report.issues) {
    console.error(issue.type, issue.blockIndex, issue.id, issue.message);
  }
}
for (const check of report.manualChecks) {
  console.error("manual", check.type, check.blockIndex, check.message);
}
```

## What it checks

Machine issues are limited to modeled facts OfficeKit can test without
guessing author intent:

- `headingLevelSkipped`: an explicit outline/Heading style jumps over a level;
- `imageAltTextMissing`: an image has empty alternative text;
- `tableHeaderRowMissing`: a table has no declared repeating header-row prefix;
- `hyperlinkTextMissing`: a hyperlink has no visible text.

Manual checks remain separate:

- `tablePurposeAndDescription`: decide whether a table needs non-visible title
  or description and whether its declared leading rows are real headers;
- `hyperlinkPurpose`: review generic phrases and raw-URL labels in context.

The result always sets `conformanceClaimed: false`. A green machine result is
not Microsoft Word Accessibility Checker, WCAG conformance, proof of correct
heading intent, or evidence about opaque/unmodeled package content.

## Existing DOCX workflow

Use the packaged workflow when the input is an existing file:

```bash
officekit run examples/officekit-accessibility-audit-workflow.mjs \
  input.docx accessibility-report.json
```

It imports through OfficeKit Codec, records the package version and source
SHA-256, performs `document.verify()`, re-hashes the source, writes a private
temporary report, and promotes it without overwrite. Its save policy is
`none`: it produces no DOCX and never mutates the input.

After review, use the narrow source-bound heading-level, image-alt,
table-header, hyperlink-text, or table-accessibility workflow documented in
`tasks/accessibility_a11y.md`. The four machine issue types each have one
bounded repair route, but none supplies the missing author intent. Do not infer
a hierarchy, invent alternative text, infer a header from bold/fill alone, or
silently use the Python package helper as another Office authoring engine.
Reopen and render any corrected DOCX before delivery.
