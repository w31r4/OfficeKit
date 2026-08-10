# Presentation accessibility audit

OfficeKit separates facts it can prove from review that still depends on
PowerPoint and author intent. Object metadata and classification are modeled;
assistive reading order and whole-deck conformance are not inferred.

## Author object metadata

Meaningful ordinary shapes, connectors, groups, images, tables, and charts use
one non-visible `accessibility` object:

```js
const chart = slide.charts.add("bar", {
  name: "revenue-chart",
  categories: ["North", "South"],
  series: [{ name: "Revenue", values: [18, 14] }],
  accessibility: {
    title: "Regional revenue comparison",
    description: "North revenue is 18 and South revenue is 14.",
    decorative: false,
  },
});

slide.shapes.add({
  name: "decorative-rule",
  geometry: "line",
  position: { left: 80, top: 180, width: 400, height: 0 },
  accessibility: { decorative: true },
});
```

`decorative: false` is an explicit meaningful classification. It is distinct
from omission and must carry a title or description to pass the machine audit.
`decorative: true` cannot coexist with either text field. Imported objects must
advertise `accessibilityCapability.editable` before metadata is changed.

## Run the bounded audit

```js
const audit = presentation.auditAccessibility({ maxChars: 200_000 });

if (!audit.machineCheckPassed) {
  console.error(audit.ndjson);
}

for (const check of audit.manualChecks) {
  console.log(check.type, check.message);
}
```

The report has stable top-level fields:

- `machineCheckPassed`: false when a modeled object is unclassified, or when an
  explicitly meaningful object has neither title nor description.
- `summary`: slide/object counts split across meaningful, decorative,
  unclassified, missing-text, and opaque-native states.
- `issues`: machine-checkable failures with slide, object kind, stable ID, name,
  and optional parent-group locator.
- `manualChecks`: separate `readingOrder` and `opaqueObjectAccessibility`
  records.
- `conformanceClaimed: false`: always explicit. Passing the machine check is not
  a claim that PowerPoint Accessibility Checker, WCAG, or another complete
  conformance review has passed.

PowerPoint does not expose an independent reading-order leaf through the
bounded OfficeKit model. Reordering the native shape tree would also change
visual z-order, so the audit never performs that mutation as a hidden fix.

## Audit an existing PPTX without modifying it

The packaged read-only workflow protects the source, imports it through the
OfficeKit Codec, writes one no-overwrite JSON report, and records package
version, source SHA-256, provider identity, save policy, and explicit limits:

```bash
node examples/officekit-accessibility-audit-workflow.mjs \
  input.pptx \
  accessibility-report.json
```

It produces no PPTX artifact and uses `savePolicy.strategy: "none"`. A report
path collision, symlink input, source mutation, invalid input package, or
unsupported import fails closed. Review every returned manual check in native
PowerPoint or another explicitly selected host before making a whole-deck
accessibility statement.
