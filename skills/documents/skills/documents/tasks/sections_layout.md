# Task: Section breaks and mixed page layout

## Goal

Use public `DocumentModel` section blocks without breaking page geometry or
section-scoped headers and footers.

## Key concept

A section block inserts a break before the content that follows it. It defines
page size, orientation, and margins for the following section. Do not append an
unused section block at the end of a document; that can produce a blank page.

## Audit an existing package

```bash
python scripts/section_audit.py input.docx
```

The script is an audit, not an authoring engine. Use it to inventory section
count, geometry, and header/footer relationships before a semantic edit.

## Public API pattern

```js
import { DocumentFile, DocumentModel } from "office-kit";

const document = DocumentModel.create({ blocks: [] });
document.addParagraph("Portrait-section evidence.");

document.addSection({
  name: "landscape-evidence",
  breakType: "nextPage",
  orientation: "landscape",
  pageSize: { widthTwips: 15840, heightTwips: 12240 },
  margins: { top: 720, right: 900, bottom: 720, left: 900 },
  lineNumbering: { countBy: 5, start: 0, distance: 360, restart: "newPage" },
  pageNumbering: { start: 1, format: "lowerRoman" },
  columns: { count: 2, spacing: 720, separator: true },
});
document.addParagraph("Landscape-section evidence.");

document.addHeader("Landscape appendix", {
  referenceType: "default",
  sectionIndex: 1,
});
document.addFooter("1", {
  referenceType: "default",
  sectionIndex: 1,
  fieldInstruction: "PAGE",
});

await (await DocumentFile.exportDocx(document)).save("out.docx");
```

Imported section relationship/linkage graphs beyond the modeled section
boundary are source-bound. If changing one would invalidate source evidence,
OfficeKit fails closed; do not flatten the document to force the edit.

`columns` has two mutually exclusive bounded profiles. Equal-width layout uses
`{ count, spacing, separator }`. Asymmetric layout uses ordered native column
definitions:

```js
columns: {
  definitions: [
    { width: 3000, spacing: 720 },
    { width: 5640, spacing: 0 },
  ],
  separator: true,
}
```

Both profiles allow 1–45 columns and use twentieths of a point. In the custom
profile, `spacing` means space after that definition; never combine
`definitions` with equal-width `count` or root `spacing`. Ordinary margins,
binding gutter, widths, and inter-column gaps
must fit the page content width. Duplicate containers, ignored equal-width
root attributes, unknown children, extension-bearing definitions, and other
ambiguous graphs remain source-owned; inspect reports that section as
`editable: false`.

`lineNumbering` owns one canonical native `w:lnNumType` leaf and places line
numbers before each text column. An empty object enables every-line numbering
with `countBy: 1`. The bounded fields are:

```js
lineNumbering: {
  countBy: 5,         // 1..32767
  start: 0,           // optional native zero-based value; first display is 1
  distance: 360,      // optional twentieths of a point
  restart: "newPage", // optional: newPage | newSection | continuous
}
```

Set `section.lineNumbering = undefined` to remove the canonical leaf. This
paragraph can opt out of the displayed sequence and its calculation, or
explicitly override suppression inherited from a named style:

```js
document.addParagraph("Unnumbered heading", {
  styleId: "Heading1",
  paragraphFormat: { suppressLineNumbers: true },
});
document.addParagraph("Numbered evidence.");

// In a style definition, the same property suppresses every inheriting paragraph.
// An explicit false on a paragraph overrides that inherited suppression.
```

`true` emits direct/style `w:suppressLineNumbers`; `false` retains an explicit
direct override; omission inherits the style/default behavior. Recognized
canonical direct and style leaves are editable. Duplicate leaves, children,
extension attributes, and invalid lexical values remain source-owned and fail
closed on semantic replacement. Duplicate or irregular `w:lnNumType` leaves,
invalid numeric values, and unknown restart values likewise make the section
read-only. Use native Word/LibreOffice pagination to check the displayed
numbers and column placement; model preview alone is not authoritative.

## Keep headings and paragraphs together

Paragraph pagination directives are separate from page geometry. Use only the
one that expresses the request; OfficeKit does not simulate pagination:

```js
document.addParagraph("Decision", {
  styleId: "Heading1",
  paragraphFormat: { keepNext: true },
});
document.addParagraph("A short conclusion that must not split across pages.", {
  paragraphFormat: { keepLinesTogether: true },
});
document.addParagraph("Avoid a single first or last line at a page boundary.", {
  paragraphFormat: { widowControl: true },
});
document.addParagraph("A compact continuation in the same body style.", {
  styleId: "BodyText",
  paragraphFormat: { contextualSpacing: true },
});
document.addParagraph("A short, deliberately highlighted review note.", {
  paragraphFormat: { shadingFill: "#FEF3C7" },
});
document.addParagraph("A standalone heading for the generated outline.", {
  paragraphFormat: { outlineLevel: 1 }, // native outline level 1 (second level)
});
document.addParagraph("Override an inherited heading style to ordinary body text.", {
  styleId: "Heading1",
  paragraphFormat: { outlineLevel: 9 }, // explicit no-outline override
});
document.addParagraph("Appendix", {
  styleId: "Heading1",
  paragraphFormat: { pageBreakBefore: true },
});
```

`keepNext` asks Word to keep this paragraph with its following paragraph.
`keepLinesTogether` writes `w:keepLines`, so one paragraph may move to the
next page rather than split across two pages. `widowControl` writes
`w:widowControl`, asking the host to avoid a single first or last line at a
page boundary. `pageBreakBefore` writes `w:pageBreakBefore`. All four accept
boolean `true`/`false`; omission inherits the named-style/default behavior.
`outlineLevel` is separate document-structure metadata: `0` through `8` are
the native outline levels, and `9` explicitly clears an inherited level;
omission inherits. It writes `w:outlineLvl` and does not restyle text or
calculate a TOC.
`contextualSpacing` is a separate presence-aware spacing setting. `true` writes
`w:contextualSpacing` and suppresses `spaceBefore`/`spaceAfter` only between
adjacent paragraphs with the same style; explicit `false` overrides an
inherited style value, while omission inherits. It does not calculate layout
or collapse spacing across different styles.
`shadingFill` is a separate `#RRGGBB` callout/background primitive. It writes
only canonical `w:shd` (`w:val="clear"`, `w:color="auto"`, and a six-digit
fill), not a generic pattern or theme-color API. A recognized ordinary direct
paragraph can add, change, or clear that fill inside its modeled
direct-formatting profile; theme/pattern markup and imported style-catalog
changes remain source-bound.
Canonical direct leaves are editable after import; source-free named styles may
use the same fields, while imported style catalogs remain source-bound. Duplicate,
child-bearing, extension-bearing, or invalid lexical `w:keepNext`,
`w:keepLines`, `w:widowControl`, `w:pageBreakBefore`, or
`w:contextualSpacing`/`w:shd`/`w:outlineLvl`/`w:suppressLineNumbers` markup stays source-owned and semantic replacement
fails closed. Native
Word/LibreOffice rendering remains the final authority on actual page breaks.

`pageNumbering` owns one canonical native `w:pgNumType` leaf. Use `start` to
restart a section at an integer from 0 through 2147483647; omit it to continue
the previous section's sequence. The optional `format` is one of `decimal`,
`upperRoman`, `lowerRoman`, `upperLetter`, or `lowerLetter`. At least one of
`start` or `format` is required:

```js
pageNumbering: { start: 1, format: "lowerRoman" }
```

This setting controls page numbers displayed by PAGE fields; it does not add a
footer, insert a field, paginate the document, or refresh cached field text.
Add a PAGE footer/header separately and use native Word/LibreOffice rendering
for final QA. Chapter style/separator attributes, unsupported number formats,
duplicate leaves, extensions, and empty `w:pgNumType` markup remain
source-owned and make the section read-only.

### Change one imported section's page numbering safely

When the user asks for a tightly bounded page-numbering correction in an
existing DOCX, inspect the document first and record the section **block
index** plus the complete current value. Do not select by a guessed page number
or by text near the section break.

For the canonical profile, use the shipped transaction rather than XML patching:

```bash
officekit run examples/officekit-section-page-numbering-edit-workflow.mjs \
  input.docx output.docx audit.json \
  1 '{"start":1,"format":"lowerRoman"}' \
  '{"start":1,"format":"decimal"}'
```

The workflow requires all of the following before it writes anything:

- the selected imported block is a section, `editable`, and still resolves by
  its inspected identity;
- its complete `pageNumbering` object exactly matches the supplied source
  value;
- the corresponding raw native `w:pgNumType` is one canonical leaf; and
- `input.docx`, output, and audit paths are distinct and the output paths do
  not already exist.

It exports through `DocumentFile`, permits only `word/document.xml` to change,
normalizes namespace-declaration placement and attribute order solely for the
residual comparison, then proves that every other XML name, attribute value,
text node, element order, section, relationship, and package part stayed
unchanged. It reimports, verifies, creates a model SVG, and writes a byte-bound
audit. Use LibreOffice/Word plus Poppler for the final PAGE-field visual review.
The workflow changes metadata only: it does not insert a PAGE field or claim to
refresh a cached field display.

### Change one imported section's page margins safely

For a constrained page-margin correction, inspect the existing DOCX first and
record the section **block index** and all five current twip values. Do not
select a section from nearby body text or a guessed page number.

```bash
officekit run examples/officekit-section-margin-edit-workflow.mjs \
  input.docx output.docx audit.json \
  1 '{"top":1440,"right":1440,"bottom":1440,"left":1440,"gutter":0}' \
  '{"top":1440,"right":1440,"bottom":1440,"left":1728,"gutter":0}'
```

The workflow requires an editable, resolvable imported section whose complete
semantic `margins` object and raw canonical `w:pgMar` leaf match the supplied
source value. It changes only the five body-margin attributes in that leaf;
the native `w:header` and `w:footer` distances, sibling/terminal sections,
every relationship, and every non-document package part remain bound. It
exports through `DocumentFile`, permits only `word/document.xml` to change,
compares a namespace-tolerant raw OPC residual, reimports the full section
projection, verifies, model-renders, and publishes DOCX/audit files without
overwrite. Use LibreOffice or Word plus Poppler to review the final page
geometry.

### Change one imported section's page geometry safely

For a portrait/landscape or paper-size correction, keep the orientation and
both page dimensions together. They are one native `w:pgSz` value; do not make
two independent partial changes or ask this workflow to rescale existing text,
tables, drawings, headers, or footers.

```bash
officekit run examples/officekit-section-page-geometry-edit-workflow.mjs \
  input.docx output.docx audit.json \
  1 '{"orientation":"portrait","pageSize":{"widthTwips":12240,"heightTwips":15840}}' \
  '{"orientation":"landscape","pageSize":{"widthTwips":15840,"heightTwips":12240}}'
```

The transaction requires one editable, resolvable imported section whose full
semantic geometry and raw canonical `w:pgSz` leaf match the supplied source
value. It accepts exactly `w:w`, `w:h`, and `w:orient`, all canonical; paper
codes, extensions, duplicate/non-Word attributes, stale geometry, and missing
or partial values fail closed. It changes only that one leaf in
`word/document.xml`, checks that all remaining XML/package content is stable,
reimports the full section projection, verifies, model-renders, and publishes
the DOCX/audit without overwrite. It never rescales the surrounding document.
Review all affected pages with LibreOffice or Word plus Poppler before delivery.

### Change one imported section's line numbering safely

For one exact display-cadence correction, first inspect the imported section
and record its **block index** plus the complete normalized line-numbering
profile. `countBy` is always explicit in this workflow, even when the source
uses Word's default omitted attribute; the optional fields are `start`,
`distance`, and `restart`.

```bash
officekit run examples/officekit-section-line-numbering-edit-workflow.mjs \
  input.docx output.docx audit.json 1 \
  '{"countBy":5,"start":0,"distance":360,"restart":"newPage"}' \
  '{"countBy":10,"start":4,"distance":480,"restart":"continuous"}'
```

The transaction requires one editable, resolvable section and exactly one raw
canonical `w:lnNumType` leaf. It accepts only `w:countBy`, `w:start`,
`w:distance`, and `w:restart`; an omitted `countBy` means the native default
of `1`. Unknown/duplicate/non-Word attributes, noncanonical numeric spelling,
child content, stale values, an absent leaf, or a noncanonical section fail
closed. Only `word/document.xml` may change; the workflow masks only that
leaf for a namespace-tolerant residual comparison, reimports the full section
projection, verifies and model-renders, and writes a no-overwrite audit. It
does not add a line-number leaf, remove one, change `suppressLineNumbers`, or
calculate displayed line numbers. Review the affected pages in Word or
LibreOffice plus Poppler before delivery.

### Change one imported section's columns safely

For one exact column-flow correction, first inspect the imported section and
record its **block index** plus one complete normalized profile. Equal-width
profiles use `{ count, spacing, separator }`; explicit-width profiles use
`{ definitions: [{ width, spacing }, ...], separator }`. The transaction keeps
the inspected profile shape: it does not convert equal-width columns into an
explicit-width graph, or the reverse.

```bash
officekit run examples/officekit-section-columns-edit-workflow.mjs \
  input.docx output.docx audit.json 1 \
  '{"count":2,"spacing":720,"separator":true}' \
  '{"count":3,"spacing":360,"separator":false}'
```

The workflow requires one editable, resolvable section and one canonical
`w:cols` element. Equal-width input accepts Word's omitted `w:equalWidth` and
`w:num` defaults but requires native `w:space`; explicit-width input requires
`w:equalWidth="false"` and 1 through 45 direct `w:col` leaves. Unknown,
duplicate, non-Word, noncanonical-numeric, child/extension-bearing, stale, or
noncanonical graphs fail closed. Only `word/document.xml` may change; the
workflow masks only that one element for a namespace-tolerant residual
comparison, reimports the full section projection, verifies and model-renders,
and writes a no-overwrite audit. It does not add/remove a columns element or
calculate the visible column flow. Review affected pages in Word or LibreOffice
plus Poppler before delivery.

### Change one imported section's break type safely

For one exact section-boundary behavior correction, first inspect the imported
section and record its **block index** plus the complete current `breakType`.
The supported values are `nextPage`, `continuous`, `evenPage`, and `oddPage`.

```bash
officekit run examples/officekit-section-break-edit-workflow.mjs \
  input.docx output.docx audit.json 1 nextPage continuous
```

The workflow requires one editable, resolvable section and exactly one raw
canonical `w:type w:val="..."/>` leaf. Unknown/duplicate/non-Word attributes,
missing or stale leaves, unsupported values, and noncanonical section graphs
fail closed. Only `word/document.xml` may change; the workflow masks only that
leaf for a namespace-tolerant residual comparison, reimports the full section
projection, verifies and model-renders, and writes a no-overwrite audit. It
does not add/remove/move a section boundary or calculate the resulting page
breaks. Review every affected page in Word or LibreOffice plus Poppler before
delivery, especially for even/odd-page behavior.

## Render review

- Only the intended pages change orientation.
- Page size and margins match the explicit twip values.
- Text flows through the requested equal-width or asymmetric columns and separator rules.
- Line numbers use the requested increment, offset, restart behavior, and distance in every text column.
- PAGE fields restart/continue and display in the requested section format.
- Headers and footers appear in the intended section.
- First/even variants behave as requested.
- No empty trailing page was introduced.

When a document mixes page sizes or orientations, pass an explicit `--dpi` to
`render_docx.py` if exact output pixel dimensions matter.
