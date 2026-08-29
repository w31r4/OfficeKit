# API catalog

Generated from `HELP_CATALOG` in `src/help/index.mjs`.

## document

| Name | Kind | Summary |
| --- | --- | --- |
| `document.addBibliography` | api | Append one canonical switch-free BIBLIOGRAPHY output placeholder, requiring at least one modeled bibliography source and enabling updateFields-on-open by default. A compatible Word host formats entries; imported fields permit cached-display edits only. |
| `document.addBibliographySource` | api | Add a canonical Word bibliography source for inspect, resolve, and native b:Sources authoring. Recognized imports allow bounded source content edits while source order, IDs, and tags remain source-bound. |
| `document.addBlockTextContentControl` | api | Append one canonical block-level Word plain-text content control around exactly one modeled paragraph and one ordinary run. The handle reports placement=block; OfficeKit preserves the w:sdt wrapper and binds native identity/topology after import. |
| `document.addBookmark` | api | Wrap exactly one paragraph-like block in a native Word bookmark for inspect, resolve, and internal hyperlinks. Recognized imported whole-block bookmarks are exposed with source identity but remain fixed-topology/read-only; cross-block, nested, crossing, table-cell, and otherwise complex ranges stay opaque-preserved and fail closed on mutation. |
| `document.addChange` | api | Append one bounded whole-paragraph tracked insertion or deletion. OfficeKit authors native w:ins/w:del markup and permits fixed-topology imported text/author/date edits; mixed or nested revision graphs remain source-bound. |
| `document.addCitation` | api | Add a whole-paragraph bibliography-backed citation exported as a native w:fldSimple CITATION field plus a bounded bookmark. Recognized imports allow display-text edits while source tags and topology remain fixed. |
| `document.addComment` | api | Attach a whole-paragraph Word comment. Classic roots remain minimal; bounded modern roots may carry resolved, durable/UTC, and provider-person metadata through OfficeKit. |
| `document.addDeletion` | api | Append one bounded whole-paragraph tracked deletion using native w:del/w:delText markup. For one exact in-paragraph replacement in existing source bytes, use DocumentFile.addTrackedReplacement; mixed, moved, nested, and property-level revisions remain outside the bounded profile. |
| `document.addEndnote` | api | Append one native plain-text endnote with 1 through 16 canonical physical paragraphs at the end of one paragraph or list item. Recognized imported canonical endnotes permit fixed-count body-text edits only; anchor, kind, native ID, and note topology remain source-bound. |
| `document.addField` | api | Append a bounded w:fldSimple block for PAGE, NUMPAGES, SECTION, date/time, and selected document-property commands. External-content and arbitrary reference commands fail closed. |
| `document.addFooter` | api | Add a default, first-page, or even-page DOCX footer, optionally section-scoped. Source-free input may be plain text, one legacy simple field, or a 2-through-32 ordered literal/simple-field segment sequence in one native paragraph. Imported multi-segment page furniture is inspectable and exactly preserved on no-op export but remains source-bound/read-only. |
| `document.addFootnote` | api | Append one native plain-text footnote with 1 through 16 canonical physical paragraphs at the end of one paragraph or list item. Recognized imported canonical footnotes permit fixed-count body-text edits only; anchor, kind, native ID, and note topology remain source-bound. |
| `document.addHeader` | api | Add a default, first-page, or even-page DOCX header, optionally section-scoped. Source-free input may be plain text, one legacy simple field, or a 2-through-32 ordered literal/simple-field segment sequence in one native paragraph. Imported multi-segment page furniture is inspectable and exactly preserved on no-op export but remains source-bound/read-only. |
| `document.addHyperlink` | api | Append a native w:hyperlink backed by an external relationship or internal bookmark anchor; native import restores URL/anchor, relationship identity, tooltip, and history state. |
| `document.addImage` | api | Append an inspectable embedded PNG/JPEG image. Images are inline by default; an explicit bounded placement authors a native foreground wp:anchor with square or top-and-bottom wrapping. |
| `document.addInsertion` | api | Append one bounded whole-paragraph tracked insertion using native w:ins markup. For one exact in-paragraph replacement in existing source bytes, use DocumentFile.addTrackedReplacement; mixed, moved, nested, and property-level revisions remain outside the bounded profile. |
| `document.addListItem` | api | Append a numbered, character-bulleted, or bounded picture-bulleted list item using native DOCX numbering definitions. Picture markers are shared numbering-level resources: every item using the same numberingId and level must agree, and recognized imported edits must update the complete group without changing embedded-versus-external source kind. |
| `document.addParagraph` | api | Append a styled paragraph with optional run spans and bounded direct paragraph formatting, including canonical solid shading and solid paragraph borders, presence-aware contextual spacing, and line-number suppression. |
| `document.addSection` | api | Append a DOCX section break with page size, orientation, margins, binding gutter, canonical equal-width or explicit-width columns, bounded page-number start/format, and break-type metadata backed by w:sectPr. Imported geometry and page numbering are writable only when their native markup is canonical. |
| `document.addTable` | api | Append a Word-style table with physical cell values, optional logical merge geometry, fixed-layout width/margin/border styling, optional left/center/right table placement, optional uniform top/center/bottom physical-cell alignment, non-clipping per-row minimum heights, an optional native repeating-header prefix, individual rows kept together across pages, and optional non-visible table alternative text. |
| `document.addTableOfContents` | api | Append one canonical one-paragraph complex TOC field with bounded heading levels/switches and enable the native updateFields-on-open hint by default. Refreshed cross-paragraph result graphs remain opaque/source-bound and read-only. |
| `document.addWatermark` | api | Add one canonical VML text watermark to a section/header-reference scope. Recognized imported watermarks permit text-only edits or whole-object removal; adding to an imported package, changing scope, shared headers, multiple objects, DrawingML, images, and irregular VML fail closed. |
| `document.applyDesignPreset` | api | Apply a clean-room report or memo design preset that updates named styles for consistent DOCX export and SVG/layout previews. |
| `document.auditAccessibility` | api | Audit modeled Word headings, image alternative text, table header semantics, and hyperlink text with stable block locators. Machine-checkable defects remain separate from table-purpose and link-purpose manual review, and the result never claims Word Accessibility Checker or WCAG conformance. |
| `document.contentControls` | api | List typed mutable handles for recognized inline or table-cell plain-text, checkbox, drop-down, combo-box, and date controls plus block plain-text controls, with explicit placement and model/native identity. |
| `document.fillContentControls` | api | Transactionally fill every recognized block, inline, or table-cell plain-text control matching an object or Map of tag-to-string entries. Checkbox, drop-down, combo-box, and date tags do not silently accept text. |
| `document.fontFamilies` | api | Return a fresh sorted, case-insensitively deduplicated list of document theme and explicit run/style font families. |
| `document.inspect` | api | Emit bounded NDJSON for document blocks including typed block/inline plain-text and inline checkbox/list/date content controls with explicit placement, fields, tracked changes, bookmark ranges, footnotes/endnotes, bibliography sources, comments, styles, headers/footers with sourceBound/editable evidence, canonical text watermarks, and layout; narrow with search/target anchors and fields with include/exclude. |
| `document.layoutJson` | api | Return page-aware layout JSON with block bounding boxes, section/page ordinals, effective inherited header/footer selections, styles, and target/search slicing. |
| `document.materializeFields` | api | Transactionally compute canonical inline SEQ counters and REF cached results from native bookmark targets, with dry-run evidence and strict missing-target failure. PAGEREF remains skipped because trustworthy page numbers require a real pagination host. |
| `document.render` | api | Render an SVG preview by default, return layout JSON with { format: 'layout' }, or use { source: 'docx', renderer } to feed native DOCX into LibreOffice/native Office render adapters for PDF/PNG outputs. |
| `document.replyToComment` | api | Add one source-free direct reply to a root comment. OfficeKit authors the bounded commentsExtended graph; nested replies and imported topology changes fail closed. |
| `document.resolve` | api | Resolve stable document, block, table-cell, content-control, bookmark, footnote/endnote, bibliography source ID/tag, header/footer, watermark, comment, style, and advertised text-range IDs. |
| `document.setCheckboxContentControls` | api | Transactionally set every recognized canonical checkbox control matching an object or Map of tag-to-boolean entries. Other control types do not silently coerce. |
| `document.setComboBoxContentControls` | api | Transactionally set every recognized canonical combo-box control from a tag-to-value string mapping. Values may select a declared choice or provide bounded custom text; unknown tags and invalid values fail before mutation. |
| `document.setDateContentControls` | api | Transactionally set every recognized canonical date control from a tag-to-YYYY-MM-DD mapping. Invalid Gregorian dates, unknown tags, and other control types fail before mutation. |
| `document.setDropdownContentControls` | api | Transactionally set every recognized canonical drop-down control from a tag-to-choice-value string mapping. Unknown tags or values outside the declared choice table fail before mutation. |
| `document.setSectionSettings` | api | Set per-section Word behavior such as different-first-page header/footer activation without changing preserved header/footer references. |
| `document.setSettings` | api | Set model settings. evenAndOddHeaders, mirrorMargins, gutterAtTop, trackRevisions, the updateFields refresh hint, and bounded passwordless documentProtection are inside the bounded OfficeKit DOCX profile. Irregular page-margin mode markup and password/cryptographic protection variants stay source-owned and fail closed on replacement. |
| `document.styles.effective` | api | Resolve a named document style through basedOn inheritance so inspect/layout/render/DOCX export share the same effective style metadata. |
| `document.textRange` | api | Inspect or resolve stable textRange anchors such as blockId/text and tableId/cell/row/column/text. Assignment is limited to fully editable text; replace() also supports explicitly advertised source-bound literal patches. |
| `document.verify` | api | Return QA issues for invalid/duplicate content-control IDs and native IDs, malformed tags/aliases, invalid block-control profiles, fake lists, invalid links/citations/bibliography sources, malformed tracked changes, duplicate/dangling/reversed bookmark ranges, invalid footnotes/endnotes, unknown styles, malformed tables, bad images/sections, invalid watermark IDs/scopes/text, dangling comments, visual overflow, and prose-like table cells. |
| `documentComment.reopen` | api | Clear the resolved state of a bounded modern comment without changing its root/reply topology or durable identity. |
| `documentComment.resolve` | api | Set resolved=true for a bounded modern comment. Imported edits re-prove source hashes and commentsExtended topology while keeping thread identity fixed. |
| `DocumentFile.addTrackedReplacement` | api | Add one exact replacement inside a direct body paragraph or bounded table-cell paragraph to hash-bound DOCX source bytes as adjacent native w:del/w:ins runs. A structured paragraph/tableCell selector, full expected text, and one unique literal contained in either one ordinary run or adjacent run fragments with identical w:rPr preserve source formatting; mixed formatting and broader topologies fail closed with exact changed-part audit. |
| `DocumentFile.exportDocx` | api | Export DocumentModel to DOCX through the single bundled OfficeKit codec. Only limits is accepted; legacy codec and lossy-fallback options fail explicitly. |
| `DocumentFile.finalizeRevisions` | api | Accept or reject bounded direct whole-paragraph one-run revisions and exact adjacent in-paragraph w:del + w:ins pairs from source bytes, including same-format fragmented deletions in direct body paragraphs or bounded table-cell paragraphs. Mandatory SHA-256 binding, decompression budgets, exact changed-part audit, and fail-closed graph checks prevent silent model reconstruction or broad package mutation. |
| `DocumentFile.importDocx` | api | Import relationship-driven core DOCX semantics through the single bundled OfficeKit codec. An imported header/footer advertises editable only for one direct unformatted text paragraph in a uniquely used source part; recognized ordered literal/simple-field page furniture is exposed as segments but remains source-bound/read-only and no-op preserved. PAGE/simple fields, rich, shared, inherited, and irregular page furniture stay read-only. Recognized inline controls, fields, revisions, notes, citations, simple tables, and exclusive canonical VML text-watermark paragraphs are fixed-topology editable; a canonical BIBLIOGRAPHY output field permits only its cached display text to change. Otherwise read-only paragraphs and complex table cells separately advertise textPatchable when at least one direct ordinary native text node can participate in a bounded literal patch. A unique literal may span adjacent same-format runs without rebuilding the surrounding graph. |
| `DocumentFile.inspectDocx` | api | Inspect bounded DOCX parts, content types, the required main-document/root officeDocument relationship, and namespace-aware source XML r:id/r:embed/r:link references after raw-input, part-count, decompression, and optional compression-ratio budgets; verifyCrc32 additionally checks ZIP entry CRCs. |
| `DocumentFile.patchDocx` | api | Apply DOCX part patches with path traversal validation for settings, classic-comment anchors, commentsExtended/commentsIds/commentsExtensible/people parts, and numbering assignments; atomically reject dangling packages and invalid comment graphs. |
| `documentHeaderFooter.setSegments` | api | Atomically replace one source-free header/footer's ordered literal/simple-field sequence. The derived visible text must remain the concatenated segment displays; imported page furniture cannot use this mutation profile. |
| `DocumentModel.create` | api | Create a document with paragraph/character styles, formatted paragraphs/runs including canonical solid paragraph shading and bounded solid paragraph borders, canonical inline and one-paragraph table-cell plain-text, checkbox, drop-down, combo-box, and ISO/Gregorian date content controls, one-paragraph block plain-text controls, canonical inline SEQ/REF/PAGEREF fields, sections, headers/footers, canonical VML text watermarks, lists, TableGrid fixed-geometry tables, links, bounded whole-block bookmarks, 1-through-16-paragraph plain-text footnotes/endnotes, canonical bibliography-backed citations plus one source-free switch-free BIBLIOGRAPHY output placeholder, simple fields, a canonical complex TOC placeholder, bounded whole-paragraph tracked insertions/deletions, classic comments, bounded modern root/direct-reply threads, and PNG/JPEG images. Nested/irregular modern threads, rich comment bodies, multi-paragraph/rich/inline-within-cell/nested/data-bound/locked/placeholder table-cell SDTs, other nested/data-bound/locked/placeholder SDTs, irregular lists, localized dates, custom checkbox symbols, image/DrawingML/irregular VML watermarks, other complex field graphs, arbitrary table-style graphs, complex bookmark/note/revision graphs, and advanced settings remain unsupported or source-bound. |
| `documentTable.setAccessibilityMetadata` | api | Set or clear non-visible Word table alternative text through w:tblCaption and w:tblDescription. It never creates a visible caption paragraph or changes layout; duplicate, empty, child-bearing, extension-bearing, or irregular imported leaves fail closed. |
| `documentTable.setHeaderRowCount` | api | Set the number of contiguous leading rows marked with native w:tblHeader repetition semantics. This is separate from headerFill styling; imported tables accept it only when their row-property profile is canonical, otherwise the edit fails closed. |
| `documentTable.setHorizontalAlignment` | api | Set or clear native table-level w:jc placement. Center/right require zero table indent so OfficeKit never relies on host-specific resolution of competing w:jc and w:tblInd values; irregular imported table-property profiles fail closed. |
| `documentTable.setMinimumRowHeight` | api | Set or clear one physical row's non-clipping minimum height through native w:trHeight hRule=atLeast. It is not a fixed exact height or a pagination calculator; imported tables accept it only under the canonical row-property profile, otherwise the edit fails closed. |
| `documentTable.setRowKeepTogether` | api | Set whether one physical table row may split across pages through native w:cantSplit. This is a per-row pagination constraint, not a row-group or pagination calculator; imported tables accept it only under the canonical row-property profile, otherwise the edit fails closed. |
| `documentTableCell.addCheckboxContentControl` | api | Wrap one source-free rectangular table cell in a canonical Word 2010+ checkbox w:sdt. OfficeKit owns the visible glyph and symbols; recognized imports permit checked/tag/alias edits while identity, type, placement, symbols, and topology remain fixed. |
| `documentTableCell.addComboBoxContentControl` | api | Wrap one source-free rectangular table cell in a canonical standard combo-box w:sdt with ordered choices and a declared-or-custom typed value. Recognized imports permit value/tag/alias edits while the choice table and topology remain fixed. |
| `documentTableCell.addDateContentControl` | api | Wrap one source-free rectangular table cell in the canonical ISO/Gregorian date w:sdt profile. Recognized imports permit dateValue/tag/alias edits while native date metadata, placement, and topology remain fixed. |
| `documentTableCell.addDropdownContentControl` | api | Wrap one source-free rectangular table cell in a canonical standard drop-down w:sdt with ordered choices and a typed selectedValue. Recognized imports permit selectedValue/tag/alias edits while the choice table and topology remain fixed. |
| `documentTableCell.addTextContentControl` | api | Wrap one source-free rectangular table cell's existing text in a canonical cell-level plain-text w:sdt. The handle reports placement=tableCell plus row/column; recognized imported controls permit fixed-topology text/tag/alias edits, while adding or removing imported control topology fails closed. |
| `documentTableCell.replaceText` | api | Apply a literal source-bound text patch to one table cell that advertises textPatchable. The search must resolve exactly once inside one ordinary native w:t node or adjacent non-empty direct runs with byte-identical w:rPr. Whole-cell replacement, mixed formatting, empty-run gaps, paragraph boundaries, fields, controls, revisions, and ambiguous matches fail closed. |
| `documentWatermark.remove` | api | Remove one modeled or recognized source-bound canonical watermark as a complete header paragraph. The source-bound operation re-proves exact element and header residual hashes and never heuristically deletes arbitrary header graphics. |
| `exportDocxWithOfficeKit` | api | Export bounded DocumentModel paragraphs/runs, fields, tables, bookmarks, notes, citations plus one canonical bibliography-output placeholder, tracked changes, comments, images, canonical text watermarks, sections, numbering, settings, and source-free ordered header/footer literal/simple-field sequences; recognized imports permit exact-profile semantic edits plus hash-bound literal patches to one unique ordinary paragraph or table-cell span inside one direct w:r/w:t or adjacent same-format runs while preserving all surrounding native markup. |
| `importDocxWithOfficeKit` | api | Import DOCX bytes through OfficeKit with source-bound blocks, recognized exclusive canonical VML text-watermark paragraphs, source-bound header/footer editable evidence, and read-only ordered header/footer literal/simple-field sequences. A header/footer edit is limited to one direct unformatted text paragraph in one uniquely used source part; fields, rich/shared/inherited page furniture, scope changes, and multiple edits to one part fail closed. Literal body/table patch capability never implies whole-paragraph/cell editability; only adjacent non-empty direct runs with byte-identical w:rPr may form one patch span, while mixed-format, gapped, cross-paragraph, ambiguous, field/control/revision text remains fail-closed. |
| `paragraph.addCheckboxContentControl` | api | Append one canonical Word 2010+ checkbox content control with typed checked state; OfficeKit owns its visible glyph and w14 symbol declarations. |
| `paragraph.addComboBoxContentControl` | api | Append one canonical inline Word combo-box content control with ordered displayText/value choices and a typed value that may be a declared choice or bounded custom text. OfficeKit derives the visible projection. |
| `paragraph.addDateContentControl` | api | Append one canonical inline Word date picker from a real Gregorian YYYY-MM-DD value. OfficeKit owns the fixed ISO display, UTC-midnight fullDate, language, mapping, and calendar projection. |
| `paragraph.addDropdownContentControl` | api | Append one canonical inline Word drop-down content control with an ordered displayText/value choice table and typed selectedValue. OfficeKit derives visible text from the selected choice. |
| `paragraph.addField` | api | Append a logical inline SEQ, REF, or PAGEREF field run. A SEQ run may add a bookmark around only its cached result for real caption-number targets. OfficeKit authors/imports the canonical native graph; imported field position, instruction, and bookmark identity remain source-bound while cached display text is editable. |
| `paragraph.addTextContentControl` | api | Append one inline plain-text Word content-control run with agent ID, tag, alias, text, and optional run formatting. OfficeKit assigns native w:id identity and authors canonical w:sdt markup. |
| `paragraph.replaceText` | api | Replace literal paragraph text without flattening formatting boundaries. Fully editable one-run paragraphs update their existing run; imported source-bound paragraphs advertise textPatchable when OfficeKit can replace one unique ordinary w:r/w:t node or adjacent non-empty direct runs with byte-identical w:rPr while preserving all native topology and surrounding markup. Mixed formatting, empty-run gaps, paragraph boundaries, fields, controls, revisions, and duplicate matches fail closed. |

### document details

#### `document.addBibliography`

Append one canonical switch-free BIBLIOGRAPHY output placeholder, requiring at least one modeled bibliography source and enabling updateFields-on-open by default. A compatible Word host formats entries; imported fields permit cached-display edits only.

**Schema parameters:**

- `display` (string) — Cached placeholder result shown until a compatible Word host refreshes the bibliography.
- `updateFields` (boolean) — Enable the updateFields-on-open hint; defaults to true.
- `styleId` (string) — Paragraph style ID.
- `id` (string) — Optional model-local field ID.
- `name` (string) — Optional inspectable field name.

**Schema returns:**

- `field` (DocumentFieldBlock) — One canonical switch-free BIBLIOGRAPHY output placeholder. Requires at least one modeled bibliography source; imported fields permit cached-display edits only.

#### `document.addBibliographySource`

Add a canonical Word bibliography source for inspect, resolve, and native b:Sources authoring. Recognized imports allow bounded source content edits while source order, IDs, and tags remain source-bound.

**Schema parameters:**

- `tag` (string) required — Unique Word source tag used by CITATION fields: 1 through 255 ASCII letters, digits, periods, underscores, colons, or hyphens.
- `sourceType` (string) required — Word bibliography source type such as Book, Report, JournalArticle, InternetSite, or Misc.
- `title` (string) — Source title.
- `authors` (object[]|string[]) — Personal contributors with first/middle/last names.
- `corporateAuthor` (string) — Corporate author used when personal authors are absent.
- `year` (string|number) — Publication year.
- `publisher` (string) — Publisher.
- `url` (string) — Source URL.
- `fields` (object) — Additional supported Word bibliography fields such as city, journalName, volume, issue, pages, edition, and standardNumber.

**Schema returns:**

- `source` (DocumentBibliographySource) — Canonical b:Source entry. Recognized imports permit bounded field/author edits with fixed source order, ID, and tag.

#### `document.addBlockTextContentControl`

Append one canonical block-level Word plain-text content control around exactly one modeled paragraph and one ordinary run. The handle reports placement=block; OfficeKit preserves the w:sdt wrapper and binds native identity/topology after import.

**Schema parameters:**

- `text` (string) required — Initial visible paragraph text, including the empty string when the template is intentionally blank.
- `blockId` (string) — Optional agent-facing paragraph block ID; generated when omitted.
- `id` (string) — Agent-facing content-control ID; generated when omitted.
- `tag` (string) required — Block plain-text SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Human title/alias, 1 to 255 characters; defaults to tag.
- `styleId` (string) — Optional modeled paragraph style ID.
- `paragraphFormat` (object) — Optional modeled paragraph formatting for the wrapped paragraph, including canonical #RRGGBB shadingFill, bounded solid borders, keepNext, boolean keepLinesTogether, boolean widowControl, presence-aware boolean contextualSpacing, 0-through-9 outlineLevel, pageBreakBefore, and presence-aware boolean suppressLineNumbers with the same direct/style inheritance and fail-closed source rules as document.addParagraph.
- `runStyle` (object) — Optional modeled formatting for the single ordinary run.

**Schema returns:**

- `paragraph` (DocumentParagraphBlock) — Appended canonical body-level block w:sdt around one paragraph/run. Multi-run, inline-field/control, non-text, nested, locked, placeholder, repeating-section, and data-bound profiles fail closed; use documentTableCell.addTextContentControl for the separate canonical cell-level profile.

#### `document.addBookmark`

Wrap exactly one paragraph-like block in a native Word bookmark for inspect, resolve, and internal hyperlinks. Recognized imported whole-block bookmarks are exposed with source identity but remain fixed-topology/read-only; cross-block, nested, crossing, table-cell, and otherwise complex ranges stay opaque-preserved and fail closed on mutation.

**Schema parameters:**

- `target` (string|object) required — Paragraph-like block ID/facade to wrap. Canonical authoring does not accept table cells or multi-block ranges.
- `name` (string) required — Unique case-insensitive Word bookmark name: ASCII letter first, then letters, digits, or underscore, at most 40 characters.
- `endTarget` (string|object) — Optional end block. Canonical authoring requires it to be the same block as target.
- `nativeId` (number) — Optional unsigned 32-bit Word bookmark numeric identity for source-free authoring; imported identity is source-bound.

**Schema returns:**

- `bookmark` (DocumentBookmark) — Native whole-block bookmark. Recognized imports are inspectable/resolvable but fixed-topology and read-only.

#### `document.addChange`

Append one bounded whole-paragraph tracked insertion or deletion. OfficeKit authors native w:ins/w:del markup and permits fixed-topology imported text/author/date edits; mixed or nested revision graphs remain source-bound.

**Schema parameters:**

- `changeType` (string) required — insert or delete.
- `text` (string) required — Revision text.
- `author` (string) — Revision author.
- `date` (string) — Revision timestamp.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `change` (DocumentChangeBlock) — Appended bounded tracked-change block authored as native whole-paragraph w:ins or w:del markup.

#### `document.addCitation`

Add a whole-paragraph bibliography-backed citation exported as a native w:fldSimple CITATION field plus a bounded bookmark. Recognized imports allow display-text edits while source tags and topology remain fixed.

**Schema parameters:**

- `text` (string) required — Visible citation text.
- `metadata` (object) required — Structured citation metadata containing a bounded ASCII tag that resolves to document.bibliographySources.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `citation` (DocumentCitationBlock) — Native whole-paragraph w:fldSimple CITATION block. Imported display text is editable while its tag and topology remain source-bound.

#### `document.addComment`

Attach a whole-paragraph Word comment. Classic roots remain minimal; bounded modern roots may carry resolved, durable/UTC, and provider-person metadata through OfficeKit.

**Schema parameters:**

- `target` (string|object) required — Stable block ID or block facade.
- `text` (string) required — Comment text.
- `author` (string) — Comment author.
- `initials` (string) — Author initials written to w:initials; derived deterministically from author when omitted.
- `date` (string) — Optional ISO-style comment timestamp written to w:date.
- `resolved` (boolean) — Optional w15:done state. Its presence selects the bounded modern comment graph.
- `parentId` (string) — Root comment model ID for a direct reply; prefer document.replyToComment().
- `paraId` (string) — Optional w14/w15 paragraph identity from 00000001 through 7FFFFFFF; generated deterministically when omitted for a modern source-free graph.
- `durableId` (string) — Optional Office 2019 durable identity from 00000001 through 7FFFFFFE; generated for the complete thread when required.
- `dateUtc` (string) — Optional ISO 8601 Office 2021 UTC metadata.
- `person` (object) — Optional presence identity for this author: providerId is 1-100 characters and userId is 1-300. Every comment with the same author must use the same identity or omit it consistently.
- `intelligentPlaceholder` (boolean) — Optional Office 2021 intelligent-placeholder flag.

**Schema returns:**

- `comment` (DocumentComment) — Attached classic or bounded modern whole-paragraph root comment. Rich bodies and irregular support-part graphs fail closed.

#### `document.addDeletion`

Append one bounded whole-paragraph tracked deletion using native w:del/w:delText markup. For one exact in-paragraph replacement in existing source bytes, use DocumentFile.addTrackedReplacement; mixed, moved, nested, and property-level revisions remain outside the bounded profile.

**Schema parameters:**

- `text` (string) required — Deleted text.
- `author` (string) — Revision author.
- `date` (string) — Revision timestamp.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `change` (DocumentChangeBlock) — Appended bounded whole-paragraph tracked deletion.

#### `document.addEndnote`

Append one native plain-text endnote with 1 through 16 canonical physical paragraphs at the end of one paragraph or list item. Recognized imported canonical endnotes permit fixed-count body-text edits only; anchor, kind, native ID, and note topology remain source-bound.

**Schema parameters:**

- `target` (string|DocumentParagraphBlock|DocumentListItemBlock) required — Paragraph or list-item ID/facade whose final run receives the native endnote reference.
- `text` (string) — One physical plain-text endnote paragraph, 1 through 1,000,000 XML-safe characters. Required unless paragraphs is supplied; it must equal the LF-joined paragraphs when both are supplied.
- `paragraphs` (string[]) — Optional 1 through 16 physical plain-text endnote paragraphs. Every item must be non-empty, XML-safe, and contain no CR/LF; imported notes keep this count source-bound.
- `name` (string) — Optional inspectable note name.
- `nativeId` (number) — Optional positive 32-bit native endnote ID for source-free authoring; imported identity is source-bound.

**Schema returns:**

- `endnote` (DocumentNote) — Native bounded endnote. Canonical imports expose paragraphs plus LF-joined text and allow fixed-count plain-text edits only.

#### `document.addField`

Append a bounded w:fldSimple block for PAGE, NUMPAGES, SECTION, date/time, and selected document-property commands. External-content and arbitrary reference commands fail closed.

**Schema parameters:**

- `instruction` (string) required — Bounded simple Word field instruction such as PAGE, NUMPAGES, SECTION, DATE, or a supported document-property command. Use addBibliography() rather than raw BIBLIOGRAPHY authoring.
- `display` (string) — Visible fallback/result text.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `field` (DocumentFieldBlock) — Appended field block.

#### `document.addFooter`

Add a default, first-page, or even-page DOCX footer, optionally section-scoped. Source-free input may be plain text, one legacy simple field, or a 2-through-32 ordered literal/simple-field segment sequence in one native paragraph. Imported multi-segment page furniture is inspectable and exactly preserved on no-op export but remains source-bound/read-only.

**Schema parameters:**

- `text` (string|HeaderFooterSegment[]) required — Plain footer text, or 2 through 32 ordered { text } / { field: { instruction, display } } items for one source-free native paragraph. Segment display concatenation becomes footer.text.
- `name` (string) — Inspectable block name.
- `styleId` (string) — Named style ID.
- `fieldInstruction` (string) — Optional legacy one-simple-field instruction such as PAGE or NUMPAGES. Mutually exclusive with segment input.
- `referenceType` (string) — default, first, or even section reference type.
- `sectionIndex` (number) — Zero-based target section. Omit to bind to the final section for backward compatibility.
- `activateVariant` (boolean) — Set false to preserve a dormant first/even reference without enabling different-first-page or even/odd behavior.

**Schema returns:**

- `footer` (DocumentHeaderFooterBlock) — Appended footer block.

#### `document.addFootnote`

Append one native plain-text footnote with 1 through 16 canonical physical paragraphs at the end of one paragraph or list item. Recognized imported canonical footnotes permit fixed-count body-text edits only; anchor, kind, native ID, and note topology remain source-bound.

**Schema parameters:**

- `target` (string|DocumentParagraphBlock|DocumentListItemBlock) required — Paragraph or list-item ID/facade whose final run receives the native footnote reference.
- `text` (string) — One physical plain-text footnote paragraph, 1 through 1,000,000 XML-safe characters. Required unless paragraphs is supplied; it must equal the LF-joined paragraphs when both are supplied.
- `paragraphs` (string[]) — Optional 1 through 16 physical plain-text footnote paragraphs. Every item must be non-empty, XML-safe, and contain no CR/LF; imported notes keep this count source-bound.
- `name` (string) — Optional inspectable note name.
- `nativeId` (number) — Optional positive 32-bit native footnote ID for source-free authoring; imported identity is source-bound.

**Schema returns:**

- `footnote` (DocumentNote) — Native bounded footnote. Canonical imports expose paragraphs plus LF-joined text and allow fixed-count plain-text edits only.

#### `document.addHeader`

Add a default, first-page, or even-page DOCX header, optionally section-scoped. Source-free input may be plain text, one legacy simple field, or a 2-through-32 ordered literal/simple-field segment sequence in one native paragraph. Imported multi-segment page furniture is inspectable and exactly preserved on no-op export but remains source-bound/read-only.

**Schema parameters:**

- `text` (string|HeaderFooterSegment[]) required — Plain header text, or 2 through 32 ordered { text } / { field: { instruction, display } } items for one source-free native paragraph. Segment display concatenation becomes header.text.
- `name` (string) — Inspectable block name.
- `styleId` (string) — Named style ID.
- `fieldInstruction` (string) — Optional legacy one-simple-field instruction such as PAGE or NUMPAGES. Mutually exclusive with segment input.
- `referenceType` (string) — default, first, or even section reference type.
- `sectionIndex` (number) — Zero-based target section. Omit to bind to the final section for backward compatibility.
- `activateVariant` (boolean) — Set false to preserve a dormant first/even reference without enabling different-first-page or even/odd behavior.

**Schema returns:**

- `header` (DocumentHeaderFooterBlock) — Appended header block.

#### `document.addHyperlink`

Append a native w:hyperlink backed by an external relationship or internal bookmark anchor; native import restores URL/anchor, relationship identity, tooltip, and history state.

**Schema parameters:**

- `text` (string) required — Visible link text.
- `url` (string|DocumentBookmark) — External HTTP(S) URL, #bookmark name, or bookmark facade.
- `anchor` (string|DocumentBookmark) — Internal bookmark name/facade; mutually exclusive with an external URL.
- `tooltip` (string) — Optional Word hyperlink tooltip, at most 260 characters.
- `history` (boolean) — Whether Word records the hyperlink as visited; defaults to true.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `hyperlink` (DocumentHyperlinkBlock) — Appended external or internal hyperlink block.

#### `document.addImage`

Append an inspectable embedded PNG/JPEG image. Images are inline by default; an explicit bounded placement authors a native foreground wp:anchor with square or top-and-bottom wrapping.

**Schema parameters:**

- `dataUrl` (string) — Embedded image data URL.
- `uri` (string) — External image URI metadata.
- `prompt` (string) — Generation/source prompt metadata.
- `alt` (string) — Alternative text.
- `widthPx` (number) — Rendered width in pixels.
- `heightPx` (number) — Rendered height in pixels.
- `styleId` (string) — Named paragraph style ID.
- `placement` (object) — Optional image placement. Omit it (or use { type: 'inline' }) for inline flow. The bounded floating profile is { type: 'floating', horizontal: { relativeTo: 'margin'|'page'|'column', offsetPx }, vertical: { relativeTo: 'margin'|'page'|'paragraph', offsetPx }, wrap: 'square'|'topAndBottom', wrapSide?: 'bothSides'|'left'|'right'|'largest', distanceFromTextPx?: { top, right, bottom, left } }. wrapSide is square-only; offsets are bounded to +/-10000 px and text distances to 0..10000 px.

**Schema returns:**

- `image` (DocumentImageBlock) — Appended embedded image block. Recognized imported floating images permit only fixed-topology placement edits; inline/floating transitions and unsupported anchor graphs fail closed.

#### `document.addInsertion`

Append one bounded whole-paragraph tracked insertion using native w:ins markup. For one exact in-paragraph replacement in existing source bytes, use DocumentFile.addTrackedReplacement; mixed, moved, nested, and property-level revisions remain outside the bounded profile.

**Schema parameters:**

- `text` (string) required — Inserted text.
- `author` (string) — Revision author.
- `date` (string) — Revision timestamp.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `change` (DocumentChangeBlock) — Appended bounded whole-paragraph tracked insertion.

#### `document.addListItem`

Append a numbered, character-bulleted, or bounded picture-bulleted list item using native DOCX numbering definitions. Picture markers are shared numbering-level resources: every item using the same numberingId and level must agree, and recognized imported edits must update the complete group without changing embedded-versus-external source kind.

**Schema parameters:**

- `text` (string) required — List item text.
- `listType` (string) — bullet or numbered.
- `level` (number) — Zero-based list nesting level.
- `numberFormat` (string) — OOXML numbering format such as bullet, decimal, upperLetter, lowerRoman, or ordinal.
- `start` (number) — Positive starting value for this numbering level.
- `levelText` (string) — OOXML level text template using placeholders such as %1 or %2.
- `numberingId` (number|string) — Optional list-instance identity used to group levels during export and preserved by native import.
- `abstractNumberingId` (number|string) — Optional abstract numbering identity used to share one compatible multilevel definition across list instances; preserved by native import.
- `numberingStyleId` (string) — Optional Word numbering-style identity resolved through styleLink/numStyleLink and flattened safely on second export.
- `pictureBullet` (string|object) — Optional embedded PNG/JPEG/GIF base64 data URL, absolute HTTP(S) URI, or { dataUrl|uri, widthPt|sizePt, heightPt, alt } marker. Width and height are 4 through 72 points; external resources are referenced but never fetched. All list items sharing numberingId and level must use the same marker. Recognized imported markers allow only a coherent full-group edit with the original embedded/external source kind; irregular VML and broader inherited numbering graphs fail closed.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `listItem` (DocumentListItemBlock) — Appended native-numbering list item, including normalized pictureBullet metadata when configured.

#### `document.addParagraph`

Append a styled paragraph with optional run spans and bounded direct paragraph formatting, including canonical solid shading and solid paragraph borders, presence-aware contextual spacing, and line-number suppression.

**Schema parameters:**

- `text` (string) required — Paragraph text.
- `styleId` (string) — Named paragraph style ID.
- `name` (string) — Inspectable block name.
- `paragraphFormat` (object) — Optional modeled paragraph formatting. keepNext keeps this paragraph with its following paragraph, keepLinesTogether keeps one paragraph from splitting across pages, widowControl asks the host to avoid orphan/widow lines, and pageBreakBefore asks the host to begin it on a new page; these independent pagination directives do not calculate pages. keepLinesTogether and widowControl accept only boolean true or false, where false explicitly overrides an inherited style setting and omission inherits. contextualSpacing accepts only boolean true or false: true suppresses before/after spacing between adjacent paragraphs with the same style; false explicitly overrides an inherited style setting and omission inherits. shadingFill accepts one #RRGGBB solid paragraph background and authors canonical w:shd with clear pattern and auto foreground. A recognized ordinary direct paragraph can add, change, or clear that fixed fill inside its modeled direct-formatting profile; theme colors, patterns, and imported style-catalog edits stay source-bound. borders accepts one nonempty object with top/left/bottom/right/between/bar edges; every edge requires a #RRGGBB color and integer size 2 through 96 in eighths of a point, with optional integer space 0 through 31 in points. It writes canonical w:pBdr children with w:val=single. A recognized ordinary direct paragraph can add, replace, or clear the whole profile; source-free paragraph styles may author it, while imported style catalogs, themes, patterns, frame/shadow, other line styles, duplicate edges, children, extensions, and unknown attributes stay source-bound. outlineLevel accepts integer 0 through 9: 0 through 8 are native outline levels, while 9 explicitly clears an inherited level; omission inherits. Canonical direct w:keepNext/w:keepLines/w:pageBreakBefore/w:widowControl/w:contextualSpacing/w:shd/w:pBdr/w:outlineLvl/w:suppressLineNumbers leaves are editable under their bounded profiles; duplicate, child-bearing, extension-bearing, or invalid lexical markup stays source-owned and semantic replacement fails closed. suppressLineNumbers true excludes this paragraph from section line-number display and calculation; false explicitly overrides inherited style suppression; omission inherits the named style/default.
- `runs` (object[]) — Optional run spans whose style may include runStyleId plus direct/theme formatting. A run may carry a bounded contentControl { id, tag, alias, nativeId?, controlType?, checked?, choices?, selectedValue?, value? } or inlineField { instruction, bookmarkName?, bookmarkNativeId? }.

**Schema returns:**

- `paragraph` (DocumentParagraphBlock) — Appended paragraph block with stable ID.

#### `document.addSection`

Append a DOCX section break with page size, orientation, margins, binding gutter, canonical equal-width or explicit-width columns, bounded page-number start/format, and break-type metadata backed by w:sectPr. Imported geometry and page numbering are writable only when their native markup is canonical.

**Schema parameters:**

- `breakType` (string) — Section break type such as nextPage or continuous.
- `orientation` (string) — portrait or landscape.
- `pageSize` (object) — Page width/height in twentieths of a point.
- `margins` (object) — Top/right/bottom/left margins plus optional non-negative binding gutter in twentieths of a point. document.settings.gutterAtTop chooses top-edge versus binding-side placement.
- `columns` (object) — Optional canonical text columns. Equal-width profile: { count: 1–45, spacing, separator }. Explicit-width profile: { definitions: [{ width, spacing }], separator }, with 1–45 ordered definitions, positive widths, and non-negative spacing-after values. All values are twentieths of a point; margins, binding-side gutter, widths, and gaps must fit the page content width. The two profiles cannot be mixed; ambiguous or extension-bearing w:cols graphs stay source-owned.
- `lineNumbering` (object) — Optional canonical line numbering before each text column: { countBy?: integer 1..32767, start?: integer 0..32767, distance?: integer 0..31680, restart?: 'newPage'|'newSection'|'continuous' }. An empty object defaults countBy to 1; start is the zero-based native value, so the first displayed line is start + 1. distance is in twentieths of a point. Use paragraphFormat.suppressLineNumbers for presence-aware paragraph/style suppression. Duplicate leaves, children, unknown values, or extension-bearing w:lnNumType markup stay source-owned.
- `pageNumbering` (object) — Optional canonical section numbering: { start?: integer 0..2147483647, format?: 'decimal'|'upperRoman'|'lowerRoman'|'upperLetter'|'lowerLetter' }. At least one property is required; omitting start continues the prior sequence. This controls PAGE-field presentation but does not add or refresh a field. Chapter numbering, unsupported formats, duplicate leaves, children, or extension-bearing w:pgNumType markup stay source-owned.

**Schema returns:**

- `section` (DocumentSectionBlock) — Appended section break block. inspect reports editable=false when imported mirrorMargins/gutterAtTop mode markup, section-column topology, line-number markup, or page-number markup is not canonical.

#### `document.addTable`

Append a Word-style table with physical cell values, optional logical merge geometry, fixed-layout width/margin/border styling, optional left/center/right table placement, optional uniform top/center/bottom physical-cell alignment, non-clipping per-row minimum heights, an optional native repeating-header prefix, individual rows kept together across pages, and optional non-visible table alternative text.

**Schema parameters:**

- `values` (unknown[][]) required — Table cell value matrix.
- `gridColumns` (number) — Logical Word table-grid width. Required for explicit authored geometry; otherwise derived from values.
- `cells` (object[]) — One record per physical value cell with zero-based row/column, gridColumn, columnSpan, rowSpan, verticalMerge none/restart/continue, and editability evidence. OfficeKit can author complete, contiguous, conforming geometry and keeps imported geometry source-bound.
- `name` (string) — Inspectable table name.
- `styleId` (string) — Table style ID.
- `widthDxa` (number) — Table width in twentieths of a point.
- `indentDxa` (number) — Leading table indent in twentieths of a point.
- `horizontalAlignment` ("left" | "center" | "right") — Optional native table-level w:jc placement. Omit for Word's normal left default. center/right require indentDxa 0, preventing a host-dependent conflict between table placement and table indentation.
- `columnWidthsDxa` (number[]) — One width per logical table-grid column in twentieths of a point; values must sum to widthDxa.
- `cellMarginsDxa` (object) — Cell margins in twentieths of a point.
- `borderColor` (string) — Table border color.
- `borderSize` (number) — Uniform border width in eighths of a point; zero disables borders.
- `headerFill` (string) — First-row fill color. This is visual styling only and does not mark rows as native repeat headers.
- `verticalAlignment` ("top" | "center" | "bottom") — Optional uniform physical-cell alignment. Omit for Word's native top default; a recognized imported table can edit or clear it only when every physical cell has the same canonical direct w:vAlign profile.
- `headerRowCount` (number) — Number of contiguous leading physical rows to mark as native Word w:tblHeader repeat headers; 0 through the table row count, default 0.
- `keepTogetherRows` (number[]) — Zero-based physical table rows that must not split across pages through native w:cantSplit. Values form a deduplicated ascending set within the table row count; this does not group rows or calculate pagination.
- `minimumRowHeightsDxa` (number[]) — One non-negative integer DXA value per physical row. Zero omits the native height leaf; a positive value writes canonical w:trHeight hRule=atLeast so wrapped content may expand instead of being clipped.
- `accessibility` (object) — Optional non-visible Word table alternative text: { title?: string, description?: string }. Each value is 1 through 32767 XML-safe characters and maps to one canonical w:tblCaption/@w:val or w:tblDescription/@w:val leaf. It does not create a visible caption paragraph or change table layout.

**Schema returns:**

- `table` (DocumentTableBlock) — Appended table block.

#### `document.addTableOfContents`

Append one canonical one-paragraph complex TOC field with bounded heading levels/switches and enable the native updateFields-on-open hint by default. Refreshed cross-paragraph result graphs remain opaque/source-bound and read-only.

**Schema parameters:**

- `levels` (string) — Ascending heading-level range such as 1-3; defaults to 1-3.
- `minLevel` (number) — Minimum level when levels is omitted.
- `maxLevel` (number) — Maximum level when levels is omitted.
- `hyperlinks` (boolean) — Include the canonical \h switch; defaults to true.
- `hidePageNumbersInWeb` (boolean) — Include the canonical \z switch; defaults to true.
- `useOutlineLevels` (boolean) — Include the canonical \u switch; defaults to true.
- `display` (string) — Cached placeholder result shown until a compatible host refreshes the TOC.
- `updateFields` (boolean) — Enable the updateFields-on-open hint; defaults to true.
- `styleId` (string) — Paragraph style ID.

**Schema returns:**

- `field` (DocumentFieldBlock) — Canonical complex TOC placeholder with complex=true.

#### `document.addWatermark`

Add one canonical VML text watermark to a section/header-reference scope. Recognized imported watermarks permit text-only edits or whole-object removal; adding to an imported package, changing scope, shared headers, multiple objects, DrawingML, images, and irregular VML fail closed.

**Schema parameters:**

- `text` (string) required — Nonblank XML-safe watermark text, 1 through 256 characters.
- `id` (string) — Optional object ID. IDs locate this model object; they are not persistent document identity across unrelated imports.
- `referenceType` (string) — default, first, or even header reference scope; defaults to default.
- `sectionIndex` (number) — Zero-based target section; defaults to 0.

**Schema returns:**

- `watermark` (DocumentWatermark) — One canonical VML text watermark. Only one object is allowed per section/reference scope.

#### `document.applyDesignPreset`

Apply a clean-room report or memo design preset that updates named styles for consistent DOCX export and SVG/layout previews.

**Schema parameters:**

- `name` (string) required — report, memo, or a custom preset name.
- `styles` (object) — Style overrides merged into the preset.

**Schema returns:**

- `document` (DocumentModel) — The mutated document facade.

#### `document.auditAccessibility`

Audit modeled Word headings, image alternative text, table header semantics, and hyperlink text with stable block locators. Machine-checkable defects remain separate from table-purpose and link-purpose manual review, and the result never claims Word Accessibility Checker or WCAG conformance.

**Schema parameters:**

- `maxChars` (number) — Maximum bounded NDJSON size across machine issues and manual-review records.

**Schema returns:**

- `report` (object) — A host-neutral report with machineCheckPassed, conformanceClaimed: false, manualReviewRequired, stable block locators, counts, machine issues, and separate manual checks. Skipped heading levels, empty image alternative text, tables without a repeating header-row prefix, and empty hyperlink text fail the machine check. Missing table descriptions plus generic/raw-URL link text remain manual author-intent checks; the audit does not claim Word Accessibility Checker or WCAG conformance.

#### `document.contentControls`

List typed mutable handles for recognized inline or table-cell plain-text, checkbox, drop-down, combo-box, and date controls plus block plain-text controls, with explicit placement and model/native identity.

**Schema returns:**

- `controls` (DocumentContentControlHandle[]) — Fresh typed handles for recognized block/inline/table-cell text, checkbox, drop-down, combo-box, and date controls. placement is block, inline, or tableCell; runIndex is present only for inline controls and row/column only for table-cell controls. Tag/alias plus type-specific text, checked, selectedValue, value, or dateValue are mutable; list choices, controlType, nativeId, native date profile, symbol declarations, placement, and topology are source identity.

#### `document.fillContentControls`

Transactionally fill every recognized block, inline, or table-cell plain-text control matching an object or Map of tag-to-string entries. Checkbox, drop-down, combo-box, and date tags do not silently accept text.

**Schema parameters:**

- `values` (object|Map) required — Tag-to-string value mapping. Duplicate tags fill every matching control.
- `strict` (boolean) — Unknown tags fail before mutation; defaults to true. Checkbox, drop-down, combo-box, and date tags are never matched by this text primitive.

**Schema returns:**

- `result` (object) — Structured { updated, matchedTags, missingTags } result.

#### `document.fontFamilies`

Return a fresh sorted, case-insensitively deduplicated list of document theme and explicit run/style font families.

**Schema returns:**

- `families` (string[]) — Font-family inventory; mutating the returned array does not mutate the document.

#### `document.inspect`

Emit bounded NDJSON for document blocks including typed block/inline plain-text and inline checkbox/list/date content controls with explicit placement, fields, tracked changes, bookmark ranges, footnotes/endnotes, bibliography sources, comments, styles, headers/footers with sourceBound/editable evidence, canonical text watermarks, and layout; narrow with search/target anchors and fields with include/exclude.

**Examples:**

- document.inspect({ kind: 'paragraph,comment', target: comment.id, maxChars: 4000 })

**Options:**

- kind
- search
- target/targetId/id/anchor
- before/after/context
- include/fields
- exclude/omit
- maxChars

**Schema parameters:**

- `kind` (string) — Comma-separated block/tableCell/comment/watermark/style/textRange/layout kinds; paragraph and table-cell records expose textEditable/textPatchable capability evidence.
- `search` (string) — Case-insensitive record filter.
- `target` (string) — Stable target ID/anchor.
- `before` (number) — Context records before matches.
- `after` (number) — Context records after matches.
- `include` (string) — Comma-separated fields to keep.
- `exclude` (string) — Comma-separated fields to omit.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `inspection` (object) — Bounded { ndjson, truncated } inspection result.

**Returns:**

{ ndjson, truncated } bounded NDJSON records

#### `document.layoutJson`

Return page-aware layout JSON with block bounding boxes, section/page ordinals, effective inherited header/footer selections, styles, and target/search slicing.

**Schema parameters:**

- `pageWidth` (number) — Modeled page width in pixels.
- `pageHeight` (number) — Modeled page height in pixels.
- `margin` (number) — Modeled page margin in pixels.
- `target` (string) — Stable target ID/anchor.
- `search` (string) — Case-insensitive element filter.
- `before` (number) — Context elements before matches.
- `after` (number) — Context elements after matches.

**Schema returns:**

- `layout` (object) — Page-aware document layout tree.

#### `document.materializeFields`

Transactionally compute canonical inline SEQ counters and REF cached results from native bookmark targets, with dry-run evidence and strict missing-target failure. PAGEREF remains skipped because trustworthy page numbers require a real pagination host.

**Schema parameters:**

- `types` (string|string[]) — SEQ and/or REF cached-result types; defaults to both. PAGEREF is rejected when requested.
- `dryRun` (boolean) — Plan and report every cache change without mutating the document.
- `strict` (boolean) — Reject unresolved or duplicate bookmark targets before any mutation; defaults to true.

**Schema returns:**

- `result` (object) — Structured { dryRun, updated, wouldUpdate, seqFields, refFields, skippedPageReferences, missingBookmarks, changes } result.

#### `document.render`

Render an SVG preview by default, return layout JSON with { format: 'layout' }, or use { source: 'docx', renderer } to feed native DOCX into LibreOffice/native Office render adapters for PDF/PNG outputs.

**Schema parameters:**

- `format` (string) — svg by default, layout, docx, pdf, png, or another renderer output.
- `source` (string) — Set to docx to render exported DOCX bytes.
- `renderer` (function) — Optional LibreOffice/native/raster renderer adapter.
- `pageWidth` (number) — Modeled SVG/layout page width.
- `pageHeight` (number) — Modeled SVG/layout page height.

**Schema returns:**

- `blob` (FileBlob) — SVG, layout JSON, DOCX, or converted renderer output.

#### `document.replyToComment`

Add one source-free direct reply to a root comment. OfficeKit authors the bounded commentsExtended graph; nested replies and imported topology changes fail closed.

**Schema parameters:**

- `parent` (string|DocumentComment) required — Existing parent comment ID or facade.
- `text` (string) required — Reply text.
- `author` (string) — Reply author.
- `initials` (string) — Reply author initials.
- `date` (string) — Optional reply timestamp.
- `resolved` (boolean) — Reply resolution state.
- `durableId` (string) — Optional preserved Office 2019 durable comment identity.
- `dateUtc` (string) — Optional Office 2021 UTC timestamp.
- `person` (object) — Optional providerId/userId presence identity for the reply author.

**Schema returns:**

- `comment` (DocumentComment) — Source-free direct reply authored through commentsExtended. Replies to replies and additions/removals in an imported fixed-topology thread fail closed.

#### `document.resolve`

Resolve stable document, block, table-cell, content-control, bookmark, footnote/endnote, bibliography source ID/tag, header/footer, watermark, comment, style, and advertised text-range IDs.

**Schema parameters:**

- `id` (string) required — Stable document, block, table-cell, header/footer, watermark, comment, style, or advertised text-range ID.

**Schema returns:**

- `object` (object|undefined) — Resolved editable facade/record or undefined.

#### `document.setCheckboxContentControls`

Transactionally set every recognized canonical checkbox control matching an object or Map of tag-to-boolean entries. Other control types do not silently coerce.

**Schema parameters:**

- `values` (object|Map) required — Tag-to-boolean checked-state mapping. Duplicate tags update every matching checkbox.
- `strict` (boolean) — Unknown checkbox tags fail before mutation; defaults to true.

**Schema returns:**

- `result` (object) — Structured { updated, matchedTags, missingTags } result.

#### `document.setComboBoxContentControls`

Transactionally set every recognized canonical combo-box control from a tag-to-value string mapping. Values may select a declared choice or provide bounded custom text; unknown tags and invalid values fail before mutation.

**Schema parameters:**

- `values` (object|Map) required — Tag-to-string value mapping. Each value may match one declared choice or be XML-safe custom text of 1 to 255 characters; duplicate tags update every matching combo-box.
- `strict` (boolean) — Unknown combo-box tags fail before mutation; defaults to true. All values are validated before any control changes.

**Schema returns:**

- `result` (object) — Structured { updated, matchedTags, missingTags } result.

#### `document.setDateContentControls`

Transactionally set every recognized canonical date control from a tag-to-YYYY-MM-DD mapping. Invalid Gregorian dates, unknown tags, and other control types fail before mutation.

**Schema parameters:**

- `values` (object|Map) required — Tag-to-string date mapping. Every value must be a real Gregorian date in exact YYYY-MM-DD form; duplicate tags update every matching date control.
- `strict` (boolean) — Unknown date tags fail before mutation; defaults to true. All dates are validated before any control changes.

**Schema returns:**

- `result` (object) — Structured { updated, matchedTags, missingTags } result.

#### `document.setDropdownContentControls`

Transactionally set every recognized canonical drop-down control from a tag-to-choice-value string mapping. Unknown tags or values outside the declared choice table fail before mutation.

**Schema parameters:**

- `values` (object|Map) required — Tag-to-string selected-value mapping. Each value must exactly match one declared choice; duplicate tags update every matching drop-down.
- `strict` (boolean) — Unknown drop-down tags fail before mutation; defaults to true. All selected values are validated before any control changes.

**Schema returns:**

- `result` (object) — Structured { updated, matchedTags, missingTags } result.

#### `document.setSectionSettings`

Set per-section Word behavior such as different-first-page header/footer activation without changing preserved header/footer references.

**Schema parameters:**

- `sectionIndex` (number) required — Zero-based section index from 0 through the number of section-break blocks.
- `differentFirstPage` (boolean) — Whether the section activates first-page header/footer references through w:titlePg.

**Schema returns:**

- `document` (DocumentModel) — Document facade with normalized per-section settings.

#### `document.setSettings`

Set model settings. evenAndOddHeaders, mirrorMargins, gutterAtTop, trackRevisions, the updateFields refresh hint, and bounded passwordless documentProtection are inside the bounded OfficeKit DOCX profile. Irregular page-margin mode markup and password/cryptographic protection variants stay source-owned and fail closed on replacement.

**Schema parameters:**

- `settings` (object) required — Partial settings object. evenAndOddHeaders, mirrorMargins, gutterAtTop, updateFields, and trackRevisions are booleans. documentProtection accepts false/null/off to remove the element, none/readOnly/comments/trackedChanges/forms, or { edit, enforcement, formatting }; password hashes, cryptographic attributes, IRM, permission exceptions, and irregular mirrorMargins/gutterAtTop markup are unsupported/source-owned. Structurally irregular page-margin mode markup also blocks sibling settings edits and makes imported section geometry read-only when exact reserialization cannot be proved.

**Schema returns:**

- `document` (DocumentModel) — Document facade with normalized facing-page/binding-gutter/header/tracking/refresh settings; updateFields is a refresh request, and passwordless documentProtection is an editing restriction rather than encryption or access control.

#### `document.styles.effective`

Resolve a named document style through basedOn inheritance so inspect/layout/render/DOCX export share the same effective style metadata.

**Schema parameters:**

- `styleId` (string) required — Named style ID to resolve through basedOn inheritance.

**Schema returns:**

- `style` (object|undefined) — Resolved effective style or undefined.

#### `document.textRange`

Inspect or resolve stable textRange anchors such as blockId/text and tableId/cell/row/column/text. Assignment is limited to fully editable text; replace() also supports explicitly advertised source-bound literal patches.

**Schema parameters:**

- `id` (string) required — Stable blockId/text or tableId/cell/row/column/text range ID.

**Schema returns:**

- `textRange` (TextRange|undefined) — Advertised text-range facade. Assignment requires textEditable; replace() may instead use the narrower textPatchable contract.

#### `document.verify`

Return QA issues for invalid/duplicate content-control IDs and native IDs, malformed tags/aliases, invalid block-control profiles, fake lists, invalid links/citations/bibliography sources, malformed tracked changes, duplicate/dangling/reversed bookmark ranges, invalid footnotes/endnotes, unknown styles, malformed tables, bad images/sections, invalid watermark IDs/scopes/text, dangling comments, visual overflow, and prose-like table cells.

**Schema parameters:**

- `visualQa` (boolean) — Include modeled layout overflow checks.
- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — Document semantic/layout QA result.

#### `documentComment.reopen`

Clear the resolved state of a bounded modern comment without changing its root/reply topology or durable identity.

**Schema returns:**

- `comment` (DocumentComment) — The same comment facade with resolved=false; root/reply, paragraph, durable, UTC, and people identity remain fixed.

#### `documentComment.resolve`

Set resolved=true for a bounded modern comment. Imported edits re-prove source hashes and commentsExtended topology while keeping thread identity fixed.

**Schema returns:**

- `comment` (DocumentComment) — The same comment facade with resolved=true; imported edits re-prove source hashes and commentsExtended topology.

#### `DocumentFile.addTrackedReplacement`

Add one exact replacement inside a direct body paragraph or bounded table-cell paragraph to hash-bound DOCX source bytes as adjacent native w:del/w:ins runs. A structured paragraph/tableCell selector, full expected text, and one unique literal contained in either one ordinary run or adjacent run fragments with identical w:rPr preserve source formatting; mixed formatting and broader topologies fail closed with exact changed-part audit.

**Schema parameters:**

- `docx` (FileBlob|Uint8Array|ArrayBuffer) required — Original DOCX bytes. OfficeKit edits this package directly and never rebuilds it from the imported JavaScript model.
- `target` (object) — Preferred structured selector: { kind: 'paragraph', blockIndex } or { kind: 'tableCell', blockIndex, row, column }. Table row/column are zero-based physical indexes from the exact imported table block.
- `targetBlockIndex` (number) — Compatibility selector for a direct body paragraph. Omit when target is supplied; the two forms are mutually exclusive.
- `expectedText` (string) required — Exact full text of the target paragraph or single-paragraph table cell; stale text fails closed before mutation.
- `search` (string) required — Non-empty literal that must occur exactly once. It may occupy one ordinary w:r/w:t or adjacent non-empty ordinary runs only when their exact w:rPr markup is identical; duplicate, empty-run-gap, and mixed-format spans fail closed.
- `replacement` (string) required — Non-empty replacement text written in a native adjacent w:ins run with the source run formatting.
- `author` (string) required — Revision author, 1 through 255 characters without control characters.
- `date` (string) — Optional ISO 8601 revision timestamp.
- `expectedSourceSha256` (string) required — Lowercase 64-hex SHA-256 of the exact input bytes; JavaScript and OfficeKit both verify it.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — Source-preserving DOCX with metadata.trackedReplacement containing the re-proved structured target, source/output and paragraph-element hashes, UTF-16 text hashes/counts, matchedSourceRunCount, package-local native revision IDs, semantic/body indexes, and the exact changed-part list. Only word/document.xml may change.

#### `DocumentFile.exportDocx`

Export DocumentModel to DOCX through the single bundled OfficeKit codec. Only limits is accepted; legacy codec and lossy-fallback options fail explicitly.

**Schema parameters:**

- `document` (DocumentModel) required — Document facade to serialize.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — DOCX package bytes.

#### `DocumentFile.finalizeRevisions`

Accept or reject bounded direct whole-paragraph one-run revisions and exact adjacent in-paragraph w:del + w:ins pairs from source bytes, including same-format fragmented deletions in direct body paragraphs or bounded table-cell paragraphs. Mandatory SHA-256 binding, decompression budgets, exact changed-part audit, and fail-closed graph checks prevent silent model reconstruction or broad package mutation.

**Schema parameters:**

- `docx` (FileBlob|Uint8Array|ArrayBuffer) required — Original DOCX bytes. The native codec operates directly on this package rather than rebuilding it from a JavaScript model.
- `mode` (string) required — accept or reject.
- `expectedSourceSha256` (string) required — Lowercase 64-hex SHA-256 of the exact input bytes; JavaScript and OfficeKit both verify it.
- `keepTracking` (boolean) — Preserve an existing trackRevisions setting after finalization. Defaults to false and never enables a setting that was absent.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — Rewritten DOCX with metadata.revisionFinalization containing source/output hashes, insertion/deletion counts, tracking before/after, and exact changed parts. Direct body whole-paragraph one-run revisions plus exact adjacent deletion/insertion pairs in direct body paragraphs or bounded table cells are accepted; the deletion may retain multiple adjacent fragments only when every fragment and the single insertion have identical w:rPr. Mixed-format, nested, moved, property-level, non-body-story, irregular-table, malformed, or absent revisions fail closed.

#### `DocumentFile.importDocx`

Import relationship-driven core DOCX semantics through the single bundled OfficeKit codec. An imported header/footer advertises editable only for one direct unformatted text paragraph in a uniquely used source part; recognized ordered literal/simple-field page furniture is exposed as segments but remains source-bound/read-only and no-op preserved. PAGE/simple fields, rich, shared, inherited, and irregular page furniture stay read-only. Recognized inline controls, fields, revisions, notes, citations, simple tables, and exclusive canonical VML text-watermark paragraphs are fixed-topology editable; a canonical BIBLIOGRAPHY output field permits only its cached display text to change. Otherwise read-only paragraphs and complex table cells separately advertise textPatchable when at least one direct ordinary native text node can participate in a bounded literal patch. A unique literal may span adjacent same-format runs without rebuilding the surrounding graph.

**Schema parameters:**

- `docx` (FileBlob|Uint8Array) required — DOCX package bytes.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `document` (DocumentModel) — Imported document facade with editable core blocks, hash-bound direct-text header/footer capability evidence, recognized canonical text watermarks, and source-bound read-only advanced content.

#### `DocumentFile.inspectDocx`

Inspect bounded DOCX parts, content types, the required main-document/root officeDocument relationship, and namespace-aware source XML r:id/r:embed/r:link references after raw-input, part-count, decompression, and optional compression-ratio budgets; verifyCrc32 additionally checks ZIP entry CRCs.

**Schema parameters:**

- `docx` (FileBlob|Uint8Array) required — DOCX package bytes.
- `includeText` (boolean) — Include bounded XML/JSON/relationship previews.
- `maxPreviewChars` (number) — Maximum preview characters per textual part.
- `maxInputBytes` (number) — Maximum compressed input bytes checked before JSZip parses the package.
- `maxParts` (number) — Maximum package part count.
- `maxPartBytes` (number) — Maximum uncompressed bytes per part.
- `maxTotalBytes` (number) — Maximum total uncompressed package bytes.
- `maxCompressionRatio` (number) — Optional maximum declared uncompressed/compressed ZIP-entry ratio; zero or omitted disables this extra check.
- `verifyCrc32` (boolean) — Verify every ZIP entry CRC32 before inspecting package structure; use for untrusted retained inputs.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `package` (object) — DOCX package result with ok, issues, parts, records, and bounded NDJSON.

#### `DocumentFile.patchDocx`

Apply DOCX part patches with path traversal validation for settings, classic-comment anchors, commentsExtended/commentsIds/commentsExtensible/people parts, and numbering assignments; atomically reject dangling packages and invalid comment graphs.

**Examples:**

- await DocumentFile.patchDocx(docx, [{ path: 'customXml/review-note.xml', text: '<review>ok</review>' }])

**Schema parameters:**

- `docx` (FileBlob|Uint8Array) required — DOCX package bytes.
- `patches` (array|object) required — Path-validated package part edits with text/xml/json/bytes/remove.
- `maxInputBytes` (number) — Maximum compressed input bytes checked before JSZip parses the package.
- `maxPatchBytes` (number) — Per-part patch size limit.
- `maxParts` (number) — Maximum resulting package part count; the source part count is checked before inflation.
- `maxPartBytes` (number) — Maximum uncompressed bytes per source or resulting part.
- `maxTotalBytes` (number) — Maximum total uncompressed source or resulting package bytes.
- `maxCompressionRatio` (number) — Optional maximum declared uncompressed/compressed ZIP-entry ratio; zero or omitted disables this extra check.
- `syncContentTypes` (boolean) — Synchronize inferred or explicit content-type declarations; defaults to true.
- `syncRelationships` (boolean) — Remove relationships to deleted parts and apply relationship recipes; defaults to true.
- `syncSourceReferences` (boolean) — Apply opt-in standard sourceReference XML mutations for supported semantic recipes; defaults to true.
- `validateResult` (boolean) — Validate final content types and relationships atomically; defaults to true. Set false only for deliberate invalid-package fixtures.
- `recipe` (string|object) — Standard OOXML part recipe with optional source/id/target and sourceReference fields; DOCX supports settings mutations, section-scoped header/footer references, batch classic-comment anchors, commentsExtended/commentsIds/commentsExtensible/people relationships, and numbering assignments for block, paragraph, or table-cell targets.
- `sourceReference` (boolean|object) — Opt-in semantic XML mutation. Settings accepts trackRevisions/updateFields/evenAndOddHeaders/mirrorMargins/gutterAtTop booleans and passwordless documentProtection; comments accepts { anchors: [...] }; numbering accepts { assignments: [...] }.
- `relationship` (object) — Per-patch source/id/type/target/targetMode relationship recipe; explicit ID collisions require replaceExisting:true. relationships accepts an array.

**Schema returns:**

- `docx` (FileBlob) — Patched DOCX FileBlob with part/relationship/content-type/source-reference update counts and validation metadata.

#### `documentHeaderFooter.setSegments`

Atomically replace one source-free header/footer's ordered literal/simple-field sequence. The derived visible text must remain the concatenated segment displays; imported page furniture cannot use this mutation profile.

**Schema parameters:**

- `segments` (HeaderFooterSegment[]) required — 2 through 32 ordered { text } or { field: { instruction, display } } entries. At least one bounded simple field is required; literal/field displays derive text and no fieldInstruction may coexist.

**Schema returns:**

- `segments` (DocumentHeaderFooterBlock) — Mutated source-free header/footer block with derived text. Imported source-bound page furniture rejects this operation.

#### `DocumentModel.create`

Create a document with paragraph/character styles, formatted paragraphs/runs including canonical solid paragraph shading and bounded solid paragraph borders, canonical inline and one-paragraph table-cell plain-text, checkbox, drop-down, combo-box, and ISO/Gregorian date content controls, one-paragraph block plain-text controls, canonical inline SEQ/REF/PAGEREF fields, sections, headers/footers, canonical VML text watermarks, lists, TableGrid fixed-geometry tables, links, bounded whole-block bookmarks, 1-through-16-paragraph plain-text footnotes/endnotes, canonical bibliography-backed citations plus one source-free switch-free BIBLIOGRAPHY output placeholder, simple fields, a canonical complex TOC placeholder, bounded whole-paragraph tracked insertions/deletions, classic comments, bounded modern root/direct-reply threads, and PNG/JPEG images. Nested/irregular modern threads, rich comment bodies, multi-paragraph/rich/inline-within-cell/nested/data-bound/locked/placeholder table-cell SDTs, other nested/data-bound/locked/placeholder SDTs, irregular lists, localized dates, custom checkbox symbols, image/DrawingML/irregular VML watermarks, other complex field graphs, arbitrary table-style graphs, complex bookmark/note/revision graphs, and advanced settings remain unsupported or source-bound.

**Schema parameters:**

- `name` (string) — Document name.
- `designPreset` (string) — Initial design preset name.
- `theme` (object) — Word theme name, 12 scheme colors, and major/minor Latin, East-Asian, and complex-script fonts.
- `defaultRunStyle` (object) — Document-wide run properties serialized as w:docDefaults/w:rPrDefault and applied before named styles.
- `styles` (object) — Named paragraph/character styles plus imported table/numbering style records with optional basedOn inheritance and numberingId/numberingLevel linkage. Source-free table blocks may select TableGrid; arbitrary custom table-style graphs are not materialized.
- `paragraphs` (string[]) — Convenience paragraph list; the first paragraph uses Title style.
- `blocks` (object[]) — Ordered paragraph/list/table/link/field/citation/image/section/change block models. Paragraph runs may carry canonical inline SEQ/REF/PAGEREF fields; bibliography-backed citations and one-paragraph complex TOC placeholders also cross OfficeKit. Other field graphs remain source-bound.
- `bookmarks` (object[]) — Whole-block bookmark ranges. Source-free authoring requires one unique valid Word name around exactly one paragraph-like block; imported bookmarks are fixed-topology/read-only.
- `notes` (object[]) — Plain-text footnote/endnote records. The bounded profile permits one note at the end of each paragraph or list item; imported note text may change, but kind, anchor, native ID, and topology are source-bound.
- `bibliography` (object) — Canonical Word bibliography SelectedStyle, StyleName, and URI metadata authored in one b:Sources Custom XML part.
- `bibliographySources` (object[]) — Bounded Word bibliography sources with ordinary personal or corporate Author data and supported scalar fields. Imported source order, IDs, and tags remain source-bound.
- `headers` (object[]) — Header block models. Imported items expose sourceBound/editable; only a uniquely used one-run direct text paragraph is text-editable, and at most one edit may target each source Header part.
- `footers` (object[]) — Footer block models. Imported items expose sourceBound/editable; only a uniquely used one-run direct text paragraph is text-editable, and at most one edit may target each source Footer part.
- `sectionSettings` (object[]) — Per-section settings with zero-based sectionIndex and differentFirstPage activation state.
- `comments` (object[]) — Classic whole-paragraph comments. Parent/reply, resolved, durable-ID, UTC/person, and modern extension metadata are outside the OfficeKit 0.2 boundary.
- `settings` (object) — evenAndOddHeaders, mirrorMargins, gutterAtTop, trackRevisions, the updateFields-on-open refresh hint, and bounded passwordless documentProtection are authorable. mirrorMargins toggles facing-page inside/outside margins; gutterAtTop chooses whether each section's gutter is added at the top edge or binding side. Irregular page-margin mode markup stays source-owned and makes section geometry read-only. Password/cryptographic protection variants cannot be replaced through the semantic model.

**Schema returns:**

- `document` (DocumentModel) — Editable document facade.

#### `documentTable.setAccessibilityMetadata`

Set or clear non-visible Word table alternative text through w:tblCaption and w:tblDescription. It never creates a visible caption paragraph or changes layout; duplicate, empty, child-bearing, extension-bearing, or irregular imported leaves fail closed.

**Schema parameters:**

- `title` (string | null) — Non-visible table alternative-text title. A 1 through 32767 character XML-safe string writes canonical w:tblCaption/@w:val; null clears it.
- `description` (string | null) — Non-visible table alternative-text description. A 1 through 32767 character XML-safe string writes canonical w:tblDescription/@w:val; null clears it.

**Schema returns:**

- `table` (DocumentTableBlock) — Sets or clears non-visible table alternative text without creating a visible caption paragraph or changing layout. A source-bound table may change this metadata only when each native leaf is absent or appears exactly once as a direct canonical non-empty w:val-only leaf; duplicates, empty values, children, extensions, or other irregular profiles remain source-owned and fail closed.

#### `documentTable.setHeaderRowCount`

Set the number of contiguous leading rows marked with native w:tblHeader repetition semantics. This is separate from headerFill styling; imported tables accept it only when their row-property profile is canonical, otherwise the edit fails closed.

**Schema parameters:**

- `count` (number) required — Integer from 0 through the table's physical row count. It sets the complete native w:tblHeader prefix, not cell fill styling.

**Schema returns:**

- `table` (DocumentTableBlock) — Sets source-free or recognized imported table repeat-header rows. Imported row properties may contain only canonical grid offsets, non-clipping w:trHeight hRule=atLeast, and no-w:val w:cantSplit/w:tblHeader leaves in native order; non-prefix, duplicate, exact-height, explicit-value, extension-bearing, or otherwise irregular profiles stay source-bound and fail closed.

#### `documentTable.setHorizontalAlignment`

Set or clear native table-level w:jc placement. Center/right require zero table indent so OfficeKit never relies on host-specific resolution of competing w:jc and w:tblInd values; irregular imported table-property profiles fail closed.

**Schema parameters:**

- `value` ("left" | "center" | "right" | null) required — left, center, or right writes canonical table-level w:jc; null removes it. center/right require the table's indentDxa to be exactly 0.

**Schema returns:**

- `table` (DocumentTableBlock) — Sets or clears source-free or recognized imported table placement. Imported table properties must be the complete canonical fixed-layout direct-formatting profile with zero-or-one canonical w:jc leaf; duplicate, malformed, extension-bearing, or center/right-plus-indent profiles stay source-bound and fail closed.

#### `documentTable.setMinimumRowHeight`

Set or clear one physical row's non-clipping minimum height through native w:trHeight hRule=atLeast. It is not a fixed exact height or a pagination calculator; imported tables accept it only under the canonical row-property profile, otherwise the edit fails closed.

**Schema parameters:**

- `rowIndex` (number) required — Zero-based physical table row index from 0 through rowCount - 1.
- `heightDxa` (number|null) — Positive integer DXA minimum from 1 through 1000000, or null to clear the native height leaf. OfficeKit always writes hRule=atLeast, never an exact clipping height.

**Schema returns:**

- `table` (DocumentTableBlock) — Sets source-free or recognized imported row minimum height. Imported row properties may contain only canonical grid offsets, one positive w:trHeight hRule=atLeast, and no-w:val w:cantSplit/w:tblHeader leaves in native order; duplicate, exact, explicit-value, reordered, extension-bearing, or otherwise irregular profiles stay source-bound and fail closed.

#### `documentTable.setRowKeepTogether`

Set whether one physical table row may split across pages through native w:cantSplit. This is a per-row pagination constraint, not a row-group or pagination calculator; imported tables accept it only under the canonical row-property profile, otherwise the edit fails closed.

**Schema parameters:**

- `rowIndex` (number) required — Zero-based physical table row index from 0 through rowCount - 1.
- `keepTogether` (boolean) — True prevents the selected physical row from splitting across pages through native w:cantSplit; false removes only that native marker. Defaults to true.

**Schema returns:**

- `table` (DocumentTableBlock) — Sets source-free or recognized imported row page-break policy. Imported row properties may contain only canonical grid offsets, non-clipping w:trHeight hRule=atLeast, and no-w:val w:cantSplit/w:tblHeader leaves in native order; duplicate, exact-height, explicit-value, reordered, extension-bearing, or otherwise irregular profiles stay source-bound and fail closed.

#### `documentTableCell.addCheckboxContentControl`

Wrap one source-free rectangular table cell in a canonical Word 2010+ checkbox w:sdt. OfficeKit owns the visible glyph and symbols; recognized imports permit checked/tag/alias edits while identity, type, placement, symbols, and topology remain fixed.

**Schema parameters:**

- `checked` (boolean) — Initial checked state; defaults to false.
- `id` (string) — Agent-facing content-control ID; generated when omitted.
- `tag` (string) required — Table-cell checkbox SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Non-empty human title/alias, at most 255 characters; defaults to tag.

**Schema returns:**

- `control` (DocumentContentControlHandle) — Canonical Word 2010+ checkbox around the cell's existing single paragraph/run. Source-free rectangular cells may add it once. Visible glyph and symbols are codec-owned; recognized imports keep native ID, type, placement, row/column, symbol profile, and topology fixed.

#### `documentTableCell.addComboBoxContentControl`

Wrap one source-free rectangular table cell in a canonical standard combo-box w:sdt with ordered choices and a declared-or-custom typed value. Recognized imports permit value/tag/alias edits while the choice table and topology remain fixed.

**Schema parameters:**

- `choices` (Array<string|object>) required — Ordered 1 to 256 choice table. A string uses the same displayText and value; objects require unique XML-safe displayText and value strings of 1 to 255 characters.
- `value` (string) — Initial value, 1 to 255 XML-safe characters; defaults to the first choice. A matching choice uses its displayText, while custom text is shown verbatim.
- `id` (string) — Agent-facing content-control ID; generated when omitted.
- `tag` (string) required — Table-cell combo-box SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Non-empty human title/alias, at most 255 characters; defaults to tag.

**Schema returns:**

- `control` (DocumentContentControlHandle) — Canonical standard combo box around the cell's existing single paragraph/run. Source-free rectangular cells may add it once; recognized imports keep native ID, type, placement, row/column, ordered choice table, and topology fixed.

#### `documentTableCell.addDateContentControl`

Wrap one source-free rectangular table cell in the canonical ISO/Gregorian date w:sdt profile. Recognized imports permit dateValue/tag/alias edits while native date metadata, placement, and topology remain fixed.

**Schema parameters:**

- `dateValue` (string) required — Real proleptic Gregorian date in exact YYYY-MM-DD form, from 0001-01-01 through 9999-12-31. Date objects and locale-formatted strings are rejected.
- `id` (string) — Agent-facing content-control ID; generated when omitted.
- `tag` (string) required — Table-cell date SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Non-empty human title/alias, at most 255 characters; defaults to tag.

**Schema returns:**

- `control` (DocumentContentControlHandle) — Canonical ISO/Gregorian date picker around the cell's existing single paragraph/run. Source-free rectangular cells may add it once; recognized imports keep native ID, type, placement, row/column, native date profile, and topology fixed.

#### `documentTableCell.addDropdownContentControl`

Wrap one source-free rectangular table cell in a canonical standard drop-down w:sdt with ordered choices and a typed selectedValue. Recognized imports permit selectedValue/tag/alias edits while the choice table and topology remain fixed.

**Schema parameters:**

- `choices` (Array<string|object>) required — Ordered 1 to 256 choice table. A string uses the same displayText and value; objects require unique XML-safe displayText and value strings of 1 to 255 characters.
- `selectedValue` (string) — Initial internal choice value; defaults to the first choice.
- `id` (string) — Agent-facing content-control ID; generated when omitted.
- `tag` (string) required — Table-cell drop-down SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Non-empty human title/alias, at most 255 characters; defaults to tag.

**Schema returns:**

- `control` (DocumentContentControlHandle) — Canonical standard drop-down around the cell's existing single paragraph/run. Source-free rectangular cells may add it once; recognized imports keep native ID, type, placement, row/column, ordered choice table, and topology fixed.

#### `documentTableCell.addTextContentControl`

Wrap one source-free rectangular table cell's existing text in a canonical cell-level plain-text w:sdt. The handle reports placement=tableCell plus row/column; recognized imported controls permit fixed-topology text/tag/alias edits, while adding or removing imported control topology fails closed.

**Schema parameters:**

- `id` (string) — Agent-facing content-control ID; generated when omitted.
- `tag` (string) required — Table-cell plain-text SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Non-empty human title/alias, at most 255 characters; defaults to tag.

**Schema returns:**

- `control` (DocumentContentControlHandle) — Canonical plain-text control around the cell's existing single paragraph/run. Source-free rectangular cells may add it once; recognized imported controls keep native ID, type, placement, row/column, and topology fixed.

#### `documentTableCell.replaceText`

Apply a literal source-bound text patch to one table cell that advertises textPatchable. The search must resolve exactly once inside one ordinary native w:t node or adjacent non-empty direct runs with byte-identical w:rPr. Whole-cell replacement, mixed formatting, empty-run gaps, paragraph boundaries, fields, controls, revisions, and ambiguous matches fail closed.

**Schema parameters:**

- `search` (string) required — Non-empty literal text that must occur exactly once in the visible cell. A source-bound match may occupy one ordinary direct w:r/w:t or adjacent non-empty direct runs only when their exact w:rPr markup is identical and it never crosses a paragraph boundary.
- `replacement` (string) required — XML-safe replacement text, up to 1,000,000 characters.

**Schema returns:**

- `cell` (DocumentTableCell) — Mutated table-cell facade with one pending source-bound text patch.

#### `documentWatermark.remove`

Remove one modeled or recognized source-bound canonical watermark as a complete header paragraph. The source-bound operation re-proves exact element and header residual hashes and never heuristically deletes arbitrary header graphics.

**Schema returns:**

- `watermark` (undefined) — Removes the complete recognized watermark paragraph after source/residual revalidation on export.

#### `exportDocxWithOfficeKit`

Export bounded DocumentModel paragraphs/runs, fields, tables, bookmarks, notes, citations plus one canonical bibliography-output placeholder, tracked changes, comments, images, canonical text watermarks, sections, numbering, settings, and source-free ordered header/footer literal/simple-field sequences; recognized imports permit exact-profile semantic edits plus hash-bound literal patches to one unique ordinary paragraph or table-cell span inside one direct w:r/w:t or adjacent same-format runs while preserving all surrounding native markup.

**Schema parameters:**

- `document` (DocumentModel) required — Document facade within the OfficeKit paragraph/run/style, inline SEQ/REF/PAGEREF field, source-free switch-free BIBLIOGRAPHY output placeholder, section, header/footer, canonical text-watermark, image, list, hyperlink, whole-block bookmark, plain-text footnote/endnote, simple-field, comment, and fixed-table boundary. Advanced imported content remains source-bound; unsupported edits fail closed.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — DOCX bytes produced by the bundled Open XML SDK NativeAOT codec, with codec diagnostics in metadata.

#### `importDocxWithOfficeKit`

Import DOCX bytes through OfficeKit with source-bound blocks, recognized exclusive canonical VML text-watermark paragraphs, source-bound header/footer editable evidence, and read-only ordered header/footer literal/simple-field sequences. A header/footer edit is limited to one direct unformatted text paragraph in one uniquely used source part; fields, rich/shared/inherited page furniture, scope changes, and multiple edits to one part fail closed. Literal body/table patch capability never implies whole-paragraph/cell editability; only adjacent non-empty direct runs with byte-identical w:rPr may form one patch span, while mixed-format, gapped, cross-paragraph, ambiguous, field/control/revision text remains fail-closed.

**Schema parameters:**

- `input` (FileBlob|Uint8Array|ArrayBuffer) required — DOCX package bytes.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `document` (DocumentModel) — Imported document facade carrying source/opaque evidence. Canonical footnote/endnote bodies, exclusive VML text watermarks, and one direct unformatted header/footer paragraph in a uniquely used source part are text-editable with fixed source-bound anchors; whole-block bookmarks are fixed-topology/read-only, and other complex graphs remain source-bound.

#### `paragraph.addCheckboxContentControl`

Append one canonical Word 2010+ checkbox content control with typed checked state; OfficeKit owns its visible glyph and w14 symbol declarations.

**Schema parameters:**

- `checked` (boolean) — Initial checked state; defaults to false.
- `id` (string) — Agent-facing model ID; generated when omitted.
- `tag` (string) required — Checkbox SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Human title/alias, at most 255 characters; defaults to tag.
- `style` (object) — Optional modeled run formatting for the canonical visible glyph.

**Schema returns:**

- `run` (object) — Appended paragraph run carrying bounded canonical checkbox content-control metadata.

#### `paragraph.addComboBoxContentControl`

Append one canonical inline Word combo-box content control with ordered displayText/value choices and a typed value that may be a declared choice or bounded custom text. OfficeKit derives the visible projection.

**Schema parameters:**

- `choices` (Array<string|object>) required — Ordered 1 to 256 choice table. A string uses the same displayText and value; objects require unique XML-safe displayText and value strings of 1 to 255 characters.
- `value` (string) — Initial value, 1 to 255 XML-safe characters; defaults to the first choice. A matching choice uses its displayText, while custom text is shown verbatim.
- `id` (string) — Agent-facing model ID; generated when omitted.
- `tag` (string) required — Combo-box SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Human title/alias, at most 255 characters; defaults to tag.
- `style` (object) — Optional modeled run formatting for the derived visible value.

**Schema returns:**

- `run` (object) — Appended paragraph run carrying bounded canonical combo-box content-control metadata.

#### `paragraph.addDateContentControl`

Append one canonical inline Word date picker from a real Gregorian YYYY-MM-DD value. OfficeKit owns the fixed ISO display, UTC-midnight fullDate, language, mapping, and calendar projection.

**Schema parameters:**

- `dateValue` (string) required — Real proleptic Gregorian date in exact YYYY-MM-DD form, from 0001-01-01 through 9999-12-31. Date objects and locale-formatted strings are rejected.
- `id` (string) — Agent-facing model ID; generated when omitted.
- `tag` (string) required — Date SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Human title/alias, at most 255 characters; defaults to tag.
- `style` (object) — Optional modeled run formatting for the codec-owned ISO visible date.

**Schema returns:**

- `run` (object) — Appended paragraph run carrying bounded canonical date content-control metadata.

#### `paragraph.addDropdownContentControl`

Append one canonical inline Word drop-down content control with an ordered displayText/value choice table and typed selectedValue. OfficeKit derives visible text from the selected choice.

**Schema parameters:**

- `choices` (Array<string|object>) required — Ordered 1 to 256 choice table. A string uses the same displayText and value; objects require unique XML-safe displayText and value strings of 1 to 255 characters.
- `selectedValue` (string) — Initial internal choice value; defaults to the first choice.
- `id` (string) — Agent-facing model ID; generated when omitted.
- `tag` (string) required — Drop-down SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Human title/alias, at most 255 characters; defaults to tag.
- `style` (object) — Optional modeled run formatting for the derived visible choice text.

**Schema returns:**

- `run` (object) — Appended paragraph run carrying bounded canonical drop-down content-control metadata.

#### `paragraph.addField`

Append a logical inline SEQ, REF, or PAGEREF field run. A SEQ run may add a bookmark around only its cached result for real caption-number targets. OfficeKit authors/imports the canonical native graph; imported field position, instruction, and bookmark identity remain source-bound while cached display text is editable.

**Schema parameters:**

- `instruction` (string) required — Canonical SEQ <label> \* ARABIC, REF <bookmark> \h, or PAGEREF <bookmark> \h instruction using a bounded Word-compatible name.
- `display` (string) — Cached visible result before host refresh; defaults to 0.
- `bookmarkName` (string) — Optional unique Word bookmark name for a SEQ field; wraps only the cached-result run so REF/PAGEREF can target the caption number.
- `bookmarkNativeId` (number) — Optional unsigned 32-bit native bookmark ID for source-free authoring; imported identity is source-bound.
- `style` (object) — Optional modeled formatting for the cached result run.

**Schema returns:**

- `run` (object) — Logical inline field run. Imported position, instruction, and optional bookmark identity are source-bound; cached display text remains editable.

#### `paragraph.addTextContentControl`

Append one inline plain-text Word content-control run with agent ID, tag, alias, text, and optional run formatting. OfficeKit assigns native w:id identity and authors canonical w:sdt markup.

**Schema parameters:**

- `text` (string) required — Initial visible control text.
- `id` (string) — Agent-facing model ID; generated when omitted.
- `tag` (string) required — Plain-text SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Human title/alias, at most 255 characters; defaults to tag.
- `style` (object) — Optional modeled run formatting.

**Schema returns:**

- `run` (object) — Appended paragraph run carrying bounded inline plain-text content-control metadata.

#### `paragraph.replaceText`

Replace literal paragraph text without flattening formatting boundaries. Fully editable one-run paragraphs update their existing run; imported source-bound paragraphs advertise textPatchable when OfficeKit can replace one unique ordinary w:r/w:t node or adjacent non-empty direct runs with byte-identical w:rPr while preserving all native topology and surrounding markup. Mixed formatting, empty-run gaps, paragraph boundaries, fields, controls, revisions, and duplicate matches fail closed.

**Schema parameters:**

- `search` (string) required — Non-empty literal text that must occur exactly once. A source-bound match may occupy one ordinary direct w:r/w:t or adjacent non-empty direct runs only when their exact w:rPr markup is identical.
- `replacement` (string) required — XML-safe replacement text, up to 1,000,000 characters.

**Schema returns:**

- `paragraph` (DocumentParagraphBlock) — Mutated paragraph facade. Source-bound patches are applied only after native-node and source-hash validation during export.

## pdf

| Name | Kind | Summary |
| --- | --- | --- |
| `createPdfjsParser` | api | Create an optional PDF.js parser adapter to extract page geometry, positioned text, heuristic tables, and bounded embedded raster or stencil-mask PNG images with placement boxes. |
| `pdf.addChart` | api | Add a modeled bar/line chart region with categories, series, title, meaningful alternative text or decorative-artifact semantics, bbox, inspect/resolve/layout records, SVG preview, and PDF metadata roundtrip. |
| `pdf.addFlowText` | api | Wrap long text into positioned lines and automatically append pages when the configured content box is full. |
| `pdf.addImage` | api | Add a modeled PDF image region with dataUrl/URI/prompt metadata, meaningful alternative text or explicit decorative-artifact semantics, and a page-space bounding box. |
| `pdf.addLink` | api | Add a meaningful visible http, https, or mailto link with stable ID, page-space bounding box, tagged Link structure, URI annotation, OBJR association, and explicit reading-order participation. |
| `pdf.addPage` | api | Append a modeled PDF page with explicit point dimensions and optional text, positioned items, regions, tables, images, charts, and links. |
| `pdf.addTable` | api | Add a modeled table with cell values, row/column spans, TH/TD roles, scopes, header associations, stable cell IDs, a page-space bounding box, and optional semanticId joining constrained consecutive-page segments into one logical Table. |
| `pdf.addText` | api | Add positioned PDF text with page-space bbox, font metadata, optional semantic H1-H6 heading level or decorative Artifact semantics, inspect/resolve/layout records, and SVG preview rendering. |
| `pdf.extractTables` | api | Extract modeled table values, normalized spanning-cell/header records, and bounding boxes across all pages or a selected page. |
| `pdf.extractText` | api | Extract modeled text across all pages or a selected page. |
| `pdf.inspect` | api | Emit bounded NDJSON for pages, text, positioned text items, reading-order entries, layout regions, tables/table cells, images, charts, and links; narrow with search/target anchors and shape fields with include/exclude. |
| `pdf.layoutJson` | api | Return modeled PDF page layout JSON with page text, positioned text items, explicit/effective reading order, layout regions, normalized table cells/spans/header IDs, images, charts, links, and target/search context slicing. |
| `pdf.page.setReadingOrder` | api | Declare the complete logical reading sequence of a page's body text, positioned text, tables, images, charts, and links by stable ID without changing visual paint order. |
| `pdf.render` | api | Render a modeled PDF page to SVG by default, return page layout JSON with { format: 'layout' }, or use { source: 'pdf', renderer } to feed the exported PDF into Poppler/PDF-capable raster adapters. |
| `pdf.resolve` | api | Resolve stable PDF artifact IDs for pages, page text blocks, positioned text items, reading-order entries, layout regions, tables/table cells, images, charts, and links. |
| `pdf.verify` | api | Return QA issues for invalid H1-H6 nesting, missing/generic Figure alternative text, meaningless/unsafe links, cross-page logical-table continuity, incomplete/duplicate/unknown reading-order targets, empty pages, text extraction sanity, geometry/bounds, invalid images, table semantics, and chart data. |
| `PdfArtifact.create` | api | Create a modeled PDF artifact with pages, text, span-aware accessible table regions, image regions, charts, links, and explicit reading order. |
| `PdfFile.editPdf` | api | Apply bounded direct-original MuPDF.js operations with explicit rewrite or byte-prefix-verified incremental save, object-level signature detection, atomic caller-controlled output, and fail-closed rejection of incremental page-tree/redaction/deletion/source-bound annotation or link operations, signed incremental edits, ambiguous radio export values, rotated-page crop requests, unsafe link destinations, clipped native appearances, and unsupported operations. add_text_annotation, add_free_text_annotation, add_area_annotation, add_text_markup, compatibility add_text_highlight, and add_link bind one exact source hash plus the inspected mupdfPage bbox/rotation snapshot and accept coordinates in its explicit mupdf-page-space after the current 0/90/180/270-degree rotation; raw MediaBox/CropBox remain unrotated PDF-space facts. Text-note, FreeText, area, and text-markup records expose provider appearanceBbox evidence, and placement fails before publication if the full native appearance would leave the visible page. add_free_text_annotation accepts one visible bbox and bounded contents with a fixed Helvetica appearance, 4–72 point font size, RGB text color, and alignment; it rejects styling it cannot represent and verifies the native appearance retains all requested text before save. add_area_annotation accepts one visible bbox plus rectangle/ellipse intent, a solid 0.5–12 point RGB outline, and optional review metadata; it creates no interior fill and rejects any requested or native appearance that leaves the page. add_text_markup accepts exactly Highlight, Underline, StrikeOut, or Squiggly over one unique native text-search selection plus optional RGB/review metadata, never caller quads or rectangles; add_text_highlight remains the compatible Highlight-only form. delete_page, duplicate_page, and rearrange_pages are source-bound single-operation full rewrites for untagged PDFs: deletion and duplication bind the selected page snapshot, while rearrangement binds a complete current-order page snapshot before applying one changed complete permutation; every output requires fresh inspection and mapped rendering. duplicate_page additionally rejects unsupported page-bound graphs or projected page/object budget overflow and does not synthesize navigation. delete_embedded_file binds the exact source SHA-256 plus one inspect-returned canonical catalog EmbeddedFiles locator and complete snapshot, verifies all non-target entries, and requires rewrite; it removes a catalog NameTree entry but never claims sanitize or physical payload erasure. set_metadata binds one exact source plus the complete mupdfDocumentMetadata snapshot, preserves non-target Info entries, and either updates Document Info alone or synchronizes requested existing properties in a field-safe XMP packet while proving all other decoded packet bytes unchanged. Inspection exposes xmpMutableFields and field-specific xmpBlockedFields: a unique x-default may be edited in a multilingual Alt, common scalar description attributes are supported, and a multi-author or irregular field blocks only itself. Missing or blocked requested properties, CDATA/DTD/invalid entities, malformed XML, stale/no-op/partial evidence, and legacy unbound payloads fail closed. update_outline binds the exact source hash, one inspect-derived mupdfOutline locator, and its complete path/title/URI/open/page/child-count snapshot; it changes only the title and, for a parent, expansion state while preserving destinations, order, nesting, and every non-target outline. Leaf expansion edits, topology or destination edits, stale/partial evidence, controls in titles, and no-op patches fail closed. delete_annotation, update_annotation, delete_link, and update_link require an inspect-returned source hash, source-bound locator, and snapshot precondition. update_annotation retains compatible non-empty contents/author/subject patches for native Text notes. A complete fixed-helvetica-v1 FreeText snapshot permits only contents/author/subject while preserving geometry/style and re-verifying all visible appearance text. A complete solid-no-fill-v1 Square/Circle snapshot permits contents/author/subject/RGB color while preserving geometry, line width/style, no-fill state, flags, page, locator, and on-page appearance. Native Highlight, Underline, StrikeOut, and Squiggly require their complete inspect-returned snapshot and may additionally change RGB color while preserving type, quadrilaterals, rectangle, appearance bounds, flags, page, and locator. Partial/stale snapshots, no-op or geometry/style patches, unsupported annotation profiles, and incremental save fail closed. update_link changes only a safe non-empty URL; native link geometry is never patched. set_page_crop changes only the visible CropBox on an unrotated page, while rotate_page writes an absolute right-angle /Rotate value; neither removes hidden original content. |
| `PdfFile.editPdf.add_area_annotation` | api | Add one source- and page-snapshot-bound unfilled native Square or Circle review outline to imported PDF bytes. The rewrite-only operation accepts rectangle/ellipse intent, explicit mupdf-page-space geometry, bounded solid RGB stroke width, and optional review metadata; stale evidence, painted bounds outside the page, fill/dash/cloud/opacity/arbitrary appearance, incremental save, and silent fallback fail closed. A fresh solid-no-fill-v1 snapshot may revise contents/author/subject/RGB color through update_annotation while geometry and border style remain fixed. |
| `PdfFile.editPdf.add_free_text_annotation` | api | Add one source- and page-snapshot-bound native FreeText review box to imported PDF bytes. The rewrite-only operation uses explicit mupdf-page-space geometry, a fixed Helvetica appearance, bounded font size/RGB text color/alignment, and native appearance-text verification; clipped text, off-page appearance, unsupported styling, stale evidence, incremental save, and silent fallback fail closed. A freshly inspected fixed-helvetica-v1 result may update only contents/author/subject through update_annotation with its complete snapshot and renewed native appearance-text verification. |
| `PdfFile.exportPdf` | api | Export a modeled artifact as a real multi-page tagged PDF 1.7 whose logical structure follows explicit page reading order without changing paint order, emits semantic H1-H6 headings, meaningful Figure /Alt text, Link annotations with OBJR associations, /Artifact marked content, and constrained logical Tables spanning consecutive pages, and preserves language/title, Table/TR/TH/TD hierarchy, optional Unicode TrueType embedding, positioned text, vector charts, and PNG/JPEG images. |
| `PdfFile.importPdf` | api | Reopen package-generated metadata losslessly or lazily use required MuPDF.js for arbitrary PDFs, producing a bounded reconstructed extraction/QA view with text geometry, raster placements and transforms, links, annotations, widgets, and heuristic table candidates; the view is never an edit representation. |
| `PdfFile.inspectPdf` | api | Inspect a path or PDF bytes after a pre-WASM input budget, combining native MuPDF page/object/annotation/widget/link/outline records, source SHA-256, source-bound Document Info records plus annotation/link/outline/canonical catalog EmbeddedFiles locators, complete XMP stream identity/profile/mutable-and-blocked-field evidence, native annotation snapshots and update capabilities plus FreeText default-appearance/alignment, area border/style, and text-markup quadrilateral/color facts, raw MediaBox/CropBox facts, and effective normalized page rotation with bounded tagged-PDF, language, reading-order, heading, Figure, Link, Artifact, font, and table-structure evidence. |
| `PdfFile.renderPdf` | api | Render one page from original PDF bytes through runtime-lazy MuPDF.js as PNG or JPEG, enforcing input, page/object, DPI, and preallocation pixel budgets before returning a FileBlob. |
| `PdfProviders.ensure` | api | Install only a previously installable, policy- and catalog-bound capability resolution into the project-private cache, then return a fresh probe. It uses only catalog-pinned release bytes, validates size/hash/archive/receipt boundaries, and never chooses another provider or obtains credentials. A blocked or ready resolution cannot be forced through this API. |
| `PdfProviders.probe` | api | Probe exactly one selected PDF provider under the requested policy without downloading, mutating the cache, importing MuPDF, or trying fallback providers. The result reports ready or blocked runtime evidence together with the pinned pack plan. |
| `PdfProviders.resolve` | api | Resolve one explicit PDF task and selected/default provider against the immutable capability catalog and project policy. It is read-only: no MuPDF initialization, network access, cache mutation, credential acquisition, or automatic provider fallback occurs. The result is ready, installable, or blocked with exact packs, platform, sizes, licenses, runtime prerequisites, consents, and operation limits. |

### pdf details

#### `createPdfjsParser`

Create an optional PDF.js parser adapter to extract page geometry, positioned text, heuristic tables, and bounded embedded raster or stencil-mask PNG images with placement boxes.

**Examples:**

- const parser = createPdfjsParser({ getDocumentOptions: { useSystemFonts: true } })

**Schema parameters:**

- `pdfjs` (object) — Injected PDF.js module; otherwise pdfjs-dist is loaded.
- `getDocumentOptions` (object) — Options merged into PDF.js getDocument().
- `textContentOptions` (object) — Options merged into getTextContent().

**Schema returns:**

- `parser` (function) — Parser adapter for PdfFile.importPdf().

#### `pdf.addChart`

Add a modeled bar/line chart region with categories, series, title, meaningful alternative text or decorative-artifact semantics, bbox, inspect/resolve/layout records, SVG preview, and PDF metadata roundtrip.

**Examples:**

- pdf.addChart({ pageIndex: 0, chartType: 'bar', categories: ['A', 'B'], series: [{ name: 'Score', values: [2, 4] }], bbox: [72, 430, 468, 180] })

**Schema parameters:**

- `pageIndex` (number) — Zero-based target page index.
- `chartType` (string) — bar or line.
- `title` (string) — Visible chart title.
- `alt` (string) — Meaningful alternative text describing the chart; required unless decorative is true.
- `decorative` (boolean) — Mark the chart as decorative PDF Artifact content and exclude it from logical reading order.
- `categories` (string[]) required — Category labels.
- `series` (object[]) required — Series with name, numeric values, and optional color.
- `bbox` (number[]) — Page-space [left, top, width, height] in points.

**Schema returns:**

- `chart` (PdfChart) — Inspectable chart facade with stable ID.

#### `pdf.addFlowText`

Wrap long text into positioned lines and automatically append pages when the configured content box is full.

**Examples:**

- pdf.addFlowText(longReport, { fontSize: 11, margins: { top: 72, right: 72, bottom: 72, left: 72 } })

**Schema parameters:**

- `text` (string) required — Paragraph text separated by newlines.
- `pageIndex` (number) — Zero-based starting page index; defaults to the first page.
- `margins` (number|object) — Uniform margin or top/right/bottom/left page margins in points.
- `left` (number) — Explicit content-box left edge overriding margins.left.
- `top` (number) — Explicit first-page top edge overriding margins.top.
- `width` (number) — Explicit content width; defaults to page width minus horizontal margins.
- `fontSize` (number) — Line font size in points.
- `lineHeight` (number) — Line advance in points.
- `paragraphGap` (number) — Extra vertical space after each paragraph.

**Schema returns:**

- `flow` (object) — Flow ID, positioned items, page IDs, page indexes, and line count.

#### `pdf.addImage`

Add a modeled PDF image region with dataUrl/URI/prompt metadata, meaningful alternative text or explicit decorative-artifact semantics, and a page-space bounding box.

**Examples:**

- pdf.addImage({ pageIndex: 0, dataUrl, alt: 'Approval mark', bbox: [430, 60, 64, 64] })

**Schema parameters:**

- `pageIndex` (number) — Zero-based target page index.
- `dataUrl` (string) — Embedded PNG or JPEG image data URL.
- `uri` (string) — External image URI metadata.
- `prompt` (string) — Image generation/extraction prompt metadata.
- `alt` (string) — Meaningful alternative text; required unless decorative is true.
- `decorative` (boolean) — Mark the image as decorative PDF Artifact content and exclude it from logical reading order.
- `bbox` (number[]) — Page-space [left, top, width, height] in points.
- `fit` (string) — contain or cover intent metadata.

**Schema returns:**

- `image` (PdfImage) — Inspectable image facade with stable ID.

#### `pdf.addLink`

Add a meaningful visible http, https, or mailto link with stable ID, page-space bounding box, tagged Link structure, URI annotation, OBJR association, and explicit reading-order participation.

**Examples:**

- pdf.addLink({ pageIndex: 0, text: 'W3C accessibility guidance', url: 'https://www.w3.org/WAI/', bbox: [72, 700, 240, 18] })

**Schema parameters:**

- `pageIndex` (number) — Zero-based target page index.
- `text` (string) required — Visible text that meaningfully describes the destination; generic text and raw URLs fail verification.
- `url` (string) required — Absolute http, https, or mailto destination.
- `bbox` (number[]) — Page-space [left, top, width, height] in points for visible text and the URI annotation.

**Schema returns:**

- `link` (PdfLink) — Inspectable link facade with stable ID.

#### `pdf.addPage`

Append a modeled PDF page with explicit point dimensions and optional text, positioned items, regions, tables, images, charts, and links.

**Examples:**

- pdf.addPage({ width: 612, height: 792, text: 'Appendix' })

**Schema parameters:**

- `width` (number) — Page width in points; defaults to 612.
- `height` (number) — Page height in points; defaults to 792.
- `text` (string) — Extractable page text.
- `textItems` (object[]) — Positioned text item models.
- `regions` (object[]) — Inspectable page-space regions.
- `tables` (object[]) — Modeled page tables.
- `images` (object[]) — Modeled page images.
- `charts` (object[]) — Modeled page charts.
- `links` (object[]) — Modeled visible URI links.
- `readingOrder` (string[]|object[]) — Optional complete logical order of all semantic page items as stable IDs or objects with IDs.

**Schema returns:**

- `page` (PdfPage) — Appended editable page facade.

#### `pdf.addTable`

Add a modeled table with cell values, row/column spans, TH/TD roles, scopes, header associations, stable cell IDs, a page-space bounding box, and optional semanticId joining constrained consecutive-page segments into one logical Table.

**Examples:**

- pdf.addTable({ name: 'gates', values: [['Evidence', '', 'Status'], ['Model', 'Native', ''], ['PDF.js', 'Poppler', 'pass']], cells: [{ row: 0, column: 0, columnSpan: 2 }, { row: 0, column: 2, rowSpan: 2 }], bbox: [72, 140, 468, 96] })

**Schema parameters:**

- `name` (string) — Inspectable table name.
- `values` (unknown[][]) required — Rectangular or ragged cell value matrix.
- `semanticId` (string) — Optional logical table identity shared by constrained segments on consecutive pages. A continuation must be first and a non-final segment last in page reading order.
- `cells` (object[]) — Optional zero-based cell overrides with id, row, column, value, rowSpan, columnSpan, TH/TD role, Row/Column/Both scope, and header ID array.
- `bbox` (number[]) — Page-space [left, top, width, height] in points.
- `source` (string) — Optional extraction/source provenance.

**Schema returns:**

- `table` (PdfTable) — Inspectable table facade with stable cell IDs and getCell(row, column).

#### `pdf.addText`

Add positioned PDF text with page-space bbox, font metadata, optional semantic H1-H6 heading level or decorative Artifact semantics, inspect/resolve/layout records, and SVG preview rendering.

**Examples:**

- pdf.addText({ pageIndex: 0, text: 'Status', bbox: [72, 72, 200, 24], fontSize: 18, bold: true })

**Schema parameters:**

- `text` (string) required — Text content.
- `pageIndex` (number) — Zero-based target page index.
- `bbox` (number[]) — Page-space [left, top, width, height] in points.
- `fontName` (string) — Font family metadata.
- `fontSize` (number) — Font size in points.
- `color` (string) — Text color.
- `bold` (boolean) — Bold text flag.
- `italic` (boolean) — Italic text flag.
- `headingLevel` (number) — Optional semantic PDF heading level from 1 through 6; visual styling remains independent.
- `artifact` (boolean) — Mark repeating/decorative text such as running headers and footers as PDF Artifact content and exclude it from reading order. Cannot be combined with headingLevel.

**Schema returns:**

- `textItem` (object) — Positioned text item with stable ID.

#### `pdf.extractTables`

Extract modeled table values, normalized spanning-cell/header records, and bounding boxes across all pages or a selected page.

**Examples:**

- pdf.extractTables({ page: 1 })

**Schema parameters:**

- `page` (number) — Optional one-based page number.

**Schema returns:**

- `tables` (object[]) — Table records with page, ID, name, values, normalized cells, and bbox.

#### `pdf.extractText`

Extract modeled text across all pages or a selected page.

**Examples:**

- pdf.extractText({ page: 2 })

**Schema parameters:**

- `page` (number) — Optional one-based page number.

**Schema returns:**

- `text` (string) — Selected page text or all page text joined with blank lines.

#### `pdf.inspect`

Emit bounded NDJSON for pages, text, positioned text items, reading-order entries, layout regions, tables/table cells, images, charts, and links; narrow with search/target anchors and shape fields with include/exclude.

**Schema parameters:**

- `kind` (string) — Comma-separated page, text, textItem, readingOrder, region, table, tableCell, image, chart, and link record kinds.
- `search` (string) — Case-insensitive record filter.
- `target` (string) — Stable ID/anchor target; targetId, id, and anchor are aliases.
- `before` (number) — Records of context before target matches.
- `after` (number) — Records of context after target matches.
- `include` (string) — Comma-separated fields to keep.
- `exclude` (string) — Comma-separated fields to omit.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `inspection` (object) — Bounded { ndjson, truncated } inspection result.

#### `pdf.layoutJson`

Return modeled PDF page layout JSON with page text, positioned text items, explicit/effective reading order, layout regions, normalized table cells/spans/header IDs, images, charts, links, and target/search context slicing.

**Examples:**

- pdf.layoutJson({ page: 1, target: table.id, context: 1 })

**Schema parameters:**

- `page` (number) — Optional one-based page selector.
- `pageIndex` (number) — Optional zero-based page selector.
- `target` (string) — Stable target ID/anchor.
- `search` (string) — Case-insensitive layout-record filter.
- `before` (number) — Context records before matches.
- `after` (number) — Context records after matches.

**Schema returns:**

- `layout` (object) — Point-based PDF page layout tree and optional slice metadata.

#### `pdf.page.setReadingOrder`

Declare the complete logical reading sequence of a page's body text, positioned text, tables, images, charts, and links by stable ID without changing visual paint order.

**Examples:**

- page.setReadingOrder([`${page.id}/text`, image.id, heading.id, table.id, chart.id])

**Schema parameters:**

- `order` (string[]|object[]) required — Complete page sequence containing each semantic body-text, positioned-text, table, image, chart, and link target exactly once; artifact text and decorative figures are excluded.

**Schema returns:**

- `page` (PdfPage) — The same editable page facade for chaining.

#### `pdf.render`

Render a modeled PDF page to SVG by default, return page layout JSON with { format: 'layout' }, or use { source: 'pdf', renderer } to feed the exported PDF into Poppler/PDF-capable raster adapters.

**Examples:**

- await pdf.render({ pageIndex: 0 })
- await pdf.render({ source: 'pdf', format: 'png', renderer: createPopplerRenderer() })

**Schema parameters:**

- `pageIndex` (number) — Zero-based page index for modeled SVG rendering.
- `page` (number) — One-based page selector used by layout/native renderer workflows.
- `format` (string) — svg by default, layout, pdf, png, ppm, or tiff with a renderer.
- `source` (string) — Set to pdf to render exported PDF bytes.
- `renderer` (function) — Optional PDF-capable renderer adapter.

**Schema returns:**

- `blob` (FileBlob) — SVG, layout JSON, PDF, or renderer output.

#### `pdf.resolve`

Resolve stable PDF artifact IDs for pages, page text blocks, positioned text items, reading-order entries, layout regions, tables/table cells, images, charts, and links.

**Examples:**

- pdf.resolve('pg-1/txt/1')

**Schema parameters:**

- `id` (string) required — Stable artifact, page, text, text-item, reading-order, region, table, table-cell, image, chart, or link ID.

**Schema returns:**

- `object` (object|undefined) — Resolved editable facade/record or undefined.

#### `pdf.verify`

Return QA issues for invalid H1-H6 nesting, missing/generic Figure alternative text, meaningless/unsafe links, cross-page logical-table continuity, incomplete/duplicate/unknown reading-order targets, empty pages, text extraction sanity, geometry/bounds, invalid images, table semantics, and chart data.

**Examples:**

- pdf.verify({ maxChars: 12000 })

**Schema parameters:**

- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — PDF semantic QA result with ok, issues, ndjson, and truncated.

#### `PdfArtifact.create`

Create a modeled PDF artifact with pages, text, span-aware accessible table regions, image regions, charts, links, and explicit reading order.

**Examples:**

- const pdf = PdfArtifact.create({ pages: [{ width: 612, height: 792, text: 'Report' }] })

**Schema parameters:**

- `id` (string) — Optional stable artifact ID.
- `metadata` (object) — Clean-room metadata preserved through generated-PDF roundtrip.
- `text` (string) — Convenience text for a single default page.
- `pages` (object[]) — Page models with width, height, text, textItems, regions, tables, images, charts, links, and optional complete readingOrder ID arrays.

**Schema returns:**

- `pdf` (PdfArtifact) — Editable modeled PDF artifact.

#### `PdfFile.editPdf`

Apply bounded direct-original MuPDF.js operations with explicit rewrite or byte-prefix-verified incremental save, object-level signature detection, atomic caller-controlled output, and fail-closed rejection of incremental page-tree/redaction/deletion/source-bound annotation or link operations, signed incremental edits, ambiguous radio export values, rotated-page crop requests, unsafe link destinations, clipped native appearances, and unsupported operations. add_text_annotation, add_free_text_annotation, add_area_annotation, add_text_markup, compatibility add_text_highlight, and add_link bind one exact source hash plus the inspected mupdfPage bbox/rotation snapshot and accept coordinates in its explicit mupdf-page-space after the current 0/90/180/270-degree rotation; raw MediaBox/CropBox remain unrotated PDF-space facts. Text-note, FreeText, area, and text-markup records expose provider appearanceBbox evidence, and placement fails before publication if the full native appearance would leave the visible page. add_free_text_annotation accepts one visible bbox and bounded contents with a fixed Helvetica appearance, 4–72 point font size, RGB text color, and alignment; it rejects styling it cannot represent and verifies the native appearance retains all requested text before save. add_area_annotation accepts one visible bbox plus rectangle/ellipse intent, a solid 0.5–12 point RGB outline, and optional review metadata; it creates no interior fill and rejects any requested or native appearance that leaves the page. add_text_markup accepts exactly Highlight, Underline, StrikeOut, or Squiggly over one unique native text-search selection plus optional RGB/review metadata, never caller quads or rectangles; add_text_highlight remains the compatible Highlight-only form. delete_page, duplicate_page, and rearrange_pages are source-bound single-operation full rewrites for untagged PDFs: deletion and duplication bind the selected page snapshot, while rearrangement binds a complete current-order page snapshot before applying one changed complete permutation; every output requires fresh inspection and mapped rendering. duplicate_page additionally rejects unsupported page-bound graphs or projected page/object budget overflow and does not synthesize navigation. delete_embedded_file binds the exact source SHA-256 plus one inspect-returned canonical catalog EmbeddedFiles locator and complete snapshot, verifies all non-target entries, and requires rewrite; it removes a catalog NameTree entry but never claims sanitize or physical payload erasure. set_metadata binds one exact source plus the complete mupdfDocumentMetadata snapshot, preserves non-target Info entries, and either updates Document Info alone or synchronizes requested existing properties in a field-safe XMP packet while proving all other decoded packet bytes unchanged. Inspection exposes xmpMutableFields and field-specific xmpBlockedFields: a unique x-default may be edited in a multilingual Alt, common scalar description attributes are supported, and a multi-author or irregular field blocks only itself. Missing or blocked requested properties, CDATA/DTD/invalid entities, malformed XML, stale/no-op/partial evidence, and legacy unbound payloads fail closed. update_outline binds the exact source hash, one inspect-derived mupdfOutline locator, and its complete path/title/URI/open/page/child-count snapshot; it changes only the title and, for a parent, expansion state while preserving destinations, order, nesting, and every non-target outline. Leaf expansion edits, topology or destination edits, stale/partial evidence, controls in titles, and no-op patches fail closed. delete_annotation, update_annotation, delete_link, and update_link require an inspect-returned source hash, source-bound locator, and snapshot precondition. update_annotation retains compatible non-empty contents/author/subject patches for native Text notes. A complete fixed-helvetica-v1 FreeText snapshot permits only contents/author/subject while preserving geometry/style and re-verifying all visible appearance text. A complete solid-no-fill-v1 Square/Circle snapshot permits contents/author/subject/RGB color while preserving geometry, line width/style, no-fill state, flags, page, locator, and on-page appearance. Native Highlight, Underline, StrikeOut, and Squiggly require their complete inspect-returned snapshot and may additionally change RGB color while preserving type, quadrilaterals, rectangle, appearance bounds, flags, page, and locator. Partial/stale snapshots, no-op or geometry/style patches, unsupported annotation profiles, and incremental save fail closed. update_link changes only a safe non-empty URL; native link geometry is never patched. set_page_crop changes only the visible CropBox on an unrotated page, while rotate_page writes an absolute right-angle /Rotate value; neither removes hidden original content.

**Examples:**

- const inspection = await PdfFile.inspectPdf(pdf); const page = inspection.records.find((record) => record.kind === 'mupdfPage' && record.page === 1); await PdfFile.editPdf(pdf, { savePolicy: 'rewrite', operations: [{ type: 'add_text_annotation', page: 1, sourceSha256: inspection.summary.sourceSha256, expectedPage: { bbox: page.bbox, rotation: page.rotation }, point: [72, 72], contents: 'Review' }] })
- const inspection = await PdfFile.inspectPdf(pdf); const page = inspection.records.find((record) => record.kind === 'mupdfPage' && record.page === 1); await PdfFile.editPdf(pdf, { savePolicy: 'rewrite', operations: [{ type: 'add_text_markup', markup: 'underline', page: 1, sourceSha256: inspection.summary.sourceSha256, expectedPage: { bbox: page.bbox, rotation: page.rotation }, text: 'Review target', color: [0.1, 0.3, 0.9] }] })
- const inspection = await PdfFile.inspectPdf(pdf); const metadata = inspection.records.find((record) => record.kind === 'mupdfDocumentMetadata' && record.updateCapability.supported); await PdfFile.editPdf(pdf, { savePolicy: 'incremental', operations: [{ type: 'set_metadata', sourceSha256: inspection.summary.sourceSha256, metadataId: metadata.id, expected: metadata.snapshot, patch: { title: 'Reviewed title' } }] })
- const inspection = await PdfFile.inspectPdf(pdf); const outline = inspection.records.find((record) => record.kind === 'mupdfOutline' && record.updateCapability.mutableFields.includes('title')); await PdfFile.editPdf(pdf, { savePolicy: 'incremental', operations: [{ type: 'update_outline', sourceSha256: inspection.summary.sourceSha256, outlineId: outline.id, expected: outline.snapshot, patch: { title: 'Reviewed section' } }] })
- const inspection = await PdfFile.inspectPdf(pdf); const pages = inspection.records.filter((record) => record.kind === 'mupdfPage'); await PdfFile.editPdf(pdf, { savePolicy: 'rewrite', operations: [{ type: 'rearrange_pages', pages: [3, 1, 2], sourceSha256: inspection.summary.sourceSha256, expectedPages: pages.map(({ page, bbox, rotation }) => ({ page, bbox, rotation })) }] })
- const inspection = await PdfFile.inspectPdf(pdf); const file = inspection.records.find((record) => record.kind === 'mupdfEmbeddedFile' && record.name === 'review'); await PdfFile.editPdf(pdf, { savePolicy: 'rewrite', operations: [{ type: 'delete_embedded_file', sourceSha256: inspection.summary.sourceSha256, embeddedFileId: file.id, expected: file.snapshot }] })

**Schema parameters:**

- `pdf` (string|FileBlob|Uint8Array|ArrayBuffer) required — Original PDF path or bytes.
- `operations` (object[]) required — Typed MuPDF operations: source-bound add_text_annotation/add_free_text_annotation/add_area_annotation/add_text_markup and compatibility add_text_highlight, legacy fill_form, source-bound update_form_field, source-bound delete_page/duplicate_page/rearrange_pages/delete_embedded_file/set_metadata/update_outline, delete_annotation, update_annotation, set_page_crop, rotate_page, add_link, delete_link, update_link, redact_text, or redact_rect. Native placement operations require the exact inspect-returned sourceSha256 plus expectedPage={bbox,rotation}. The inspected bbox has coordinateSpace=mupdf-page-space: upper-left origin, y downward, with the current 0/90/180/270-degree page rotation already applied; raw MediaBox/CropBox remain unrotated PDF-space facts. add_text_annotation accepts a [x,y] pin and non-empty contents with optional non-empty author/subject, verifies exactly one native Text annotation, and returns its provider-normalized rect plus a conservative appearanceBbox. A requested bbox, text alias, icon selection, stale evidence, an appearance that could clip outside the visible box, or incremental save is rejected. add_free_text_annotation accepts a visible [x,y,width,height] bbox, non-empty <=4096-character contents, 4–72 point fixed Helvetica, RGB textColor, left|center|right alignment, and optional review metadata; it verifies all requested text survives in the native appearance and rejects unsupported styling, clipping, provider-unencodable content, stale evidence, or incremental save. add_area_annotation requires shape=rectangle|ellipse and a visible [x,y,width,height] bbox; optional strokeColor, 0.5–12 point borderWidth, and non-empty contents/author/subject are supported. It verifies one unfilled native Square/Circle with a solid border, exact rectangle/style, unique xref/count, and complete on-page appearanceBbox. Fill/dash/cloud/opacity/arbitrary appearance, unknown styles, edge-clipped strokes, stale evidence, and incremental save fail closed. add_text_markup requires markup=highlight|underline|strikeout|squiggly and one non-empty <=4096-character text string that native search finds exactly once on the visible page; optional RGB [red,green,blue] components in [0,1] and optional non-empty contents/author/subject are supported. It verifies one native Highlight, Underline, StrikeOut, or Squiggly plus quadrilaterals/color/appearanceBbox and rejects caller quads/rectangles, unknown or mis-cased styles, zero/multiple matches, stale evidence, an out-of-window appearance, and incremental save. add_text_highlight remains the same source-bound Highlight-only compatibility operation and rejects a markup field. add_link accepts an in-page-space [x,y,width,height], a non-duplicate target, and a safe internal #... or absolute http/https/mailto URL. These placement operations support right-angle rotated pages, report coordinateSpace/pageRotation, require rewrite, and require fresh output inspection. delete_page and duplicate_page take the exact source hash plus selected expectedPage={bbox,rotation}; rearrange_pages takes the same source hash, one changed complete 1-based page permutation, and expectedPages=[{page,bbox,rotation},...] for every current page in current order. Each page-tree mutation rejects Tagged PDFs, must be the only operation in a full rewrite, and requires fresh output inspect plus mapped render evidence. duplicate_page additionally takes an optional 1-based insertAt, defaults immediately after the source, requires a right-angle page without annotations, links, widgets/forms, page actions, associated files, article beads, transitions, or template steps, and adds no navigation. delete_embedded_file requires one canonical direct unique catalog EmbeddedFiles NameTree, the exact sourceSha256, one inspect-returned mupdf-embedded-file-... locator, and its complete expected snapshot; it verifies every non-target entry and requires rewrite. This removes that catalog entry only: payloadErasureClaimed and sanitizeClaimed remain false, while duplicate/nested/associated/portfolio graphs route explicitly to pikepdf cleanup or PyMuPDF sanitize. set_metadata requires the exact source hash, one mupdf-document-info locator, and the complete inspect-returned snapshot. Its patch accepts only the eight standard keys and uses null to clear. It fingerprints all non-target Info entries; for XMP, only field-safe existing properties listed in xmpMutableFields are accepted, both representations change together, and all non-target decoded packet bytes are proven unchanged. xmpBlockedFields gives field-specific refusal reasons: a unique x-default in a multilingual Alt, common scalar description attributes are supported, and a multi-author or irregular property blocks only itself. Missing or blocked requested properties, CDATA/DTD/invalid entities, malformed XML, partial/stale/unknown/empty-string/no-op evidence, and legacy unbound payloads fail closed. This is metadata editing, not sanitization. update_outline requires the exact source hash, one mupdf-outline-... locator, and the complete inspect-returned path/title/URI/open/page/child-count snapshot. Its patch accepts a non-empty control-free title and, only for a parent, boolean open; URI/page/path/child count, ordering, nesting, and every non-target outline remain fixed. It supports rewrite or unsigned byte-prefix-verified incremental save and rejects destination/topology edits, leaf open, partial/stale/unknown/no-op evidence, or more than the configured outline budget. update_form_field requires the exact source hash, one mupdf-form-field-<xref> locator, and its full field snapshot; it supports exactly one non-password text widget, one combo field whose display/export options are identical, or one checkbox, then verifies the saved in-memory field state. Shared-widget groups, radio/list/multi-select fields, mismatched choice exports, stale snapshots, and unsafe values fail closed and route to pypdf. delete_annotation and update_annotation each require one inspect-returned sourceSha256 and one source-bound mupdf-annotation-<page>-<xref> locator. Text-note update retains compatible non-empty contents/author/subject patches. Fixed-helvetica-v1 FreeText update requires the complete inspect-returned snapshot, accepts only contents/author/subject, preserves rectangle/appearance bounds/default appearance/alignment/flags/page/locator, and rejects clipped or provider-unencodable rebuilt text. Solid-no-fill-v1 Square/Circle update also requires the complete snapshot and accepts contents/author/subject/RGB color while preserving type/geometry/appearance bounds/no-fill state/border width/style/flags/page/locator and rechecking the on-page appearance. Native Highlight/Underline/StrikeOut/Squiggly update requires its complete inspect-returned expected snapshot and may patch those review fields plus RGB color, while type/quadrilaterals/rect/appearanceBbox/flags/page/locator remain invariant. Partial or stale snapshots, no-op/color-equivalent patches, caller geometry or border/fill styles, unsupported annotation profiles, and incremental save fail closed. delete_link and update_link require sourceSha256, a source-bound mupdf-link-<page>-<fingerprint> locator, and matching expected url/bbox/external facts. update_link accepts only one safe non-empty URL patch field; link bounds are snapshot evidence, never mutable geometry. set_page_crop remains unrotated-only and accepts a raw unrotated PDF-space bbox [x,y,width,height] fully inside the inspected MediaBox; it changes only CropBox and is never redaction. rotate_page accepts an absolute 0, 90, 180, or 270 degree /Rotate value; it does not transform or remove content.
- `savePolicy` (string) — rewrite or incremental. Incremental is forbidden for delete_page, duplicate_page, rearrange_pages, delete_embedded_file, redaction, source-bound annotation/link creation or mutation (including add_free_text_annotation, add_area_annotation, add_text_markup, and compatibility add_text_highlight), other delete operations, and signed input; source-bound single-widget form-field update, bounded Info/field-safe-XMP set_metadata, fixed-topology update_outline, set_page_crop, and rotate_page are unsigned operations that may be byte-prefix-verified incremental.
- `allowSigned` (boolean) — Acknowledge signed input after external review; never bypasses the incremental prohibition.
- `invalidateSignatures` (boolean) — Required with allowSigned for a deliberate signed-PDF rewrite.
- `password` (string) — Password for an encrypted PDF.
- `limits` (object) — Input/page/object budgets.

**Schema returns:**

- `blob` (FileBlob) — Edited PDF bytes with provider, save policy, signature state, byte counts, and applied-operation evidence.

#### `PdfFile.editPdf.add_area_annotation`

Add one source- and page-snapshot-bound unfilled native Square or Circle review outline to imported PDF bytes. The rewrite-only operation accepts rectangle/ellipse intent, explicit mupdf-page-space geometry, bounded solid RGB stroke width, and optional review metadata; stale evidence, painted bounds outside the page, fill/dash/cloud/opacity/arbitrary appearance, incremental save, and silent fallback fail closed. A fresh solid-no-fill-v1 snapshot may revise contents/author/subject/RGB color through update_annotation while geometry and border style remain fixed.

**Examples:**

- const inspection = await PdfFile.inspectPdf(pdf); const page = inspection.records.find((record) => record.kind === 'mupdfPage' && record.page === 1); await PdfFile.editPdf(pdf, { savePolicy: 'rewrite', operations: [{ type: 'add_area_annotation', page: 1, sourceSha256: inspection.summary.sourceSha256, expectedPage: { bbox: page.bbox, rotation: page.rotation }, shape: 'rectangle', bbox: [72, 196, 260, 96], strokeColor: [0.85, 0.1, 0.1], borderWidth: 3, contents: 'Confirm this region.' }] })

**Schema parameters:**

- `sourceSha256` (string) required — Exact SHA-256 from the current PdfFile.inspectPdf result.
- `expectedPage` (object) required — Exact inspect-returned {bbox, rotation} for the target mupdfPage.
- `shape` (string) required — Exactly rectangle or ellipse; mapped to one native Square or Circle annotation.
- `bbox` (number[]) required — Visible [x,y,width,height] in rotation-aware mupdf-page-space; both requested and painted appearance bounds must remain inside the page.
- `strokeColor` (number[]) — RGB outline color with each component in [0,1]; defaults to dark red.
- `borderWidth` (number) — Solid outline width from 0.5 through 12 points; defaults to 2.
- `contents` (string) — Optional non-empty review contents, at most 4,096 characters; not painted as page text.
- `author` (string) — Optional non-empty review author.
- `subject` (string) — Optional non-empty review subject.

**Schema returns:**

- `operation` (object) — Applied-operation evidence with native type, normalized rectangle, solid/no-fill style, appearanceBbox, coordinate space, unique xref/count, and appearanceBboxVerified=true. Re-inspect and render before delivery. A fresh solid-no-fill-v1 snapshot may update only contents/author/subject/RGB color through update_annotation, which re-verifies the on-page appearance; complete-snapshot deletion also remains available.

#### `PdfFile.editPdf.add_free_text_annotation`

Add one source- and page-snapshot-bound native FreeText review box to imported PDF bytes. The rewrite-only operation uses explicit mupdf-page-space geometry, a fixed Helvetica appearance, bounded font size/RGB text color/alignment, and native appearance-text verification; clipped text, off-page appearance, unsupported styling, stale evidence, incremental save, and silent fallback fail closed. A freshly inspected fixed-helvetica-v1 result may update only contents/author/subject through update_annotation with its complete snapshot and renewed native appearance-text verification.

**Examples:**

- const inspection = await PdfFile.inspectPdf(pdf); const page = inspection.records.find((record) => record.kind === 'mupdfPage' && record.page === 1); await PdfFile.editPdf(pdf, { savePolicy: 'rewrite', operations: [{ type: 'add_free_text_annotation', page: 1, sourceSha256: inspection.summary.sourceSha256, expectedPage: { bbox: page.bbox, rotation: page.rotation }, bbox: [72, 128, 260, 56], contents: 'Review this assumption.', fontSize: 12, textColor: [0.1, 0.2, 0.8], alignment: 'left' }] })

**Schema parameters:**

- `sourceSha256` (string) required — Exact SHA-256 from the current PdfFile.inspectPdf result.
- `expectedPage` (object) required — Exact inspect-returned {bbox, rotation} for the target mupdfPage.
- `bbox` (number[]) required — Visible [x,y,width,height] in rotation-aware mupdf-page-space, fully inside the inspected page bbox.
- `contents` (string) required — Non-empty visible review text, at most 4,096 characters; unsupported controls and clipped native appearance fail closed.
- `fontSize` (number) — Helvetica appearance size from 4 through 72 points; defaults to 12.
- `textColor` (number[]) — RGB text color with each component in [0,1]; defaults to black.
- `alignment` (string) — left, center, or right; defaults to left.
- `author` (string) — Optional non-empty review author.
- `subject` (string) — Optional non-empty review subject.

**Schema returns:**

- `operation` (object) — Applied-operation evidence with normalized native annotation, appearance bbox, coordinate space, style, and appearanceTextVerified=true. Re-inspect and render the rewrite before delivery. A fresh fixed-helvetica-v1 record may update only contents/author/subject through update_annotation with its complete snapshot and renewed appearanceTextVerified=true evidence.

#### `PdfFile.exportPdf`

Export a modeled artifact as a real multi-page tagged PDF 1.7 whose logical structure follows explicit page reading order without changing paint order, emits semantic H1-H6 headings, meaningful Figure /Alt text, Link annotations with OBJR associations, /Artifact marked content, and constrained logical Tables spanning consecutive pages, and preserves language/title, Table/TR/TH/TD hierarchy, optional Unicode TrueType embedding, positioned text, vector charts, and PNG/JPEG images.

**Examples:**

- const blob = await PdfFile.exportPdf(pdf, { language: 'en-US', title: 'Accessible report' })

**Schema parameters:**

- `pdf` (PdfArtifact) required — Modeled PDF artifact to serialize.
- `tagged` (boolean) — Emit StructTreeRoot/ParentTree/MCID tagging; defaults to true.
- `language` (string) — Catalog language; defaults to artifact metadata language or en-US.
- `title` (string) — Document Info title; defaults to artifact metadata title or first text line.
- `font` (string|FileBlob|Uint8Array|ArrayBuffer|object) — Optional standalone glyf-based TrueType .ttf source for Unicode Type0/CIDFontType2 embedding; accepts a path, bytes, FileBlob, or {path|bytes|base64}.
- `maxFontBytes` (number) — Maximum accepted embedded font input size; defaults to 16 MiB.
- `subsetFont` (boolean) — Subset the embedded TrueType font to used glyphs and composite dependencies; defaults to true. Set false only for diagnostics/interoperability comparison.

**Schema returns:**

- `blob` (FileBlob) — application/pdf bytes with modeled content, clean-room metadata, and tagged-export metadata.

#### `PdfFile.importPdf`

Reopen package-generated metadata losslessly or lazily use required MuPDF.js for arbitrary PDFs, producing a bounded reconstructed extraction/QA view with text geometry, raster placements and transforms, links, annotations, widgets, and heuristic table candidates; the view is never an edit representation.

**Examples:**

- await PdfFile.importPdf('third-party.pdf', { limits: { maxBytes: 64 * 1024 * 1024 }, includeImages: true })
- await PdfFile.importPdf(blob, { parser: createPdfjsParser(), preferParser: true })

**Schema parameters:**

- `blob` (string|FileBlob|Uint8Array|ArrayBuffer) required — PDF path or input bytes. Paths and Blob-like inputs are size-checked before materialization.
- `parser` (function) — Optional parser adapter returning pages/textItems/tables/images.
- `preferParser` (boolean) — Use parser even if clean-room metadata is embedded.
- `parserName` (string) — Name recorded in artifact metadata.
- `password` (string) — Password for an encrypted PDF.
- `includeImages` (boolean) — Extract bounded raster placements; defaults to true.
- `limits` (object) — maxBytes, maxPages, maxObjects, maxImages, maxImagePixels, maxTotalImagePixels, and maxTotalImageBytes budgets.

**Schema returns:**

- `pdf` (PdfArtifact) — Modeled PDF artifact with inspect/resolve/render/verify APIs.

#### `PdfFile.inspectPdf`

Inspect a path or PDF bytes after a pre-WASM input budget, combining native MuPDF page/object/annotation/widget/link/outline records, source SHA-256, source-bound Document Info records plus annotation/link/outline/canonical catalog EmbeddedFiles locators, complete XMP stream identity/profile/mutable-and-blocked-field evidence, native annotation snapshots and update capabilities plus FreeText default-appearance/alignment, area border/style, and text-markup quadrilateral/color facts, raw MediaBox/CropBox facts, and effective normalized page rotation with bounded tagged-PDF, language, reading-order, heading, Figure, Link, Artifact, font, and table-structure evidence.

**Examples:**

- await PdfFile.inspectPdf(pdf, { maxObjects: 200, maxChars: 12000 })

**Schema parameters:**

- `pdf` (string|FileBlob|Uint8Array|ArrayBuffer) required — PDF path or bytes.
- `limits` (object) — Input, page, object, annotation/widget, link, and outline budgets applied before or during native inspection.
- `maxObjects` (number) — Maximum indirect object records to inspect.
- `maxLinks` (number) — Maximum native link records to inspect or reconstruct.
- `maxOutlines` (number) — Maximum flattened native outline records; defaults to 100,000 with a separate 256-level depth bound.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `inspection` (object) — PDF file summary with sourceSha256, Document Info and metadata-update capability, outline count, tagged/language/structure evidence, bounded indirect object records, one complete source-bound mupdfDocumentMetadata snapshot, flattened source-bound mupdfOutline path/snapshot/capability records, source-bound mupdfAnnotation/mupdfLink/mupdfWidget records, and grouped mupdfFormField snapshots. The metadata record fingerprints raw Info entries and, for catalog XMP, records stream object/hash/length, field-safe profile, decoded standard values, mutable fields, and field-specific blocked reasons. A unique x-default in a multilingual Alt and common scalar description attributes are editable; a multi-author or irregular standard property blocks only that field, while malformed packet structure remains fail-closed. Outline records retain path, title, URI, open state, resolved page, child count, and a fingerprinted locator; title is mutable, while open is mutable only for a parent, and destination/topology stay fixed. Native annotation records carry a complete current-source snapshot and field-level updateCapability. Text-note and text-markup annotations expose a provider appearanceBbox; Highlight, Underline, StrikeOut, and Squiggly records also expose their quadrilateral selection and RGB color, and only those text-markup types advertise bounded color mutation. Native page records include raw unrotated PDF-space MediaBox/CropBox [x,y,width,height] values, normalized right-angle rotation, and an effective visible bbox with coordinateSpace=mupdf-page-space after that rotation. Re-inspect after any rewrite because locators cannot identify a later byte sequence.

#### `PdfFile.renderPdf`

Render one page from original PDF bytes through runtime-lazy MuPDF.js as PNG or JPEG, enforcing input, page/object, DPI, and preallocation pixel budgets before returning a FileBlob.

**Examples:**

- await PdfFile.renderPdf(pdf, { page: 1, dpi: 144, format: 'png' })

**Schema parameters:**

- `pdf` (string|FileBlob|Uint8Array|ArrayBuffer) required — Original PDF path or bytes.
- `page` (number) — One-based page number; defaults to 1.
- `dpi` (number) — Resolution greater than 0 and no more than 1200; defaults to 144.
- `format` (string) — png or jpeg.
- `quality` (number) — JPEG quality from 1 through 100.
- `password` (string) — Password for an encrypted PDF.
- `limits` (object) — Input/page/object and maxRenderPixels budgets.

**Schema returns:**

- `blob` (FileBlob) — Native PNG or JPEG page bytes with provider, page, DPI, and dimensions metadata.

#### `PdfProviders.ensure`

Install only a previously installable, policy- and catalog-bound capability resolution into the project-private cache, then return a fresh probe. It uses only catalog-pinned release bytes, validates size/hash/archive/receipt boundaries, and never chooses another provider or obtains credentials. A blocked or ready resolution cannot be forced through this API.

**Examples:**

- const installed = await PdfProviders.ensure({ resolution, policyPath: '.office-kit/pdf-providers.json' })

**Schema parameters:**

- `resolution` (object) required — The exact current-package resolution returned with status=installable. Its catalog digest, policy fingerprint, provider, and pack plan must still match.
- `policyPath` (string) — The same persistent project policy file used to resolve. It is re-read before any cache mutation.

**Schema returns:**

- `result` (object) — Fresh ready/blocked provider probe plus verified installation receipts. It can only use pinned catalog assets, a bounded project-private cache, hash/size checks, safe extraction, and atomic publication; it never downloads credentials or falls back.

#### `PdfProviders.probe`

Probe exactly one selected PDF provider under the requested policy without downloading, mutating the cache, importing MuPDF, or trying fallback providers. The result reports ready or blocked runtime evidence together with the pinned pack plan.

**Examples:**

- const state = await PdfProviders.probe({ provider: 'qpdf', task: 'repair', policyPath: '.office-kit/pdf-providers.json' })

**Schema parameters:**

- `provider` (string) required — One exact catalog provider ID; probing does not search for alternates.
- `task` (string) — Optional catalog task used to include OCR language-pack requirements in the plan.
- `policyPath` (string) — Explicit project policy file; default-missing remains disabled.
- `languages` (string[]) — Explicit OCR language list, required for an OCR task and checked against policy plus catalogued language packs.

**Schema returns:**

- `state` (object) — Ready or blocked status with local/system/managed runtime evidence and the selected pack plan. It performs no network request, cache write, MuPDF import, or provider fallback.

#### `PdfProviders.resolve`

Resolve one explicit PDF task and selected/default provider against the immutable capability catalog and project policy. It is read-only: no MuPDF initialization, network access, cache mutation, credential acquisition, or automatic provider fallback occurs. The result is ready, installable, or blocked with exact packs, platform, sizes, licenses, runtime prerequisites, consents, and operation limits.

**Examples:**

- const resolution = await PdfProviders.resolve({ task: 'repair', provider: 'qpdf', savePolicy: 'rewrite', mutationAuthorized: true, invalidateSignaturesAuthorized: true, policyPath: '.office-kit/pdf-providers.json' })

**Schema parameters:**

- `task` (string) required — One catalog task such as inspect, repair, ocr, sign, sanitize, or validate-conformance.
- `provider` (string) — Optional provider ID. A task default is a declared preference only; the resolver never substitutes a different provider when it is unavailable.
- `inspection` (object) — Exact-source inspection/preflight evidence. Required for every existing-PDF task except inspect; it must carry a 64-hex sourceSha256 at inspection.summary.sourceSha256 or sourceSha256. A failed MuPDF parse may use a bounded preflight hash record only to route explicit repair.
- `savePolicy` (string) required — One strategy allowed by the selected task, such as read-only, rewrite, incremental, or sanitize.
- `policyPath` (string) — Explicit project policy file. The conventional path is .office-kit/pdf-providers.json; a missing conventional file means disabled, never implicit authorization.
- `languages` (string[]) — Explicit OCR languages. eng and chi_sim are policy defaults; every language must be policy-authorized and catalogued.
- `mutationAuthorized` (boolean) — Required true for a task that mutates source PDF bytes.
- `invalidateSignaturesAuthorized` (boolean) — Required true for a task whose operation can invalidate signatures.
- `credentials` (string[]) — Caller-declared credential kinds such as local-pkcs12. Credentials, private keys, HSMs, remote-signing access, and TSA/LTV access are never installed or acquired.

**Schema returns:**

- `resolution` (object) — Read-only ready, installable, or blocked resolution with one provider, catalog digest, policy fingerprint, no-fallback guarantee, precise pack/platform/download/unpacked/license/runtime plan, and required consents.

## presentation

| Name | Kind | Summary |
| --- | --- | --- |
| `chart.accessibilityCapability` | api | Report sourceBound/editable/addable preflight for chart graphic-frame p:cNvPr title/description/decorative metadata; export re-proves it. |
| `chart.delete` | api | Explicitly remove a source-free chart or one capability-proven imported chart frame. The source-bound transaction removes its exact relationship and only ChartPart descendants without another package parent; external, repeated, nested, or identity-sensitive graphs fail closed. |
| `chart.deletionCapability` | api | Report whether one imported top-level chart frame owns one uniquely used internal ChartPart relationship. Export re-proves relationship use and the exclusively owned descendant closure; shared ChartParts survive. |
| `chart.setAccessibilityMetadata` | api | Transactionally add, change, or clear non-visible chart title/description/decorative metadata independently of its visible chart title. Imported irregular graphic-frame p:cNvPr graphs fail closed. |
| `compose.box` | api | Create a materialized shape surface with optional children inset by padding; use a named box as a stable connector or edit target. |
| `compose.chart` | api | Create a materialized chart in a resolved compose frame; encode quantitative claims as data relationships rather than decorative labels. |
| `compose.column` | api | Create a vertical compose container. Use width/height fill, hug, or fixed pixels; gap and padding are in pixels. |
| `compose.grid` | api | Create a grid compose container with bounded row/column tracks, spans, gaps, padding, and an optional surface. |
| `compose.image` | api | Create a materialized image node with frame, fit/crop, alt text, and an explicit user or template asset; a prompt creates only a marked placeholder. |
| `compose.layers` | api | Create a layered compose container whose children share the inner frame; use it for overlays and explicit z-order. |
| `compose.paragraph` | api | Create an editable text block with name, className/style text tokens, and stable inspect output. |
| `compose.row` | api | Create a horizontal compose container. Use fixed, hug, or fill child widths with an explicit gap and optional surface. |
| `compose.rule` | api | Create a thin horizontal or vertical rule as a materialized shape, using the resolved frame and stroke token. |
| `compose.shape` | api | Create a materialized native shape, including text-bearing shapes and straight connectors, from a declarative compose node. |
| `compose.table` | api | Create a materialized table in a resolved compose frame; keep the table data and column/row budget explicit. |
| `compose.text` | api | Create the same editable paragraph node through the reference-template-compatible children-first text(children, props) helper. |
| `connector.accessibilityCapability` | api | Report sourceBound/editable/addable preflight for connector p:cNvPr title/description/decorative metadata; export re-proves it. |
| `connector.bringToFront` | api | Move a connector above peers in its slide/group scene stack. An imported direct connector may move only when its fresh zOrderCapability is editable; unsupported or nested native topology rejects. |
| `connector.delete` | api | Explicitly remove a source-free connector or one capability-proven imported direct p:cxnSp. Relationship-bearing or nested connectors and connector/comment/timing/extension identity consumers fail closed; endpoint shapes remain untouched. |
| `connector.deletionCapability` | api | Report whether one imported top-level canonical relationship-free connector can be deleted, with a package-local native ID used for post-write absence proof. Export recomputes the source-bound capability. |
| `connector.sendToBack` | api | Move a connector behind peers in its slide/group scene stack. New shape-connected connectors start behind their nodes; an imported direct connector requires an editable zOrderCapability. |
| `connector.setAccessibilityMetadata` | api | Transactionally add, change, or clear non-visible connector title/description/decorative metadata. Imported irregular p:cNvPr graphs fail closed without disabling unrelated supported edits. |
| `connector.setConnectorFrom` | api | Atomically bind a connector start to a modeled same-tree shape and explicit connection-site index. |
| `connector.setConnectorTo` | api | Atomically bind a connector end to a modeled same-tree shape and explicit connection-site index. |
| `element.bringToFront` | api | Move a shape, image, table, chart, connector, or group to the front of its current slide/group scene stack. Imported direct elements require a current editable zOrderCapability; unsupported native topology fails closed. |
| `element.moveAfter` | api | Move one presentation element immediately in front of a different peer in the same scene stack, subject to the same source-bound capability and source-prefix checks. |
| `element.moveBefore` | api | Move one presentation element immediately behind a different peer in the same scene stack, subject to the same source-bound capability and source-prefix checks. |
| `element.sendToBack` | api | Move a shape, image, table, chart, connector, or group to the back of its current slide/group scene stack. Imported direct elements require a current editable zOrderCapability; authored overlays on an imported slide cannot move below the complete source-bound prefix. |
| `element.stackIndex` | api | Return an element's current zero-based position in its owning slide or group scene stack, where zero is farthest back. |
| `element.zOrderCapability` | api | Return fresh { sourceBound, known, editable, blockedReason } evidence for moving an element in its owner scene stack. Imported direct elements are editable only when the codec issued and export can re-prove the capability. |
| `exportPptxWithOfficeKit` | api | Export bounded direct slide backgrounds, textbox/rectangle/roundRect/ellipse shapes, free-positioned p:sp lines with the shared six-style/line-end/cap/join outline profile, rich text and lists, basic fills/lines/shadows, straight/elbow/curved p:cxnSp connectors with target connection sites through that same line profile, embedded pictures with native crop/contain/cover semantics, fixed-grid plain-text tables, recursive native p:grpSp trees, relationship-free rich speaker notes, legacy annotations, Office 2021 modern root/direct-reply threads, source-free bar/line/pie charts, the bounded literal clustered bar+line combo profile with either shared primary axes or a canonical secondary line pair, validated payload-only replacement for eligible imported OLE XLSX workbooks plus the uniquely bound DOCX Office-package profile, and bounded source-bound text updates for canonical SmartArt document nodes. Recognized imported modern threads allow only existing text/status edits; their identity, author/date metadata, anchor/range, position, topology, relationships, and source hashes remain fixed. Inherited or complex graphs remain preserved and fail closed on unsupported mutation. |
| `group.accessibilityCapability` | api | Report sourceBound/editable/addable preflight for group-frame p:cNvPr title/description/decorative metadata; export re-proves it. |
| `group.delete` | api | Delete one source-free or capability-proven imported group as a complete recursive ownership tree. Shared media and ChartParts survive; nested groups, outside connector/comment targets, relationship reuse, identity-sensitive graphs, and raw collection mutation fail closed. |
| `group.deletionCapability` | api | Report whether one imported top-level canonical recursive p:grpSp exclusively owns its complete native-ID, relationship-reference, and multi-root OPC graph. Export recomputes the source-bound capability. |
| `group.setAccessibilityMetadata` | api | Transactionally add, change, or clear non-visible group-frame title/description/decorative metadata. Imported irregular p:cNvPr graphs fail closed without disabling unrelated supported edits. |
| `image.accessibilityCapability` | api | Report sourceBound/editable/addable preflight for picture p:cNvPr title/description/decorative metadata; export re-proves the residual-protected picture profile. |
| `image.delete` | api | Explicitly remove a source-free image or one capability-proven imported top-level embedded picture. The source-bound transaction removes the p:pic subtree and exact relationship, garbage-collects only exclusively owned media descendants, preserves shared media, and rejects external/ambiguous/identity-sensitive graphs or raw array mutation. |
| `image.deletionCapability` | api | Report whether one imported top-level embedded picture can be deleted with its exact SlidePart relationship and exclusively owned media closure. Shared media survives; export re-proves relationship use, native identity, comments, connectors, timing, and extensions from source bytes. |
| `image.editSvgLeaf` | api | Replace one issued SVG RGB, opacity, or transform scalar after expectedHash verification. The exact token splice preserves all other SVG bytes and rejects stale, cross-image, invalid, unsupported, and no-op edits. |
| `image.editSvgText` | api | Replace one issued direct SVG text/tspan leaf after expectedHash verification with an escaped value. The bounded image-byte transaction preserves the rest of the SVG, rejects active/external content and stale/no-op edits, and remains verifiable after PPTX export/reimport. |
| `image.getSvgEditLeaves` | api | Return defensive source-issued SVG style and transform leaves for an image. Each leaf identifies its typed value and exact expectedHash without exposing XML selectors or arbitrary attributes. |
| `image.getSvgTextNodes` | api | Return defensive source-issued SVG text/tspan leaves for an image. Each leaf has a stable image-local ID, text, tag, and expectedHash; the returned records cannot mutate the image. |
| `image.setAccessibilityMetadata` | api | Transactionally add, change, or clear a picture's non-visible title/description/decorative metadata. The legacy image.alt property reads and writes the same description state rather than creating a second metadata source. |
| `image.svgEditCapability` | api | Report source-revision-bound direct SVG fill, stroke, opacity, and single transform-scalar leaves for a base64 SVG image. Each issued leaf carries an exact replacement hash; active content, external references, stylesheets, classes, and unsupported transform topology remain blocked. |
| `image.svgTextCapability` | api | Report bounded direct SVG text/tspan leaves for a base64 SVG image, including the image-byte SHA-256 and exact replacement hashes. Active content, external references, oversized SVGs, and nested/non-text leaves remain unsupported. |
| `importPptxWithOfficeKit` | api | Import PPTX bytes with editable bounded direct slide backgrounds, shapes, free-positioned p:sp lines including bounded line ends/caps/joins, rich text, recognized owner-local SlidePart placeholder text, rectangular pictures and native source rectangles, tables, target-bound p:cxnSp connectors, recursive canonical p:grpSp groups, bar/line/pie charts, the canonical literal clustered bar+line combo profile with either shared primary axes or a secondary line pair, legacy text-only speaker notes plus fixed-topology relationship-free rich notes and a re-proven addable capability for eligible notes-absent slides, unchanged-only legacy comments, fixed-topology modern comment text/status edits, defensive payload access for eligible OLE XLSX workbooks plus one uniquely bound DOCX Office-package profile, and a source-bound SmartArt text capability only for a canonical closed four-part DiagramDataPart whose nodes use fixed direct paragraphs with optional empty paragraphs, between one and 256 total direct plain runs, and canonical fixed a:br leaves. Compound/theme/custom-dash/effect/extension outlines and all other unsupported content remain source-bound and read-only rather than being flattened. |
| `nativeObject.getEmbeddedOfficePackage` | api | Read a defensive FileBlob copy from an eligible source-bound top-level OLE package. It is compatible with the legacy XLSX workbook profile and currently adds one uniquely bound DOCX profile; it never exposes arbitrary OLE or native-part mutation. |
| `nativeObject.getEmbeddedWorkbook` | api | Read a defensive FileBlob copy of the XLSX payload from an eligible source-bound top-level OLE object without exposing arbitrary native-part mutation. |
| `nativeObject.replaceEmbeddedOfficePackage` | api | Replace only a source-bound Office package on an eligible imported top-level OLE object. The current generic profile validates DOCX bytes and exact content type while preserving the OLE shell, relationships, preview, and all other native parts; malformed, shared, ambiguous, or unsupported package graphs fail closed. |
| `nativeObject.replaceEmbeddedWorkbook` | api | Replace only the XLSX payload of an eligible imported top-level OLE object. OfficeKit validates the new workbook and immutable source binding, preserves the OLE shell, relationships, preview, and all other native parts, and fails closed for malformed or ambiguous graphs. |
| `nativeObject.setDiagramNodeRunText` | api | Replace one existing direct a:r/a:t value by zero-based source-order run index across a proven SmartArt node's fixed direct paragraphs. Empty paragraphs, paragraph/run topology, a:pPr, a:rPr, canonical fixed a:br, and a:endParaRPr stay source-bound; wholly empty nodes, fields, noncanonical breaks, topology changes, and unsupported diagrams reject without fallback. |
| `nativeObject.setDiagramNodeText` | api | Replace a one-run source-bound SmartArt document node after its top-level four-part graph and fixed direct-paragraph/run DiagramDataPart profile are proven. Multi-run nodes reject so OfficeKit never guesses a formatting boundary. |
| `nativeObject.setName` | api | Native OLE, SmartArt/diagram, contentPart, and media objects imported through OfficeKit are source-bound and read-only for names; setName rejects instead of mutating the preserved package graph. Separate bounded SmartArt node/run text methods own the only modeled diagram mutation. |
| `nativeObject.setPosition` | api | Native OLE, SmartArt/diagram, contentPart, and media objects imported through OfficeKit are source-bound and read-only; setPosition rejects instead of rewriting their geometry or payload graph. |
| `presentation.auditAccessibility` | api | Audit modeled slide objects for explicit meaningful/decorative classification and non-visible title/description coverage, while separating native-object and reading-order checks that still require manual host review. It never claims whole-deck accessibility conformance. |
| `Presentation.create` | api | Create a deck model whose canonical OfficeKit export supports ordinary slides, the complete ECMA-376 base slide-transition vocabulary, direct solid/style-reference slide backgrounds, shapes, rich text, tables, images, connectors, recursive native p:grpSp groups, plain-text speaker notes, native custom shows with canonical run links, literal bar/line/pie/standard-area/fixed-doughnut/marker-scatter/2D-bubble charts, and a bounded literal clustered bar+line combo profile. Combo bars stay on the primary pair; all lines share either that pair or the canonical secondary top/right pair. Formula/external chart data, custom themes, Master/Layout authoring, comments, custom-show topology mutation, advanced plot geometry, mixed line groups, secondary bars, irregular combo graphs, and other package-level features remain outside the source-free PPTX boundary. |
| `presentation.customShows.add` | api | Define an ordered native p:custShowLst playback route for source-free OfficeKit export. Text runs may target a show by exact name with optional returnToSlide. Canonical imported shows may change only their name and ordered retained-slide membership; fixed native identity keeps existing run links bound across a rename, while irregular graphs stay opaque. |
| `presentation.customShows.getItem` | api | Resolve a source-free or canonical imported custom show by zero-based index, stable facade ID, or exact name. |
| `presentation.designProfile` | api | Return a bounded read-only design-language profile for the current deck: source revision binding when imported, canvas, palette, typography, density, normalized geometry rhythm, layout families, slide archetypes, repeated visual candidates, and opaque native summaries. The profile is evidence for template-conditioned generation only; it contains no XML selectors, package paths, source bytes, or mutation authority. |
| `presentation.editComponentOccurrence` | api | Apply one atomic batch of typed native-leaf edits to a repeated component occurrence issued by presentation.inspect({ includeComponentCandidates: true }). The occurrence editCapability and each leafId, targetId, and expectedHash are source-revision-bound; all values are validated before any leaf is changed. Only codec-issued text, color, geometry, chart, SmartArt, or other bounded leaf kinds are accepted. Raw XML, selectors, part paths, foreign leaves, duplicate leaves, stale hashes, and edits outside the selected component fail closed. |
| `presentation.editNativeLeaf` | api | Change one native leaf issued by presentation.inspect({ includeNativeLeaves: true }) using its targetId, leafId, expectedHash, and a typed value. Leaf IDs are bound to the exact imported revision and target. Repeat the call for a coordinated move/resize; one export sorts all issued leaves into one deterministic Edit Plan. The current profile changes existing text leaves, including group children and shapes with source-owned outer styling, shape RGB/local-geometry scalars, picture local-geometry scalars (including opaque pictures whose payload and effects remain source-owned), direct rich chart-title runs, direct numeric bar-chart cache points proven against one exact cell in a uniquely bound embedded XLSX, direct SmartArt text runs from one canonical closed DiagramDataPart with a unique inbound owner, explicit bare text-body AutoFit choices, direct column-direction flags, direct vertical-text modes, or explicit literal paragraph/run font-size/typeface/style/color/decoration/alignment leaves (`paragraphAlignment`, `verticalAnchor`, `fontSizePoints`, `fontFamily`, `fontFamilyEastAsia`, `fontBold`, `fontItalic`, `fontColorRgb`, `fontColorScheme`, `fontUnderline`, `fontStrike`, `fontKerningPoints`) proven on one direct text run. Paragraph alignment is limited to a direct canonical `a:pPr/@algn` token (`left`, `center`, `right`, or `justify`); vertical text anchoring is limited to a direct canonical `a:bodyPr/@anchor` token (`top`, `center`, or `bottom`); text-body AutoFit is limited to a direct bare `a:noAutofit`, `a:normAutofit`, or `a:spAutoFit` child (`none`, `shrinkText`, or `resizeShape`); column direction is limited to a direct canonical `a:bodyPr/@rtlCol` token (`0` or `1`); vertical text is limited to a direct canonical `a:bodyPr/@vert` token (`horz`, `vert`, or `vert270`, exposed as `horizontal`, `vertical`, or `vertical270`); underline and strike are limited to standard direct DrawingML tokens; kerning is limited to a direct non-negative `a:rPr/@kern` token, exposed in points and spliced as hundredths of a point. Inherited, malformed, effect-bearing, or otherwise irregular style graphs remain opaque. A chartDataValue operation changes both the ChartPart cache and that worksheet cell. A diagramText operation token-splices only its issued a:t and does not reserialize the diagram part. Separate typed imported-table, embedded-image, and element-delete facades lower to tableCellText, imageAsset, and deleteElement operations in the same Edit Plan; those operation kinds are not arbitrary native-leaf selectors. The compiler binds the complete ownership tree and dependent parts. Stale hashes, concurrent non-leaf changes, foreign IDs, raw XML, XPath, part paths, arbitrary attributes or cells, relationship fields, formulas, namespaces, and topology changes reject. |
| `presentation.export` | api | Export a slide SVG preview, deck SVG montage via { format: 'montage' }, or target/search-sliced layout JSON. |
| `presentation.fontFamilies` | api | Return a fresh sorted, case-insensitively deduplicated list of explicitly used presentation text and bullet font families. |
| `presentation.inspect` | api | Emit NDJSON for deck, custom shows, PowerPoint sections, slides, cross-type layers, direct slide transitions, textboxes, shapes, grouped shapes, tables, charts, images, and native contentPart/OLE/diagram/media objects with bounded editability, relationship-reference, root-relationship, preserved-part, eligible embedded Office-package summaries, and each slide's continuationCapability; narrow with search/target anchors and shape fields with include/exclude. Layer records expose bottom-to-top stackIndex and zOrderCapability without exposing package paths. On a trusted imported source, includeNativeLeaves: true returns revision-bound safe leaves without exposing part paths or XML selectors, while includeComponentCandidates: true returns repeated visual primitives with source hashes, occurrences, and explicit reuse limits; only closed top-level candidates can issue the bounded reuseSourceComponent operation. |
| `presentation.layout.clearBackground` | api | Clear a direct background on a bounded source-free layout. Imported-layout mutation remains source-bound and fails closed. |
| `presentation.layout.placeholders.add` | api | Append a direct-frame title/body/ctrTitle/subTitle text placeholder to a source-free layout. It becomes a native p:ph and must be materialized on each slide through applyLayout/setLayout; object/media/chart/table placeholders remain source-bound. |
| `presentation.layout.placeholders.summary` | api | Return a defensive layout-placeholder discovery snapshot with stable IDs, names, native types/indexes, required flags, and direct-frame presence/geometry; editing the snapshot cannot mutate the model. |
| `presentation.layout.setBackground` | api | Set a direct background on a bounded source-free layout. Imported-layout mutation remains source-bound and fails closed. |
| `presentation.layouts.add` | api | Create one bounded source-free layout under the canonical master. Use blank, title, titleOnly, or obj/titleAndContent plus direct-frame text placeholders; imported layouts remain source-bound and read-only. |
| `presentation.layouts.getById` | api | Resolve a layout by its stable ID without falling back to a same-named or same-typed layout. |
| `presentation.master` | api | Access the one canonical source-free Slide Master. It may author a direct background, bounded text styles, and direct-frame title/body/ctrTitle/subTitle placeholders; imported Master graphs remain source-bound and read-only. |
| `presentation.master.clearBackground` | api | Clear the direct background of the one canonical source-free master. Imported-master mutation remains source-bound and fails closed. |
| `presentation.master.setBackground` | api | Set the direct background of the one canonical source-free master. Imported-master mutation remains source-bound and fails closed. |
| `presentation.master.setTheme` | api | Set a model-level master theme override for preview only. Canonical PPTX export rejects that source-free override; imported-master mutation remains source-bound and fails closed. |
| `presentation.masters.add` | api | Append a model-level Slide Master. Source-free PPTX authoring requires exactly one master, so use Presentation.create({ master }) or presentation.master for the canonical profile; multiple masters and imported-master edits fail closed. |
| `presentation.masters.getItem` | api | Resolve a model-level or imported Slide Master by stable ID or name. |
| `presentation.planTemplateGeneration` | api | Build a source-bound, read-only multi-page frame map from a trusted imported PPTX: choose clone-safe source slides by role, archetype, content density, and preferred visual kinds; issue bounded text-run targets and reusable-component candidates; report heuristic text-fit warnings, alternatives, opaque-object limits, and blocked requests without mutating the deck. |
| `presentation.resolve` | api | Map stable inspect anchor IDs back to facade objects, including custom shows, PowerPoint sections, and slide transitions; imported advanced package objects may be read-only. |
| `presentation.resolveComponentCandidate` | api | Resolve one candidateId issued by presentation.inspect({ includeComponentCandidates: true }) to a defensive source-revision-bound reference. Candidates describe repeated visual structure without exposing raw XML or asset bytes; only an inspect-only candidate with a closed top-level graph can be passed to presentation.reuseSourceComponent, while ambiguous, opaque, or relationship-bound graphs carry an explicit blocked reason. |
| `presentation.reuseSourceComponent` | api | Create a new source-bound slide containing one exact top-level repeated component occurrence from presentation.inspect({ includeComponentCandidates: true }). The candidateId, occurrenceIndex, source revision, closed-graph ownership, sibling deletion proofs, and retained connector targets are checked before a complete source slide clone is projected by deleting only codec-proven sibling elements. Nested, opaque, ambiguous, comment-bound, relationship-bound, or stale candidates fail closed; the original slide and all non-target source parts remain untouched. |
| `presentation.reuseSourceSlide` | api | Reuse one inspected imported slide as a source-bound complete graph after matching its exact slideId, sourceRevisionSha256, and optional clone-capability ownership evidence. The operation delegates to the codec-proven slide clone profile; stale revisions, unsupported graphs, and mismatched ownership evidence fail closed before the pending clone is created. |
| `presentation.sections.add` | api | Define a native PowerPoint p14:sectionLst entry for source-free OfficeKit export. Sections together must form the complete ordered slide partition. Canonical imported sections may change only existing names and contiguous boundaries while count, order, stable facade identity, and native GUID stay fixed; irregular graphs remain opaque. |
| `presentation.sections.getItem` | api | Resolve a source-free or canonical imported PowerPoint section by zero-based index, stable facade ID, or exact name. |
| `presentation.slides.add` | api | Append an editable core slide with optional hidden slideshow state, a bounded source-free layout, direct ECMA-376 base transition, solid/style-reference background, and plain-text speaker notes. A supplied layout is resolved and materialized transactionally; effective imported Layout/Master inheritance is never flattened. |
| `presentation.slides.insert` | api | Insert a source-free slide after an existing Slide or 0-based index, or at the beginning with after: null. It uses the same hidden-state, transactional layout, direct base-transition, notes, and background profile as slides.add; imported additions fail closed, while slide.duplicate and slide.delete each have their own narrow source-preserving OPC profiles. |
| `presentation.slideSize` | api | Read or set the deck canvas in pixels. On a trusted imported PPTX, a changed size is a deliberately canvas-only source-bound operation: OfficeKit updates only ppt/presentation.xml p:sldSz, clears an old preset type, and leaves slide, layout, master, chart, and shape coordinates unchanged. It never silently rescales or reflows content; callers must make any layout edits explicitly. |
| `presentation.textRange` | api | Inspect or resolve stable textRange anchors such as shapeId/text for editable slide text frames. |
| `presentation.theme` | api | Inspect the model theme and theme inheritance. Custom source-free themes are not authored by OfficeKit 0.2, and imported themes are source-bound and read-only. |
| `presentation.validateLayout` | api | Detect layout QA issues across slides, including off-canvas elements, geometry overlaps, and basic text overflow. Explicit text-free accessibility.decorative objects are excluded from overlap and partial-bleed errors; confirm their crop in the rendered slide. |
| `presentation.verify` | api | Return QA issues for layout validation, missing master/layout references, placeholder fidelity, chart/data consistency, table shape, image data, and dangling comments. |
| `presentation.view` | api | Control local editor gridline/guide visibility and inspect imported PowerPoint grid spacing, snap settings, and guides. Visibility is local model state; a separately capability-gated fixed-topology source-bound edit profile may change only already-present grid/snap values and guide positions in viewProps.xml. |
| `presentation.view.capability` | api | Return defensive sourceBound, partPresent, editable, existing-field, and guide-count evidence for the imported PPTX view-properties part. It is preflight evidence only; export re-proves hashes, topology, and the non-editable XML residual. |
| `presentation.view.setSourceProperties` | api | Change already-present imported grid spacing, snap flags, and existing guide positions only when view.capability.editable is true. It cannot create viewProps.xml, add/remove/reorient guides, write showGuides, or reconstruct extensions/relationships; unsupported profiles fail closed. |
| `PresentationFile.exportPptx` | api | Serialize PPTX through the single bundled OfficeKit codec. Only limits is accepted; legacy codec and lossy-fallback options fail explicitly. |
| `PresentationFile.importPptx` | api | Import PPTX through the single bundled OfficeKit codec with bounded free-positioned p:sp lines including direct line ends/caps/joins, source-bound opaque preservation, speaker-notes edit/add capability evidence, bounded text-only edits for recognized local SlidePart placeholders and canonical SmartArt nodes whose fixed direct paragraphs retain optional empty paragraphs and contain between one and 256 total plain runs plus canonical fixed a:br leaves, eligible OLE XLSX payload access/replacement plus uniquely bound DOCX Office-package access/replacement, and fail-closed unsupported edits. |
| `PresentationFile.inspectPptx` | api | Inspect bounded PPTX parts, content types, the required presentation/root officeDocument relationship, namespace-aware source XML references, legacy notes/comments evidence, and Office 2021 modern author/thread/anchor semantics after raw-input, part-count, decompression, and optional compression-ratio budgets; verifyCrc32 additionally checks ZIP entry CRCs. |
| `PresentationFile.patchPptx` | api | Apply path-validated PPTX part patches, including safe slide/master/layout ID lists and slide image/chart DrawingML mutations, and atomically reject dangling package references or invalid notes/comments semantics. |
| `shape.accessibilityCapability` | api | Report sourceBound/editable/addable preflight for ordinary-shape p:cNvPr title/description/decorative metadata; export re-proves it. |
| `shape.delete` | api | Explicitly remove a source-free shape or one capability-proven imported top-level ordinary shape. Relationship-owning shapes, connector/comment/timing/extension identity graphs, nested children, and raw collection mutation fail closed; pictures, connectors, tables, and charts expose their own typed deletion capability. |
| `shape.deletionCapability` | api | Report whether one imported top-level ordinary shape is inside the bounded element-deletion profile, with a package-local native ID used for post-write absence proof. Export recomputes the capability from source bytes. |
| `shape.setAccessibilityMetadata` | api | Transactionally add, change, or clear non-visible ordinary-shape title/description/decorative metadata. Imported irregular p:cNvPr graphs fail closed. |
| `shape.text.set` | api | Set plain or structured text with ordered text, field, and line-break inlines; bounded run formatting; character, picture-bullet, or auto-numbered lists; levels, indents, spacing; and external URI, internal-slide, relative-action, or existing custom-show hyperlinks. Missing, opaque, malformed, relationship-bearing, or dangling custom-show targets and unmodeled text graphs fail closed in canonical PPTX export. |
| `shape.useBackgroundFill` | api | Read the presence-aware imported PresentationML p:sp useBgFill flag. It affects preview paint but remains source-bound and read-only; source-free authoring or wire mutation fails closed. |
| `slide.addNotes` | api | Set speaker notes as text or relationship-free paragraph/run data for inspect, preview, and canonical PPTX output. OfficeKit authors source-free notes, preserves the legacy text-only edit path, and edits a fixed imported rich paragraph/run topology; fields, hyperlinks, picture bullets, notes-body list styles/layout, and unsafe NotesMaster graphs remain source-bound and fail closed. |
| `slide.animations.add` | api | Add one bounded native object animation for fade, wipe, fly, zoom, or pulse. Use withPrevious, afterPrevious, or onClick to express speaking order; textBuild reveals whole text or paragraphs, and chartBuild reveals chart content by all-at-once, series, category, series-element, or category-element. The typed surface writes canonical PowerPoint timing and never accepts raw XML. |
| `slide.animations.remove` | api | Remove one animation issued by slide.animations or identified by its stable animation ID. Imported timing must be capability-editable; opaque timing is preserved and rejects mutation. |
| `slide.applyLayout` | api | Bind a slide to a bounded source-free layout and materialize its effective direct-frame placeholder shapes. Applying the same layout is idempotent; switching a materialized layout fails closed. The resulting p:ph identities and direct frames export natively; imported Layout relationships remain preservation-only. |
| `slide.autoLayout` | api | Place existing shapes inside a frame using horizontal or vertical flow, gap, padding, and alignment options. |
| `slide.charts.add` | api | Add a source-free literal bar, line, pie, standard area, fixed 50%-hole doughnut, marker-only scatter, bounded 2D bubble, or clustered bar+line combo chart. Category families use shared literal categories; scatter and bubble use aligned per-series numeric X/Y values, with positive area-based bubble sizes. Bar and line series, including combo members, accept up to 16 bounded native linear, exponential, logarithmic, power, polynomial, or moving-average trendlines plus one fixed/percentage/standard-deviation/standard-error/custom-literal errorBars projection. Imported trendline count and error-bar presence are fixed; unsupported labels/extensions/unknown children/complex lines remain source-owned. Supported variants retain title, legend, bounded axes, basic series styling, chart-level data labels, layout JSON, error-bar-aware SVG preview, and native ChartPart output across import/edit/re-export. Formula-backed custom error bars without an explicit embedded-workbook route, other formula/external data, advanced family geometry, topology changes, and unsupported styling fail closed rather than being flattened. |
| `slide.clearBackground` | api | Remove the direct slide background so preview and PPTX output inherit from the preserved Layout/Master chain. Unsupported imported background graphs fail closed rather than being flattened or discarded. |
| `slide.clearBackgroundImage` | api | Remove the image previously authored by slide.setBackgroundImage without changing the slide's solid/theme background. |
| `slide.clearMorph` | api | Clear a source-free or capability-approved Morph transition. Imported unknown Morph extensions remain preserved and reject mutation. |
| `slide.clearNativeBackgroundImage` | api | Remove the direct native p:bg image while preserving the inherited Layout/Master background and leaving any ordinary setBackgroundImage layer untouched. |
| `slide.clearTransition` | api | Remove one canonical direct imported or source-free slide transition. A transition-absent imported slide remains a no-op until an explicit capability-approved add; timing, sound, extension, and opaque-effect graphs remain byte-preserved and reject mutation. |
| `slide.cloneCapability` | api | Report whether an imported SlidePart can be copied as one ownership-checked OPC graph. The Codec copies every uniquely owned descendant, DataPart, and external relationship while rebinding proven shared layout, NotesMaster, image, and retained-slide targets. Sections, modern comments, outside-owned nodes, removed slide-jump targets, and over-budget graphs fail closed before the model changes. |
| `slide.comments.addThread` | api | Create either a bounded legacy PPTX annotation or an Office 2021 modern thread. A comment-free imported presentation may add canonical legacy review comments only when comments.capability.addable is true; a canonical imported legacy leaf with comments.capability.editable permits only existing root-text replacement, never addThread/replies/metadata edits. Modern mode supports a top-level element/text-range/textMatch anchor, one root, direct replies, independent people/timestamps, and active/resolved/closed state; imported modern graphs permit only fixed-topology text/status edits. |
| `slide.comments.capability` | api | Inspect defensive source-bound comment-family evidence before authoring or editing. A comment-free imported presentation may advertise legacy addability; one closed imported legacy leaf may instead advertise editable, which permits only its existing root text to change while author/time/coordinate/native identity/order/topology remain fixed. Modern graphs retain their separate fixed-topology edit contract. |
| `slide.compose` | api | Materialize a clean-room compose tree with row, column, grid, layers, box, paragraph/text, shape, table, chart, image, and rule nodes into editable slide objects. Capture the returned elements for later edits or connector targets; compose nodes remain declarative and are not Shape facades. |
| `slide.connectors.add` | api | Legacy low-level connector authoring from explicit points or target centers. Prefer slide.shapes.connect or geometry: connector when DrawingML target-plus-site identity matters. |
| `slide.continuationCapability` | api | Report full-authoring, pending-clone (export/reimport first), or bounded-overlay. Bounded overlay token-preserves the tree and allows one clean export of listed basic shapes/images. Separate SlidePart edits by reviewed revision. |
| `slide.delete` | api | Remove this slide. Source-free decks may remove any non-final slide. An imported PPTX first requires deletionCapability.supported, then removes the real SlidePart and every exclusively owned descendant (including closed notes/comments/chart/OLE/diagram/media leaves) while retaining shared parts. Inbound slide references and presentation-level custom-show/section/extension identity remain fail closed. |
| `slide.deletionCapability` | api | Report whether an imported SlidePart and its exclusively owned OPC descendant closure can be deleted. The count includes the slide plus owned OpenXml/DataPart descendants; shared layout/master/theme/media remain outside the closure. Export re-proves the graph from source bytes and aggregates all requested slide deletions into one ownership transaction. |
| `slide.duplicate` | api | Clone one original imported PPTX slide after slide.cloneCapability proves a bounded ownership graph. The JavaScript model copies the unchanged semantic element tree and resolves connector targets to fresh clone-local identities; the OfficeKit Codec then creates a distinct SlidePart, recursively byte-copies every uniquely owned OpenXmlPart and DataPart with exact local relationship IDs and external links, and rebinds only proven shared layout, NotesMaster, image, slide-jump, and other identity resources. Custom-show membership is unchanged. The pending clone cannot be edited, cloned twice, or lose its origin before export/reimport. Source-free slides, sections, modern comments, outside-owned unknown nodes, removed slide-jump targets, unresolved semantic elements/connectors, pending native payload replacements, and over-budget graphs fail closed. |
| `slide.elements` | api | Read the slide's direct cross-type scene stack from back to front. Shapes, textboxes, images, tables, charts, connectors, and groups share this order; type-specific collections remain indexes over the same elements. |
| `slide.groups.add` | api | Author recursive native DrawingML p:grpSp trees with optional non-visible group title/description/decorative metadata, outer off/ext, and local chOff/chExt coordinates. The bounded profile supports modeled shapes, connectors, images, tables, charts, and nested groups; canonical imported groups allow fixed-topology semantic edits, while group-level fills/effects, locks, transforms, unknown extensions, or unsupported descendants remain opaque and read-only. |
| `slide.hide` | api | Hide this slide from the ordinary slide show through the same source-bound p:sld/@show primitive as slide.setHidden(true). |
| `slide.images.add` | api | Add an embedded image with accessibility metadata, fit/crop, frame, rotation/flips, layout, preview, and PPTX output. Ready bounded-overlay accepts rectangular images in a clean export. OfficeKit writes native p:cNvPr, decorative metadata, and a:srcRect. |
| `slide.moveTo` | api | Move this slide to an existing 0-based deck index. On an imported PPTX, OfficeKit rewrites only the retained source SlidePart order in the presentation slide-ID list; unrelated topology changes and broad graph clones remain fail-closed. |
| `slide.placeholders.getItem` | api | Resolve a slide placeholder shape by stable ID, name, placeholder type, or numeric index. Imported placeholder.textEditable reports a verified local SlidePart text capability; identity, geometry, formatting, layout binding, and inherited Master/Layout graphs remain source-bound. |
| `slide.setBackground` | api | Set a direct slide background to a six-digit RGB/theme color solid fill or a native style reference. Recognized imported direct backgrounds are hash-bound and editable; inherited Layout/Master backgrounds remain inherited. |
| `slide.setBackgroundImage` | api | Add or replace one full-slide embedded image at the bottom of a source-free scene stack. Combine it with a translucent shape and editable foreground objects for image-led pages. Imported slides reject authored underlays because they cannot be placed beneath the complete source-bound prefix without changing native order. |
| `slide.setHidden` | api | Set whether this slide is skipped by the ordinary slide show. OfficeKit writes only p:sld/@show, uses absence for visible and show=0 for hidden, and re-proves the source-bound SlidePart before export. |
| `slide.setLayout` | api | Alias of slide.applyLayout(layout): bind and materialize a bounded source-free layout for native PPTX export. |
| `slide.setMorph` | api | Author a bounded cross-slide Morph transition between adjacent slides with real source and destination objects and unique named object pairs. OfficeKit gives both objects the same Selection Pane identity; unknown imported Morph extensions remain source-bound and are not reconstructed. |
| `slide.setNativeBackgroundImage` | api | Set a direct native p:bg/p:bgPr/a:blipFill image stretched across the slide. It stays behind all slide content and is not a reorderable or animatable scene-layer picture; use slide.setBackgroundImage when you need a movable or animated image layer. |
| `slide.setTransition` | api | Set one direct p:transition from the complete 21-effect ECMA-376 base vocabulary, with effect-specific direction/orientation/throughBlack/spokes plus speed, Office 2010+ durationMs, and click/timer advancement. Source-free slides may author it; imported slides may replace one canonical existing direct transition or add one only when transition.capability.addable is true. Timing, sound, Office-extension effects, non-integer-unit duration, and irregular source graphs fail closed. |
| `slide.shapes.add` | api | Add a shape/textbox, free-positioned p:sp line, or exact-site p:cxnSp connector with accessibility metadata. Ready bounded-overlay accepts only textbox/rect/roundRect/ellipse in a clean export. Lines support dash/ends/cap/join; custom geometry supports ordered adjustment/guide formulas, XY/polar adjustment handles, and connection sites. Only a connector retains target-plus-site identity. |
| `slide.shapes.connect` | api | Connect two modeled shapes in the same slide/group tree by preset side or exact DrawingML connection-site index. Custom shapes require an explicit index into customConnectionSites. `head` is the from/start end and `tail` is the to/end end; use tail for a forward arrow, and bringToFront() when a background shape would hide the route. The target-plus-site pair survives import, edit, clone, and second import; moved or re-parameterized modeled targets reroute before render/export. |
| `slide.shapes.getConnectionSiteIndex` | api | Resolve top/left/bottom/right to a stable bounded preset connection-site index for rect, roundRect, textbox, or ellipse. Custom shapes expose an ordered site table but require its explicit numeric index; other geometries fail closed. |
| `slide.show` | api | Show this slide in the ordinary slide show by clearing the source-bound p:sld/@show leaf through slide.setHidden(false). |
| `slide.speakerNotes.capability` | api | Return defensive sourceBound, partPresent, editable, and addable evidence. addable identifies an imported notes-absent slide whose source NotesMaster/SlideMaster Theme graph can safely receive a canonical NotesSlide. Export independently re-proves the package graph, so mutating model or wire data cannot grant authority. |
| `slide.tables.add` | api | Add an inspectable table facade with rows, columns, values, cells, rectangular merges, layout JSON, SVG preview, and canonical OfficeKit plain-text PPTX output. |
| `slide.visibilityCapability` | api | Report whether the imported p:sld/@show state is known and editable. OfficeKit exposes the inverse Agent-facing hidden boolean; invalid native lexical values stay source-owned and fail closed. |
| `slideCommentThread.addReply` | api | Append a direct reply to a source-free Office 2021 modern comment thread. Imported reply topology is fixed: existing reply text/status may change, but adding or removing replies fails closed. |
| `slideCommentThread.reopen` | api | Set the modern root comment status back to active while preserving fixed imported identity, anchor, position, and reply topology. |
| `slideCommentThread.resolve` | api | Set the modern root comment status to resolved. Imported export re-proves author/date/anchor/position/topology and source-part hashes before changing only status. |
| `table.accessibilityCapability` | api | Report sourceBound/editable/addable preflight for table graphic-frame p:cNvPr title/description/decorative metadata; export re-proves it. |
| `table.delete` | api | Explicitly remove a source-free table or one capability-proven imported direct table p:graphicFrame. Relationship-bearing, irregular, nested, or identity-sensitive frames and raw collection mutation fail closed. |
| `table.deletionCapability` | api | Report whether one imported top-level bounded relationship-free DrawingML table can be deleted, with a package-local native ID used for post-write absence proof. Export recomputes the source-bound capability. |
| `table.merge` | api | Merge one inclusive rectangular table range, retain the upper-left value, clear and lock covered cells, and emit canonical DrawingML merge topology. |
| `table.setAccessibilityMetadata` | api | Transactionally add, change, or clear non-visible table title/description/decorative metadata. Imported irregular graphic-frame p:cNvPr graphs fail closed. |

### presentation details

#### `chart.accessibilityCapability`

Report sourceBound/editable/addable preflight for chart graphic-frame p:cNvPr title/description/decorative metadata; export re-proves it.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, editable, addable } preflight; export revalidates the chart graphic-frame p:cNvPr.

#### `chart.delete`

Explicitly remove a source-free chart or one capability-proven imported chart frame. The source-bound transaction removes its exact relationship and only ChartPart descendants without another package parent; external, repeated, nested, or identity-sensitive graphs fail closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `chart` (ChartElement) — The removed ChartElement facade. Imported deletion requires chart.deletionCapability.supported and records explicit intent; export removes the exact p:graphicFrame and SlidePart relationship, garbage-collects only ChartPart descendants without outside package parents, preserves shared ChartParts, validates native-ID absence, and rejects external/repeated/nested/identity-sensitive graphs or direct array splicing.

#### `chart.deletionCapability`

Report whether one imported top-level chart frame owns one uniquely used internal ChartPart relationship. Export re-proves relationship use and the exclusively owned descendant closure; shared ChartParts survive.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, known, supported, blockedReason, nativeId } preflight. nativeId is package-local p:cNvPr evidence. Export ignores caller claims and re-proves one direct chart p:graphicFrame, one uniquely used internal ChartPart relationship, the descendant ownership closure, and absence of connector/comment/timing/extension identity consumers.

#### `chart.setAccessibilityMetadata`

Transactionally add, change, or clear non-visible chart title/description/decorative metadata independently of its visible chart title. Imported irregular graphic-frame p:cNvPr graphs fail closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `update` (object) required — { title?, description?, decorative? }; null clears a field, strings require 1-1,024 XML-safe characters, decorative requires a boolean, and a classification change plus its text clears/additions must be one transaction.

**Schema returns:**

- `chart` (ChartElement) — Same chart. Source-free and canonical imported metadata is editable; unsupported graphic-frame p:cNvPr profiles fail closed without disabling unrelated supported chart edits.

#### `compose.box`

Create a materialized shape surface with optional children inset by padding; use a named box as a stable connector or edit target.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) — Stable materialized ID.
- `name` (string) — Stable materialized name.
- `geometry` (string) — Native geometry such as rect or roundRect.
- `fill` (string) — Solid fill token or color.
- `line` (object) — Optional line style; defaults to no visible line.
- `padding` (number|object) — Inset applied to child nodes.
- `children` (object[]) — Optional child nodes materialized inside the box.

**Schema returns:**

- `node` (object) — Materialized box surface and optional child container.

#### `compose.chart`

Create a materialized chart in a resolved compose frame; encode quantitative claims as data relationships rather than decorative labels.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) — Stable materialized ID.
- `name` (string) — Stable materialized name.
- `chartType` (string) required — Supported chart type such as bar, line, or pie.
- `categories` (string[]) — Category labels.
- `series` (object[]) — Named data series and styles.

**Schema returns:**

- `node` (object) — Materialized chart node. Use a chart for quantitative relationships rather than a decorative card list.

#### `compose.column`

Create a vertical compose container. Use width/height fill, hug, or fixed pixels; gap and padding are in pixels.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) — Optional stable materialized ID for a filled container surface.
- `name` (string) — Optional stable name; filled surfaces receive a `-surface` suffix.
- `children` (object[]) — Ordered child compose nodes.
- `width` (string|number) — fill, hug, or fixed pixel width.
- `height` (string|number) — fill, hug, or fixed pixel height.
- `gap` (number) — Child gap in pixels.
- `padding` (number|object) — Container padding.
- `fill` (string) — Optional solid surface fill; the container materializes a background shape behind its children.
- `geometry` (string) — Optional surface geometry when fill is set; defaults to rect.

**Schema returns:**

- `node` (object) — Vertical compose node. A declared fill is exported as a background surface behind the children.

#### `compose.grid`

Create a grid compose container with bounded row/column tracks, spans, gaps, padding, and an optional surface.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) — Optional stable materialized ID for a filled container surface.
- `name` (string) — Optional stable name; filled surfaces receive a `-surface` suffix.
- `children` (object[]) — Grid child nodes; each may set row, column, rowSpan, or columnSpan.
- `columns` (object[]|number[]) — Bounded fixed/fr column tracks.
- `rows` (object[]|number[]) — Bounded fixed/fr row tracks.
- `gap` (number) — Shared row and column gap in pixels.
- `padding` (number|object) — Container padding.
- `fill` (string) — Optional solid surface fill behind children.

**Schema returns:**

- `node` (object) — Grid compose node with bounded tracks and spans.

#### `compose.image`

Create a materialized image node with frame, fit/crop, alt text, and an explicit user or template asset; a prompt creates only a marked placeholder.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) — Stable materialized ID.
- `name` (string) — Stable materialized name.
- `dataUrl` (string) — Embedded PNG/JPEG/GIF/SVG data URL.
- `uri` (string) — Explicit local or approved asset URI.
- `fit` (string) — stretch, contain, cover, or crop semantics.
- `alt` (string) — Accessibility description.
- `prompt` (string) — Creates a marked placeholder only; it is not a generation tool.

**Schema returns:**

- `node` (object) — Materialized image node with explicit asset and accessibility boundary.

#### `compose.layers`

Create a layered compose container whose children share the inner frame; use it for overlays and explicit z-order.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) — Optional stable materialized ID for a filled container surface.
- `name` (string) — Optional stable name; filled surfaces receive a `-surface` suffix.
- `children` (object[]) — Ordered overlay nodes; later materialized children are foreground.
- `padding` (number|object) — Container padding.
- `fill` (string) — Optional solid surface fill behind children.

**Schema returns:**

- `node` (object) — Layered compose node. Use explicit child names and returned elements when later operations need identity.

#### `compose.paragraph`

Create an editable text block with name, className/style text tokens, and stable inspect output.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `text` (string) required — Editable paragraph text.
- `name` (string) — Stable element name.
- `className` (string) — Text style token string.
- `style` (object) — Explicit text style metadata.

**Schema returns:**

- `node` (object) — Paragraph compose node.

#### `compose.row`

Create a horizontal compose container. Use fixed, hug, or fill child widths with an explicit gap and optional surface.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) — Optional stable materialized ID for a filled container surface.
- `name` (string) — Optional stable name; filled surfaces receive a `-surface` suffix.
- `children` (object[]) — Ordered child compose nodes.
- `width` (string|number) — fill, hug, or fixed pixel width.
- `height` (string|number) — fill, hug, or fixed pixel height.
- `gap` (number) — Child gap in pixels.
- `padding` (number|object) — Container padding.
- `fill` (string) — Optional solid surface fill behind children.

**Schema returns:**

- `node` (object) — Horizontal compose node. Capture the materialized elements returned by slide.compose for later edits or connector targets.

#### `compose.rule`

Create a thin horizontal or vertical rule as a materialized shape, using the resolved frame and stroke token.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `stroke` (string) — Rule color.
- `weight` (number) — Rule thickness in pixels.

**Schema returns:**

- `node` (object) — Materialized horizontal or vertical rule.

#### `compose.shape`

Create a materialized native shape, including text-bearing shapes and straight connectors, from a declarative compose node.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) — Stable materialized ID.
- `name` (string) — Stable materialized name.
- `geometry` (string) required — Native geometry or `straightConnector1`.
- `fill` (string) — Solid fill token or color.
- `line` (object) — Optional line and arrow style.
- `text` (string) — Optional text for a text-bearing shape.
- `children` (object[]) — Optional rich-text children.

**Schema returns:**

- `node` (object) — Materialized native shape. For connectors between shapes, use slide.shapes.connect with materialized return values.

#### `compose.table`

Create a materialized table in a resolved compose frame; keep the table data and column/row budget explicit.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) — Stable materialized ID.
- `name` (string) — Stable materialized name.
- `rows` (object[]|string[][]) — Bounded table row data.
- `columns` (object[]|string[]) — Optional column definitions.

**Schema returns:**

- `node` (object) — Materialized table node; keep row count and text budget within the resolved frame.

#### `compose.text`

Create the same editable paragraph node through the reference-template-compatible children-first text(children, props) helper.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `children` (string|string[]|object[]) required — Text or run-like children passed as the first argument.
- `props` (object) — Paragraph name, className, style, sizing, and placement metadata passed as the second argument.

**Schema returns:**

- `node` (object) — Reference-template-compatible alias that returns the same paragraph compose node.

#### `connector.accessibilityCapability`

Report sourceBound/editable/addable preflight for connector p:cNvPr title/description/decorative metadata; export re-proves it.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, editable, addable } preflight; export revalidates the connector p:nvCxnSpPr/p:cNvPr.

#### `connector.bringToFront`

Move a connector above peers in its slide/group scene stack. An imported direct connector may move only when its fresh zOrderCapability is editable; unsupported or nested native topology rejects.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/layered-composition.md#public-surface

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `connector` (ConnectorElement) — Move the connector to the front of its owner scene stack. Imported direct connectors require fresh editable zOrderCapability evidence.

#### `connector.delete`

Explicitly remove a source-free connector or one capability-proven imported direct p:cxnSp. Relationship-bearing or nested connectors and connector/comment/timing/extension identity consumers fail closed; endpoint shapes remain untouched.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `connector` (ConnectorElement) — The removed connector facade. Source-free deletion checks current comment/connector references. Imported deletion requires connector.deletionCapability.supported and records explicit intent; export removes only the direct p:cxnSp, validates native-ID absence, leaves its endpoint shapes unchanged, and rejects relationships, nested owners, identity-sensitive graphs, or direct array splicing.

#### `connector.deletionCapability`

Report whether one imported top-level canonical relationship-free connector can be deleted, with a package-local native ID used for post-write absence proof. Export recomputes the source-bound capability.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, known, supported, blockedReason, nativeId } preflight. nativeId is package-local p:cNvPr evidence. Export ignores caller claims and re-proves one direct relationship-free p:cxnSp, a unique native ID, and absence of connector/comment/timing/extension identity consumers.

#### `connector.sendToBack`

Move a connector behind peers in its slide/group scene stack. New shape-connected connectors start behind their nodes; an imported direct connector requires an editable zOrderCapability.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/layered-composition.md#public-surface

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `connector` (ConnectorElement) — Move the connector to the back of its owner scene stack. Imported direct connectors require fresh editable zOrderCapability evidence.

#### `connector.setAccessibilityMetadata`

Transactionally add, change, or clear non-visible connector title/description/decorative metadata. Imported irregular p:cNvPr graphs fail closed without disabling unrelated supported edits.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `update` (object) required — { title?, description?, decorative? }; null clears a field, strings require 1-1,024 XML-safe characters, decorative requires a boolean, and a classification change plus its text clears/additions must be one transaction.

**Schema returns:**

- `connector` (ConnectorElement) — Same ConnectorElement. Source-free and canonical imported metadata is editable; unsupported connector p:cNvPr profiles fail closed without disabling unrelated supported connector edits.

#### `connector.setConnectorFrom`

Atomically bind a connector start to a modeled same-tree shape and explicit connection-site index.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `target` (Shape|string) required — Modeled same-tree start shape.
- `index` (number) required — Unsigned connection-site index valid for that shape's modeled site table.

**Schema returns:**

- `connector` (ConnectorElement) — The same connector with its start target and site changed atomically.

#### `connector.setConnectorTo`

Atomically bind a connector end to a modeled same-tree shape and explicit connection-site index.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `target` (Shape|string) required — Modeled same-tree end shape.
- `index` (number) required — Unsigned connection-site index valid for that shape's modeled site table.

**Schema returns:**

- `connector` (ConnectorElement) — The same connector with its end target and site changed atomically.

#### `element.bringToFront`

Move a shape, image, table, chart, connector, or group to the front of its current slide/group scene stack. Imported direct elements require a current editable zOrderCapability; unsupported native topology fails closed.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/layered-composition.md#public-surface

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `element` (object) — The same element at the front of its owner stack. Source-bound moves require editable capability evidence.

#### `element.moveAfter`

Move one presentation element immediately in front of a different peer in the same scene stack, subject to the same source-bound capability and source-prefix checks.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/layered-composition.md#public-surface

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `target` (object) required — A different element in the same direct slide or group scene stack.

**Schema returns:**

- `element` (object) — The same element immediately in front of target.

#### `element.moveBefore`

Move one presentation element immediately behind a different peer in the same scene stack, subject to the same source-bound capability and source-prefix checks.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/layered-composition.md#public-surface

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `target` (object) required — A different element in the same direct slide or group scene stack.

**Schema returns:**

- `element` (object) — The same element immediately behind target.

#### `element.sendToBack`

Move a shape, image, table, chart, connector, or group to the back of its current slide/group scene stack. Imported direct elements require a current editable zOrderCapability; authored overlays on an imported slide cannot move below the complete source-bound prefix.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/layered-composition.md#public-surface

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `element` (object) — The same element at the back of its owner stack. Authored overlays on imported slides remain above the complete source-bound prefix.

#### `element.stackIndex`

Return an element's current zero-based position in its owning slide or group scene stack, where zero is farthest back.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/layered-composition.md#public-surface

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `stackIndex` (number) — Current zero-based position in the owner scene stack; zero is farthest back.

#### `element.zOrderCapability`

Return fresh { sourceBound, known, editable, blockedReason } evidence for moving an element in its owner scene stack. Imported direct elements are editable only when the codec issued and export can re-prove the capability.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/layered-composition.md#public-surface

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, known, editable, blockedReason } evidence. Export independently re-proves imported direct-element order against the exact source revision.

#### `exportPptxWithOfficeKit`

Export bounded direct slide backgrounds, textbox/rectangle/roundRect/ellipse shapes, free-positioned p:sp lines with the shared six-style/line-end/cap/join outline profile, rich text and lists, basic fills/lines/shadows, straight/elbow/curved p:cxnSp connectors with target connection sites through that same line profile, embedded pictures with native crop/contain/cover semantics, fixed-grid plain-text tables, recursive native p:grpSp trees, relationship-free rich speaker notes, legacy annotations, Office 2021 modern root/direct-reply threads, source-free bar/line/pie charts, the bounded literal clustered bar+line combo profile with either shared primary axes or a canonical secondary line pair, validated payload-only replacement for eligible imported OLE XLSX workbooks plus the uniquely bound DOCX Office-package profile, and bounded source-bound text updates for canonical SmartArt document nodes. Recognized imported modern threads allow only existing text/status edits; their identity, author/date metadata, anchor/range, position, topology, relationships, and source hashes remain fixed. Inherited or complex graphs remain preserved and fail closed on unsupported mutation.

**Adoption tier:** `compatibility`

**Use when:**

- A package-level or legacy interoperability operation is explicitly required.
- The caller can provide source-bound evidence and perform a second import.

**Avoid when:**

- Do not use as the default authoring route; use the typed Presentation facade first.
- Do not infer that an opaque or unsupported object became editable.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- Re-import the output and compare package/source evidence.
- Report unsupported or preserved content explicitly.

**Recipes:**

- skills/presentations/skills/presentations/tasks/review-deliver.md#evidence

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `presentation` (Presentation) required — Presentation facade within the bounded direct-slide-background/shape/free-positioned-line/rich-text/picture/fixed-table/connector/recursive-group/plain-text-notes/legacy-comment/Office-2021-modern-comment and literal native-chart boundary. Charts cover bar, line, pie, standard area, fixed 50%-hole doughnut, marker-only scatter, 2D bubble, and clustered bar+line combo. A combo supports only primary bars plus all-primary lines or all-secondary lines with the canonical top/right axis pair; formula/external data, advanced plots, irregular combos, and other imported package graphs must remain unchanged.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — PPTX bytes produced by the bundled Open XML SDK NativeAOT codec, including bounded embedded-picture, fixed-grid plain-text-table, and recursive native-group profiles, with codec diagnostics in metadata.

#### `group.accessibilityCapability`

Report sourceBound/editable/addable preflight for group-frame p:cNvPr title/description/decorative metadata; export re-proves it.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, editable, addable } preflight; export revalidates the group p:nvGrpSpPr/p:cNvPr.

#### `group.delete`

Delete one source-free or capability-proven imported group as a complete recursive ownership tree. Shared media and ChartParts survive; nested groups, outside connector/comment targets, relationship reuse, identity-sensitive graphs, and raw collection mutation fail closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `group` (GroupShape) — The removed GroupShape facade. Source-free deletion checks connector and comment references to every descendant. Imported deletion requires group.deletionCapability.supported and records one recursive intent; export removes the complete p:grpSp subtree plus owned relationship edges, garbage-collects only package descendants without outside parents, preserves shared media/ChartParts, validates every native descendant ID is absent, and rejects nested, externally referenced, ambiguously shared, identity-sensitive, or raw-array deletion.

#### `group.deletionCapability`

Report whether one imported top-level canonical recursive p:grpSp exclusively owns its complete native-ID, relationship-reference, and multi-root OPC graph. Export recomputes the source-bound capability.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, known, supported, blockedReason, nativeId } preflight. Export ignores caller claims and re-proves one direct p:grpSp, unique native IDs for its complete descendant tree, absence of outside connector/comment/timing/extension identity consumers, exclusive relationship use, and the multi-root OPC ownership closure.

#### `group.setAccessibilityMetadata`

Transactionally add, change, or clear non-visible group-frame title/description/decorative metadata. Imported irregular p:cNvPr graphs fail closed without disabling unrelated supported edits.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `update` (object) required — { title?, description?, decorative? }; null clears a field, strings require 1-1,024 XML-safe characters, decorative requires a boolean, and a classification change plus its text clears/additions must be one transaction.

**Schema returns:**

- `group` (GroupShape) — Same GroupShape. Source-free and canonical imported metadata is editable; unsupported group p:cNvPr profiles fail closed without disabling unrelated supported fixed-topology group edits.

#### `image.accessibilityCapability`

Report sourceBound/editable/addable preflight for picture p:cNvPr title/description/decorative metadata; export re-proves the residual-protected picture profile.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, editable, addable } preflight for picture title/description/decorative metadata.

#### `image.delete`

Explicitly remove a source-free image or one capability-proven imported top-level embedded picture. The source-bound transaction removes the p:pic subtree and exact relationship, garbage-collects only exclusively owned media descendants, preserves shared media, and rejects external/ambiguous/identity-sensitive graphs or raw array mutation.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `image` (ImageElement) — The removed ImageElement facade. Source-free deletion checks current comment/connector references. Imported deletion requires image.deletionCapability.supported and records explicit intent; export removes the exact p:pic and SlidePart relationship, deletes only the media closure without outside parents, preserves shared media, validates absence by native ID, and rejects direct array splicing.

#### `image.deletionCapability`

Report whether one imported top-level embedded picture can be deleted with its exact SlidePart relationship and exclusively owned media closure. Shared media survives; export re-proves relationship use, native identity, comments, connectors, timing, and extensions from source bytes.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, known, supported, blockedReason, nativeId } preflight. nativeId is package-local p:cNvPr evidence. Export ignores caller claims and re-proves one direct p:pic, one uniquely used embedded-image relationship, media-part ownership, and absence of connector/comment/timing/extension identity consumers.

#### `image.editSvgLeaf`

Replace one issued SVG RGB, opacity, or transform scalar after expectedHash verification. The exact token splice preserves all other SVG bytes and rejects stale, cross-image, invalid, unsupported, and no-op edits.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `leafId` (string) required — Exact SVG leaf ID returned by getSvgEditLeaves() for the current image revision.
- `update` (object) required — Exactly { expectedHash, value }. RGB accepts #RGB/#RRGGBB, opacity accepts 0..1, and transform leaves accept only the bounded scalar component already issued for one translate, scale, or rotate attribute.

**Schema returns:**

- `image` (ImageElement) — Same image after one exact SVG token replacement. Reinspect before the next edit; stale hashes, cross-image IDs, unsafe SVG, stylesheets/classes, unsupported transform topology, invalid bounds, and no-op values fail closed.

#### `image.editSvgText`

Replace one issued direct SVG text/tspan leaf after expectedHash verification with an escaped value. The bounded image-byte transaction preserves the rest of the SVG, rejects active/external content and stale/no-op edits, and remains verifiable after PPTX export/reimport.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `nodeId` (string) required — Exact image-local SVG leaf ID returned by getSvgTextNodes().
- `update` (object) required — Exactly { expectedHash, value }; value is escaped as SVG text and is bounded to 32767 characters without XML controls.

**Schema returns:**

- `image` (ImageElement) — Same image after one image-byte-bound direct SVG text/tspan replacement. The expected hash must match the current image bytes; active/external SVG, stale, missing, and no-op edits fail closed.

#### `image.getSvgEditLeaves`

Return defensive source-issued SVG style and transform leaves for an image. Each leaf identifies its typed value and exact expectedHash without exposing XML selectors or arbitrary attributes.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `SVG edit leaves` (object[]) — Defensive copies of the currently issued typed SVG leaves. Reinspect after every image replacement or edit; leaf IDs are capabilities, not raw XML, XPath, part paths, CSS selectors, or arbitrary attribute access.

#### `image.getSvgTextNodes`

Return defensive source-issued SVG text/tspan leaves for an image. Each leaf has a stable image-local ID, text, tag, and expectedHash; the returned records cannot mutate the image.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `image text nodes` (object[]) — Defensive copies of the currently issued direct SVG text/tspan leaves. Reinspect after every image replacement or edit; this is not a raw XML selector API.

#### `image.setAccessibilityMetadata`

Transactionally add, change, or clear a picture's non-visible title/description/decorative metadata. The legacy image.alt property reads and writes the same description state rather than creating a second metadata source.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `update` (object) required — { title?, description?, decorative? }; null clears a field, strings require 1-1,024 XML-safe characters, decorative requires a boolean, and a classification change plus its text clears/additions must be one transaction.

**Schema returns:**

- `image` (ImageElement) — Same image. The legacy alt property is the description alias; residual-protected unknown cNvPr children are preserved across a bounded metadata edit.

#### `image.svgEditCapability`

Report source-revision-bound direct SVG fill, stroke, opacity, and single transform-scalar leaves for a base64 SVG image. Each issued leaf carries an exact replacement hash; active content, external references, stylesheets, classes, and unsupported transform topology remain blocked.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { supported, reason, sourceSha256, sourceRevisionSha256?, leaves[] } evidence. Leaves are typed as svgFillRgb, svgStrokeRgb, svgOpacity, or svgTransformScalar and are bound to the current image bytes, owning image, slide, presentation, and imported package revision.

#### `image.svgTextCapability`

Report bounded direct SVG text/tspan leaves for a base64 SVG image, including the image-byte SHA-256 and exact replacement hashes. Active content, external references, oversized SVGs, and nested/non-text leaves remain unsupported.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { supported, reason, sourceSha256, nodes[] } evidence for a bounded base64 SVG image. Nodes are direct text/tspan leaves with image-local IDs and expectedHash values.

#### `importPptxWithOfficeKit`

Import PPTX bytes with editable bounded direct slide backgrounds, shapes, free-positioned p:sp lines including bounded line ends/caps/joins, rich text, recognized owner-local SlidePart placeholder text, rectangular pictures and native source rectangles, tables, target-bound p:cxnSp connectors, recursive canonical p:grpSp groups, bar/line/pie charts, the canonical literal clustered bar+line combo profile with either shared primary axes or a secondary line pair, legacy text-only speaker notes plus fixed-topology relationship-free rich notes and a re-proven addable capability for eligible notes-absent slides, unchanged-only legacy comments, fixed-topology modern comment text/status edits, defensive payload access for eligible OLE XLSX workbooks plus one uniquely bound DOCX Office-package profile, and a source-bound SmartArt text capability only for a canonical closed four-part DiagramDataPart whose nodes use fixed direct paragraphs with optional empty paragraphs, between one and 256 total direct plain runs, and canonical fixed a:br leaves. Compound/theme/custom-dash/effect/extension outlines and all other unsupported content remain source-bound and read-only rather than being flattened.

**Adoption tier:** `compatibility`

**Use when:**

- A package-level or legacy interoperability operation is explicitly required.
- The caller can provide source-bound evidence and perform a second import.

**Avoid when:**

- Do not use as the default authoring route; use the typed Presentation facade first.
- Do not infer that an opaque or unsupported object became editable.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- Re-import the output and compare package/source evidence.
- Report unsupported or preserved content explicitly.

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `input` (FileBlob|Uint8Array|ArrayBuffer) required — PPTX package bytes.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `presentation` (Presentation) — Imported presentation facade with editable bounded direct slide backgrounds, shapes, free-positioned p:sp lines with bounded line ends/caps/joins, rich text, recognized owner-local SlidePart placeholder text, pictures, tables, target-bound p:cxnSp connectors, recursive canonical groups, literal bar/line/pie/standard-area/fixed-doughnut/marker-scatter/2D-bubble charts, and the clustered bar+line combo profile with either shared primary axes or a secondary line pair. Formula/external data, compound/theme/custom-dash/effect/extension line outlines, and advanced plot topology remain source-bound. Placeholder identity/geometry/formatting and inherited template graphs remain source-bound; advanced package graphs are read-only except for validated payload-only replacement on eligible XLSX OLE workbooks through getEmbeddedWorkbook/replaceEmbeddedWorkbook and one uniquely bound DOCX profile through getEmbeddedOfficePackage/replaceEmbeddedOfficePackage.

#### `nativeObject.getEmbeddedOfficePackage`

Read a defensive FileBlob copy from an eligible source-bound top-level OLE package. It is compatible with the legacy XLSX workbook profile and currently adds one uniquely bound DOCX profile; it never exposes arbitrary OLE or native-part mutation.

**Adoption tier:** `compatibility`

**Use when:**

- A package-level or legacy interoperability operation is explicitly required.
- The caller can provide source-bound evidence and perform a second import.

**Avoid when:**

- Do not use as the default authoring route; use the typed Presentation facade first.
- Do not infer that an opaque or unsupported object became editable.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- Re-import the output and compare package/source evidence.
- Report unsupported or preserved content explicitly.

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `officePackage` (FileBlob) — Defensive Office-package FileBlob copy with source part-path and SHA-256 metadata. It is compatible with the legacy eligible XLSX workbook profile and currently adds only one uniquely bound DOCX profile; arbitrary OLE payloads remain unavailable.

#### `nativeObject.getEmbeddedWorkbook`

Read a defensive FileBlob copy of the XLSX payload from an eligible source-bound top-level OLE object without exposing arbitrary native-part mutation.

**Adoption tier:** `compatibility`

**Use when:**

- A package-level or legacy interoperability operation is explicitly required.
- The caller can provide source-bound evidence and perform a second import.

**Avoid when:**

- Do not use as the default authoring route; use the typed Presentation facade first.
- Do not infer that an opaque or unsupported object became editable.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- Re-import the output and compare package/source evidence.
- Report unsupported or preserved content explicitly.

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `workbook` (FileBlob) — Defensive XLSX FileBlob copy with source part-path and SHA-256 metadata. Available only for a uniquely bound top-level OLE package relationship.

#### `nativeObject.replaceEmbeddedOfficePackage`

Replace only a source-bound Office package on an eligible imported top-level OLE object. The current generic profile validates DOCX bytes and exact content type while preserving the OLE shell, relationships, preview, and all other native parts; malformed, shared, ambiguous, or unsupported package graphs fail closed.

**Adoption tier:** `compatibility`

**Use when:**

- A package-level or legacy interoperability operation is explicitly required.
- The caller can provide source-bound evidence and perform a second import.

**Avoid when:**

- Do not use as the default authoring route; use the typed Presentation facade first.
- Do not infer that an opaque or unsupported object became editable.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- Re-import the output and compare package/source evidence.
- Report unsupported or preserved content explicitly.

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `officePackage` (FileBlob|Uint8Array|ArrayBuffer|ArrayBufferView) required — Replacement package bytes, copied defensively and limited to 16 MiB. A DOCX FileBlob must retain application/vnd.openxmlformats-officedocument.wordprocessingml.document; current generic export validates the DOCX OPC package with the Microsoft Open XML SDK.

**Schema returns:**

- `nativeObject` (NativePresentationObject) — Queues one payload-only replacement on an eligible source-bound top-level OLE object. The generic profile currently accepts only DOCX; it re-proves the original part path, relationship ID, MIME type, source digest, and exclusive inbound relationship before changing the embedded bytes. The OLE shell, relationship topology, preview image, and other native parts remain fixed; unsupported graphs or changed bindings fail closed.

#### `nativeObject.replaceEmbeddedWorkbook`

Replace only the XLSX payload of an eligible imported top-level OLE object. OfficeKit validates the new workbook and immutable source binding, preserves the OLE shell, relationships, preview, and all other native parts, and fails closed for malformed or ambiguous graphs.

**Adoption tier:** `compatibility`

**Use when:**

- A package-level or legacy interoperability operation is explicitly required.
- The caller can provide source-bound evidence and perform a second import.

**Avoid when:**

- Do not use as the default authoring route; use the typed Presentation facade first.
- Do not infer that an opaque or unsupported object became editable.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- Re-import the output and compare package/source evidence.
- Report unsupported or preserved content explicitly.

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `workbook` (FileBlob|Uint8Array|ArrayBuffer|ArrayBufferView) required — Replacement XLSX bytes, copied defensively and limited to 16 MiB before canonical export validation.

**Schema returns:**

- `nativeObject` (NativePresentationObject) — Queues one payload-only replacement on an eligible source-bound top-level OLE object. Export preserves the OLE shell, relationship topology, preview image, and other native parts; invalid XLSX or changed source bindings fail closed.

#### `nativeObject.setDiagramNodeRunText`

Replace one existing direct a:r/a:t value by zero-based source-order run index across a proven SmartArt node's fixed direct paragraphs. Empty paragraphs, paragraph/run topology, a:pPr, a:rPr, canonical fixed a:br, and a:endParaRPr stay source-bound; wholly empty nodes, fields, noncanonical breaks, topology changes, and unsupported diagrams reject without fallback.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `nodeId` (string) required — Exact existing SmartArt DiagramDataPart dgm:pt/@modelId from nativeObject.diagramText.nodes.
- `runIndex` (integer) required — Zero-based index into that node's immutable diagramText.nodes[].runs array. The index must already exist; run creation, removal, and reordering are unsupported.
- `text` (string) required — Replacement value for exactly one existing a:r/a:t, limited so the complete node remains at most 32,767 XML-safe characters.

**Schema returns:**

- `nativeObject` (NativePresentationObject) — Queues one source-ordered run-local update without changing empty paragraphs, a:pPr, a:rPr, canonical fixed a:br, a:endParaRPr, or guessing style ownership. Export re-proves the source digest, node IDs/order, fixed paragraph/run/break topology, and closed graph; only the bound DiagramDataPart may change. Wholly empty nodes, paragraph or break changes, fields, noncanonical breaks, and unsupported markup fail closed.

#### `nativeObject.setDiagramNodeText`

Replace a one-run source-bound SmartArt document node after its top-level four-part graph and fixed direct-paragraph/run DiagramDataPart profile are proven. Multi-run nodes reject so OfficeKit never guesses a formatting boundary.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `nodeId` (string) required — Exact existing SmartArt DiagramDataPart dgm:pt/@modelId from nativeObject.diagramText.nodes. Node creation, removal, ordering, and identity changes are not supported.
- `text` (string) required — Replacement plain text, limited to 32,767 XML-safe characters. Tabs, line feeds, and carriage returns are allowed; other XML control characters and invalid Unicode scalars reject.

**Schema returns:**

- `nativeObject` (NativePresentationObject) — Queues one text-only update for a one-run node in a source-bound top-level SmartArt frame. Export re-proves the closed dm/lo/qs/cs graph and fixed direct-paragraph/run/break profile, may rewrite only its bound DiagramDataPart, and preserves paragraph/node/run/break topology, formatting, frame, relationships, layout, quick-style, colors, and all non-data parts. Multi-run nodes require setDiagramNodeRunText; fields, noncanonical breaks, connected, nested, or otherwise unrecognized graphs fail closed.

#### `nativeObject.setName`

Native OLE, SmartArt/diagram, contentPart, and media objects imported through OfficeKit are source-bound and read-only for names; setName rejects instead of mutating the preserved package graph. Separate bounded SmartArt node/run text methods own the only modeled diagram mutation.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `name` (string) required — Requested native-object display name, limited to 1,024 characters. Imported native objects are read-only, so the method rejects.

**Schema returns:**

- `nativeObject` (NativePresentationObject) — No mutation is performed; imported native OLE/diagram/contentPart objects are source-bound and read-only.

#### `nativeObject.setPosition`

Native OLE, SmartArt/diagram, contentPart, and media objects imported through OfficeKit are source-bound and read-only; setPosition rejects instead of rewriting their geometry or payload graph.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `position` (object) required — Requested outer pixel frame. Imported native objects are read-only, so the method rejects.

**Schema returns:**

- `nativeObject` (NativePresentationObject) — No mutation is performed; native geometry and payload graphs remain source-bound and read-only.

#### `presentation.auditAccessibility`

Audit modeled slide objects for explicit meaningful/decorative classification and non-visible title/description coverage, while separating native-object and reading-order checks that still require manual host review. It never claims whole-deck accessibility conformance.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/review-deliver.md#evidence

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `maxChars` (number) — Maximum bounded NDJSON size across machine issues and manual-review records.

**Schema returns:**

- `report` (object) — A host-neutral report with machineCheckPassed, conformanceClaimed: false, manualReviewRequired, counts, machine issues, and separate manual checks. Unclassified modeled objects and explicit meaningful objects without title/description fail the machine check. Opaque native objects and multi-object slide reading order remain manual checks; the audit does not mutate shape-tree order or claim PowerPoint/PDF accessibility conformance.

#### `Presentation.create`

Create a deck model whose canonical OfficeKit export supports ordinary slides, the complete ECMA-376 base slide-transition vocabulary, direct solid/style-reference slide backgrounds, shapes, rich text, tables, images, connectors, recursive native p:grpSp groups, plain-text speaker notes, native custom shows with canonical run links, literal bar/line/pie/standard-area/fixed-doughnut/marker-scatter/2D-bubble charts, and a bounded literal clustered bar+line combo profile. Combo bars stay on the primary pair; all lines share either that pair or the canonical secondary top/right pair. Formula/external chart data, custom themes, Master/Layout authoring, comments, custom-show topology mutation, advanced plot geometry, mixed line groups, secondary bars, irregular combo graphs, and other package-level features remain outside the source-free PPTX boundary.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `slideSize` (object) — Slide width and height in pixels; defaults to 1280x720. On a trusted imported PPTX, changing it updates only the source-bound p:sldSz canvas and never rescales existing coordinates.
- `theme` (object) — Model theme metadata. OfficeKit 0.2 source-free export requires the default theme; imported themes are read-only.
- `master` (object) — The one canonical source-free Slide Master: name/background, bounded title/body/ctrTitle/subTitle direct-frame placeholders, and bounded textParagraphStyles. Theme overrides are unsupported.
- `masters` (object[]) — Model-level Slide Master definitions. Source-free PPTX authoring accepts exactly one master; imported master graphs remain source-bound and read-only.
- `layouts` (object[]) — Bounded source-free layouts linked to the canonical master. Each uses blank, title, titleOnly, or obj/titleAndContent plus direct-frame text placeholders; imported layouts remain source-bound and read-only.
- `commentFormat` (string) — Comment wire family: legacy (default) or modern. Modern selects the bounded Office 2021 author/comments graph; the two families cannot be mixed.

**Schema returns:**

- `presentation` (Presentation) — Editable presentation facade.

#### `presentation.customShows.add`

Define an ordered native p:custShowLst playback route for source-free OfficeKit export. Text runs may target a show by exact name with optional returnToSlide. Canonical imported shows may change only their name and ordered retained-slide membership; fixed native identity keeps existing run links bound across a rename, while irregular graphs stay opaque.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `name` (string) required — Unique custom-show name, compared case-insensitively.
- `slides` (PresentationSlide[]|string[]) required — Ordered non-empty list of slide facades or stable slide IDs from this presentation.
- `nativeId` (number) — Optional preserved unsigned 32-bit p:custShow ID; new IDs are allocated collision-free.

**Schema returns:**

- `customShow` (PresentationCustomShow) — Appended native custom show for source-free PPTX authoring. Imported additions fail closed; use name assignment and setSlides(...) only on an existing canonical show.

#### `presentation.customShows.getItem`

Resolve a source-free or canonical imported custom show by zero-based index, stable facade ID, or exact name.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `idOrNameOrIndex` (string|number) required — Stable custom-show ID, exact name, or zero-based collection index.

**Schema returns:**

- `customShow` (PresentationCustomShow|undefined) — Matching custom-show facade or undefined.

#### `presentation.designProfile`

Return a bounded read-only design-language profile for the current deck: source revision binding when imported, canvas, palette, typography, density, normalized geometry rhythm, layout families, slide archetypes, repeated visual candidates, and opaque native summaries. The profile is evidence for template-conditioned generation only; it contains no XML selectors, package paths, source bytes, or mutation authority.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `maxItems` (number) — Maximum number of bounded profile entries; defaults to 256.
- `includeComponentCandidates` (boolean) — Include source-bound candidate ID summaries when the presentation was imported; defaults to true. This is descriptive evidence, not mutation authority.

**Schema returns:**

- `profile` (object) — Deterministic office-kit/pptx-design-profile/v1 evidence. Imported profiles carry sourceBound=true and the exact revisionSha256; source-free profiles remain descriptive and have no source revision.

#### `presentation.editComponentOccurrence`

Apply one atomic batch of typed native-leaf edits to a repeated component occurrence issued by presentation.inspect({ includeComponentCandidates: true }). The occurrence editCapability and each leafId, targetId, and expectedHash are source-revision-bound; all values are validated before any leaf is changed. Only codec-issued text, color, geometry, chart, SmartArt, or other bounded leaf kinds are accepted. Raw XML, selectors, part paths, foreign leaves, duplicate leaves, stale hashes, and edits outside the selected component fail closed.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `candidateId` (string) required — Exact candidateId from a trusted componentCandidate inspection.
- `occurrenceIndex` (number) — Zero-based occurrence index from the candidate record; defaults to 0.
- `expectedCandidate` (object) — Optional complete candidate record from the same inspection; it is compared to current source-bound evidence.
- `edits` (object[]) required — One through 256 records with targetId, leafId, expectedHash, and value; every leafId must belong to the selected occurrence editCapability.

**Schema returns:**

- `receipt` (object) — Atomic componentEdit receipt containing immutable nativeLeafEdit receipts. All issued leaves are validated before mutation, then compiled into one deterministic source-bound Edit Plan at export. This batches only bounded leaves already issued by inspection; it does not add a group schema, raw XML surface, relationship access, or topology editing.

#### `presentation.editNativeLeaf`

Change one native leaf issued by presentation.inspect({ includeNativeLeaves: true }) using its targetId, leafId, expectedHash, and a typed value. Leaf IDs are bound to the exact imported revision and target. Repeat the call for a coordinated move/resize; one export sorts all issued leaves into one deterministic Edit Plan. The current profile changes existing text leaves, including group children and shapes with source-owned outer styling, shape RGB/local-geometry scalars, picture local-geometry scalars (including opaque pictures whose payload and effects remain source-owned), direct rich chart-title runs, direct numeric bar-chart cache points proven against one exact cell in a uniquely bound embedded XLSX, direct SmartArt text runs from one canonical closed DiagramDataPart with a unique inbound owner, explicit bare text-body AutoFit choices, direct column-direction flags, direct vertical-text modes, or explicit literal paragraph/run font-size/typeface/style/color/decoration/alignment leaves (`paragraphAlignment`, `verticalAnchor`, `fontSizePoints`, `fontFamily`, `fontFamilyEastAsia`, `fontBold`, `fontItalic`, `fontColorRgb`, `fontColorScheme`, `fontUnderline`, `fontStrike`, `fontKerningPoints`) proven on one direct text run. Paragraph alignment is limited to a direct canonical `a:pPr/@algn` token (`left`, `center`, `right`, or `justify`); vertical text anchoring is limited to a direct canonical `a:bodyPr/@anchor` token (`top`, `center`, or `bottom`); text-body AutoFit is limited to a direct bare `a:noAutofit`, `a:normAutofit`, or `a:spAutoFit` child (`none`, `shrinkText`, or `resizeShape`); column direction is limited to a direct canonical `a:bodyPr/@rtlCol` token (`0` or `1`); vertical text is limited to a direct canonical `a:bodyPr/@vert` token (`horz`, `vert`, or `vert270`, exposed as `horizontal`, `vertical`, or `vertical270`); underline and strike are limited to standard direct DrawingML tokens; kerning is limited to a direct non-negative `a:rPr/@kern` token, exposed in points and spliced as hundredths of a point. Inherited, malformed, effect-bearing, or otherwise irregular style graphs remain opaque. A chartDataValue operation changes both the ChartPart cache and that worksheet cell. A diagramText operation token-splices only its issued a:t and does not reserialize the diagram part. Separate typed imported-table, embedded-image, and element-delete facades lower to tableCellText, imageAsset, and deleteElement operations in the same Edit Plan; those operation kinds are not arbitrary native-leaf selectors. The compiler binds the complete ownership tree and dependent parts. Stale hashes, concurrent non-leaf changes, foreign IDs, raw XML, XPath, part paths, arbitrary attributes or cells, relationship fields, formulas, namespaces, and topology changes reject.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Examples:**

- presentation.editNativeLeaf(leaf.targetId, leaf.leafId, { expectedHash: leaf.expectedHash, value: 'Reviewed title' })

**Schema parameters:**

- `targetId` (string) required — Exact targetId from an issued nativeLeaf record.
- `leafId` (string) required — Opaque revision-bound leafId from the same inspect result.
- `update` (object) required — Exactly { expectedHash, value }; raw XML, selectors, part paths, attributes, and topology fields are rejected.

**Schema returns:**

- `receipt` (object) — Immutable nativeLeafEdit receipt. Repeated authorized calls may compile into one deterministic source-bound Edit Plan; the Codec independently re-proves every leaf and its owning part. A picture geometry leaf changes only one direct a:off/a:ext scalar, including when the picture itself remains opaque because its payload and effects are outside the semantic image profile. A `rotationDegrees` leaf changes only one direct `a:xfrm/@rot` scalar on a supported source-bound shape or picture; the source shape-tree path, frame, flips, payload, and all other XML remain fixed. A chartTitleText leaf changes only the issued a:t token in one uniquely bound internal ChartPart. A chartDataValue leaf changes one numeric cache token and the matching direct numeric cell in its uniquely bound embedded XLSX, with a separate nested footprint. A diagramText leaf changes only one issued a:t token in a canonical closed DiagramDataPart; its node identity, run topology, relationships, layout, quick-style, colors, and owning graphicFrame remain fixed. A verticalAnchor leaf changes only a direct canonical a:bodyPr/@anchor token (top, center, or bottom). The textBodyInsetLeftEmu, textBodyInsetTopEmu, textBodyInsetRightEmu, and textBodyInsetBottomEmu leaves change only their corresponding direct non-negative a:bodyPr/@lIns, @tIns, @rIns, or @bIns EMU token. A textBodyWrap leaf changes only a direct canonical a:bodyPr/@wrap token (`square` or `none`). A textBodyColumnCount leaf changes only a direct canonical a:bodyPr/@numCol token (1 through 16). A textBodyAutoFit leaf changes only the local name of one direct bare a:noAutofit, a:normAutofit, or a:spAutoFit child (`none`, `shrinkText`, or `resizeShape`). A textBodyColumnDirection leaf changes only a direct canonical a:bodyPr/@rtlCol token (`0` or `1`). A textBodyVerticalText leaf changes only a direct canonical a:bodyPr/@vert token (`horz`, `vert`, or `vert270`, exposed as `horizontal`, `vertical`, or `vertical270`). These text-body leaves leave the rest of the text body and package untouched. A fontSizePoints, fontFamily, fontFamilyEastAsia, fontBold, fontItalic, fontColorRgb, fontColorScheme, fontUnderline, fontStrike, or fontKerningPoints leaf changes only one direct a:rPr scalar (`sz`, `a:latin/@typeface`, `a:ea/@typeface`, `b`, `i`, `a:solidFill/a:srgbClr/@val` or `a:solidFill/a:schemeClr/@val`, `u`, `strike`, or `kern`) and leaves the rest of the run and package untouched; underline and strike accept only standard direct tokens, kerning accepts a direct non-negative hundredths-of-a-point token exposed in points, while inherited, transformed, effect-bearing, and irregular font graphs remain blocked. Typed imported table-cell text, same-format embedded-image replacement, and capability-proven element deletion are compiled separately as tableCellText, imageAsset, and deleteElement operations with their own asset or deletion proof; they are not accepted through this native-leaf call. Chart identity, relationships, formulas, and plot topology likewise remain fixed. Concurrent changes outside issued leaves reject.

**Returns:**

immutable nativeLeafEdit receipt

**Notes:**

- A fillOpacityThousandthPercent leaf is available for a direct solid RGB fill with one bounded alpha token. Pass a 0..1 fraction; only that alpha token changes and irregular or effect-bearing fills remain blocked.
- An imported connector or group descendant may expose lineStyle, lineCap, lineJoin, lineStartArrow, or lineEndArrow leaves when an existing prstDash/cap/join/endpoint token and simple solid outline are proven. Use solid, dashed, dotted, dash-dot, dash-dot-dot, flat, round, square, bevel, miter, none, triangle, stealth, diamond, oval, or arrow; only the selected token changes, while custom dash/effect graphs, miter limits, endpoint width/length changes, and other irregular line graphs stay opaque.

#### `presentation.export`

Export a slide SVG preview, deck SVG montage via { format: 'montage' }, or target/search-sliced layout JSON.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/review-deliver.md#evidence

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `format` (string) — svg by default, montage, or layout.
- `slide` (Slide) — Slide facade to export; defaults to the first slide.
- `columns` (number) — Montage column count.
- `scale` (number) — Montage thumbnail scale.
- `gap` (number) — Montage gap in pixels.

**Schema returns:**

- `blob` (FileBlob) — SVG montage/slide preview or layout JSON.

#### `presentation.fontFamilies`

Return a fresh sorted, case-insensitively deduplicated list of explicitly used presentation text and bullet font families.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `families` (string[]) — Explicit font-family inventory; theme tokens such as +mj-lt are excluded.

#### `presentation.inspect`

Emit NDJSON for deck, custom shows, PowerPoint sections, slides, cross-type layers, direct slide transitions, textboxes, shapes, grouped shapes, tables, charts, images, and native contentPart/OLE/diagram/media objects with bounded editability, relationship-reference, root-relationship, preserved-part, eligible embedded Office-package summaries, and each slide's continuationCapability; narrow with search/target anchors and shape fields with include/exclude. Layer records expose bottom-to-top stackIndex and zOrderCapability without exposing package paths. On a trusted imported source, includeNativeLeaves: true returns revision-bound safe leaves without exposing part paths or XML selectors, while includeComponentCandidates: true returns repeated visual primitives with source hashes, occurrences, and explicit reuse limits; only closed top-level candidates can issue the bounded reuseSourceComponent operation.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/review-deliver.md#evidence

**Example paths:**

- examples/create-pptx-compose.mjs

**Examples:**

- presentation.inspect({ includeNativeLeaves: true, target: shape.id })
- presentation.inspect({ includeComponentCandidates: true, kind: 'componentCandidate' })

**Options:**

- kind
- search
- target/targetId/id/anchor
- before/after/context
- include/fields
- exclude/omit
- includeNativeLeaves
- includeComponentCandidates
- maxChars

**Schema parameters:**

- `kind` (string) — Comma-separated deck/theme/layout/slide/layer/zOrder/transition/textbox/textRange/shape/groupShape/table/chart/image/connector/animation/morph/nativeObject/nativeLeaf/componentCandidate/contentPart/oleObject/diagram/comment/notes/customShow/section kinds.
- `search` (string) — Case-insensitive record filter.
- `target` (string) — Stable target ID/anchor.
- `before` (number) — Context records before matches.
- `after` (number) — Context records after matches.
- `include` (string) — Comma-separated fields to keep.
- `exclude` (string) — Comma-separated fields to omit.
- `includeNativeLeaves` (boolean) — On a trusted imported PPTX, include revision-bound safe text leaves, direct text-body inset EMU leaves (`textBodyInsetLeftEmu`, `textBodyInsetTopEmu`, `textBodyInsetRightEmu`, `textBodyInsetBottomEmu`), explicit `textBodyWrap` leaves (`square` or `none`), direct `textBodyColumnCount` leaves (`numCol` 1 through 16), bare `textBodyAutoFit` leaves (`none`, `shrinkText`, or `resizeShape`), direct `textBodyColumnDirection` leaves (`false`/`true` for canonical `rtlCol` 0/1), and direct `textBodyVerticalText` leaves (`horizontal`/`vertical`/`vertical270` for canonical `vert` `horz`/`vert`/`vert270`) for canonical a:bodyPr tokens, bounded direct `rotationDegrees`, `flipHorizontal`, and `flipVertical` leaves for supported source-bound shape/picture `a:xfrm` tokens, shape RGB/local-geometry leaves, picture local-geometry leaves (including opaque pictures with a separately proven direct frame), direct rich-title text leaves from a uniquely bound internal ChartPart, direct numeric bar-chart cache points proven against exact cells in one uniquely bound embedded XLSX, direct SmartArt text runs from one canonical closed DiagramDataPart with a unique inbound owner, and direct run `fontKerningPoints` leaves for canonical non-negative `a:rPr/@kern` tokens. Missing, inherited, no* and malformed attributes remain opaque.
- `includeComponentCandidates` (boolean) — On a trusted imported PPTX, include repeated visual primitives as source-revision-bound componentCandidate records. Ambiguous, opaque, nested, and relationship-bound graphs are blocked; only a closed top-level candidate can authorize reuseSourceComponent.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `inspection` (object) — Bounded { ndjson, truncated } inspection result.

**Returns:**

{ ndjson, truncated } bounded NDJSON records

#### `presentation.layout.clearBackground`

Clear a direct background on a bounded source-free layout. Imported-layout mutation remains source-bound and fails closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `layout` (SlideLayoutTemplate) — Clears a direct background on a bounded source-free layout. Imported-layout edits fail closed.

#### `presentation.layout.placeholders.add`

Append a direct-frame title/body/ctrTitle/subTitle text placeholder to a source-free layout. It becomes a native p:ph and must be materialized on each slide through applyLayout/setLayout; object/media/chart/table placeholders remain source-bound.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `type` (string) required — title, body, ctrTitle, or subTitle; common aliases centeredTitle and subtitle normalize to native tokens.
- `idx` (number) — Native unsigned placeholder index; index is accepted as an alias.
- `index` (number) — Alias of idx.
- `position` (object) required — Required direct pixel frame { left, top, width, height } for source-free export.
- `text` (string|string[]|object|object[]) — Optional prompt/default text using the bounded presentation text profile.
- `style` (object) — Optional bounded default run/paragraph style.

**Schema returns:**

- `placeholder` (object) — Appended source-free layout placeholder definition. Use slide.applyLayout/setLayout to materialize it on a slide.

#### `presentation.layout.placeholders.summary`

Return a defensive layout-placeholder discovery snapshot with stable IDs, names, native types/indexes, required flags, and direct-frame presence/geometry; editing the snapshot cannot mutate the model.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `summary` (object) — Fresh defensive snapshot of the layout placeholder collection. It reports ownerId, count, requiredCount, sorted types, and copied items; imported inherited placeholders explicitly report hasDirectPosition: false.

#### `presentation.layout.setBackground`

Set a direct background on a bounded source-free layout. Imported-layout mutation remains source-bound and fails closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `background` (string|object) required — Direct solid RGB/scheme background or native style reference with index.

**Schema returns:**

- `layout` (SlideLayoutTemplate) — Sets a direct background on a bounded source-free layout. Imported-layout edits fail closed.

#### `presentation.layouts.add`

Create one bounded source-free layout under the canonical master. Use blank, title, titleOnly, or obj/titleAndContent plus direct-frame text placeholders; imported layouts remain source-bound and read-only.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `name` (string) required — Layout name; passing a name string is also accepted.
- `type` (string) — Source-free type: blank, title, titleOnly, obj, or aliases object/content/titleAndContent. Imported layouts retain their native type read-only.
- `masterId` (string) — Master identity.
- `background` (string|object) — Optional layout background overriding the linked master background.
- `placeholders` (object[]) — Direct-frame title/body/ctrTitle/subTitle source-free text placeholders. Each needs type, idx/index, and position left/top/width/height; object/chart/table/media placeholders are not authored.
- `slideGuides` (object[]) — Imported layouts expose the presentation's read-only native guide definitions. Canonical export preserves them through the source-bound view-properties part.

**Schema returns:**

- `layout` (SlideLayoutTemplate) — Appended bounded source-free layout under the canonical master. Imported layout graphs remain source-bound and read-only.

#### `presentation.layouts.getById`

Resolve a layout by its stable ID without falling back to a same-named or same-typed layout.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) required — Exact stable layout ID.

**Schema returns:**

- `layout` (SlideLayoutTemplate|undefined) — Matching layout or undefined.

#### `presentation.master`

Access the one canonical source-free Slide Master. It may author a direct background, bounded text styles, and direct-frame title/body/ctrTitle/subTitle placeholders; imported Master graphs remain source-bound and read-only.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) — Stable master identity used by layouts.
- `name` (string) — Native Slide Master name.
- `background` (string|object) — Solid RGB/scheme background or native background reference with index.
- `theme` (object) — Optional model theme override. Canonical source-free export rejects master-specific theme overrides.
- `placeholders` (object[]) — Source-free direct-frame title/body/ctrTitle/subTitle text placeholders. Each requires type, idx/index, and left/top/width/height; imported placeholders remain source-bound and read-only.
- `textParagraphStyles` (object) — title/body/other level maps (0-8) using the structured paragraph style fields, including embedded or external bulletImage values.
- `slideGuides` (object[]) — Read-only imported PowerPoint guide definitions with horizontal/vertical orientation and raw native position. Source-free authoring and imported mutation are unsupported.

**Schema returns:**

- `master` (PresentationSlideMaster) — One canonical source-free Slide Master or a source-bound imported master. Source-free output supports a direct background, bounded text styles, and direct-frame textual placeholders; imported masters are read-only.

#### `presentation.master.clearBackground`

Clear the direct background of the one canonical source-free master. Imported-master mutation remains source-bound and fails closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `master` (PresentationSlideMaster) — Clears the direct background of the one canonical source-free master. Imported-master edits fail closed.

#### `presentation.master.setBackground`

Set the direct background of the one canonical source-free master. Imported-master mutation remains source-bound and fails closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `background` (string|object) required — Direct solid RGB/scheme background or native style reference with index.

**Schema returns:**

- `master` (PresentationSlideMaster) — Sets the direct background of the one canonical source-free master. Imported-master edits fail closed.

#### `presentation.master.setTheme`

Set a model-level master theme override for preview only. Canonical PPTX export rejects that source-free override; imported-master mutation remains source-bound and fails closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `theme` (object|null) required — Partial master theme override, or null to inherit presentation.theme.

**Schema returns:**

- `master` (PresentationSlideMaster) — Model-only theme override for preview; canonical export rejects source-free master-specific themes and imported-master edits.

#### `presentation.masters.add`

Append a model-level Slide Master. Source-free PPTX authoring requires exactly one master, so use Presentation.create({ master }) or presentation.master for the canonical profile; multiple masters and imported-master edits fail closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) required — Stable unique master identity used by layouts.
- `name` (string) — Native Slide Master name.
- `background` (string|object) — Solid RGB/scheme background or native background reference with index.
- `theme` (object) — Optional model theme override; source-free master-specific themes are unsupported.
- `placeholders` (object[]) — Direct-frame title/body/ctrTitle/subTitle source-free text placeholders. A second master makes source-free export fail closed.
- `textParagraphStyles` (object) — title/body/other level maps (0-8) using the structured paragraph style fields, including embedded or external bulletImage values.

**Schema returns:**

- `master` (PresentationSlideMaster) — Appended model-level Slide Master. Canonical source-free export accepts exactly one master, so adding another deliberately fails closed.

#### `presentation.masters.getItem`

Resolve a model-level or imported Slide Master by stable ID or name.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `idOrName` (string) required — Stable master ID or native master name.

**Schema returns:**

- `master` (PresentationSlideMaster|undefined) — Matching Slide Master or undefined.

#### `presentation.planTemplateGeneration`

Build a source-bound, read-only multi-page frame map from a trusted imported PPTX: choose clone-safe source slides by role, archetype, content density, and preferred visual kinds; issue bounded text-run targets and reusable-component candidates; report heuristic text-fit warnings, alternatives, opaque-object limits, and blocked requests without mutating the deck.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `slides` (object[]) required — One through 64 page requests. Each request accepts role, title/body, optional sourceSlideOrdinal or archetypeSignature, preferredKinds, and assetIntent; unknown fields are rejected.
- `maxItems` (number) — Maximum bounded asset candidates per planned page; defaults to 64.

**Schema returns:**

- `templatePlan` (object) — Deterministic office-kit/pptx-template-plan/v1 frame map. It is source-revision-bound read-only evidence: pages include clone-safe source slide locators, bounded text targets, reusable visual candidates, heuristic fit status, alternatives, and explicit rejected requests. Export/reimport must re-resolve locators; the plan never grants raw XML or mutation authority.

#### `presentation.resolve`

Map stable inspect anchor IDs back to facade objects, including custom shows, PowerPoint sections, and slide transitions; imported advanced package objects may be read-only.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) required — Stable deck, theme, layout, slide, transition, element, custom-show, section, comment, or text-range ID.

**Schema returns:**

- `object` (object|undefined) — Resolved editable facade/record or undefined.

#### `presentation.resolveComponentCandidate`

Resolve one candidateId issued by presentation.inspect({ includeComponentCandidates: true }) to a defensive source-revision-bound reference. Candidates describe repeated visual structure without exposing raw XML or asset bytes; only an inspect-only candidate with a closed top-level graph can be passed to presentation.reuseSourceComponent, while ambiguous, opaque, or relationship-bound graphs carry an explicit blocked reason.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Examples:**

- presentation.resolveComponentCandidate(candidate.candidateId)

**Schema parameters:**

- `candidateId` (string) required — Exact candidateId from a trusted imported presentation inspection.

**Schema returns:**

- `componentCandidate` (object|undefined) — Defensive repeated-visual reference bound to the imported source SHA-256. Only a closed top-level candidate can authorize presentation.reuseSourceComponent; raw XML and arbitrary partial-graph mutation remain unavailable.

**Returns:**

defensive componentCandidate record or undefined

#### `presentation.reuseSourceComponent`

Create a new source-bound slide containing one exact top-level repeated component occurrence from presentation.inspect({ includeComponentCandidates: true }). The candidateId, occurrenceIndex, source revision, closed-graph ownership, sibling deletion proofs, and retained connector targets are checked before a complete source slide clone is projected by deleting only codec-proven sibling elements. Nested, opaque, ambiguous, comment-bound, relationship-bound, or stale candidates fail closed; the original slide and all non-target source parts remain untouched.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `candidateId` (string) required — Exact candidateId from presentation.inspect({ includeComponentCandidates: true }).
- `occurrenceIndex` (number) — Zero-based occurrence index from the candidate record; defaults to 0.
- `expectedCandidate` (object) — Optional complete candidate record from the same inspection; it is compared to current source-bound ownership evidence.

**Schema returns:**

- `slide` (Slide) — Pending source-bound slide clone containing only the selected top-level component. Sibling elements are removed only when their source deletion proofs and retained connector targets are safe; stale, nested, opaque, ambiguous, or relationship-bound candidates fail closed. Export/reimport before further edits.

#### `presentation.reuseSourceSlide`

Reuse one inspected imported slide as a source-bound complete graph after matching its exact slideId, sourceRevisionSha256, and optional clone-capability ownership evidence. The operation delegates to the codec-proven slide clone profile; stale revisions, unsupported graphs, and mismatched ownership evidence fail closed before the pending clone is created.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `slideId` (string) required — Exact slideId from trusted presentation inspection.
- `sourceRevisionSha256` (string) required — Exact 64-character source revision SHA-256 from the same inspection.
- `expectedCloneCapability` (object) — Optional complete cloneCapability record from the same inspection; it is compared to the current source-bound ownership evidence before reuse.

**Schema returns:**

- `slide` (Slide) — Pending source-bound slide clone inserted after the selected slide. The clone must remain unchanged until export/reimport; unsupported graphs and stale source or capability evidence fail closed.

#### `presentation.sections.add`

Define a native PowerPoint p14:sectionLst entry for source-free OfficeKit export. Sections together must form the complete ordered slide partition. Canonical imported sections may change only existing names and contiguous boundaries while count, order, stable facade identity, and native GUID stay fixed; irregular graphs remain opaque.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `name` (string) required — Unique 1-255-character section name, compared case-insensitively.
- `slides` (PresentationSlide[]|string[]) required — One or more slide facades or stable slide IDs from this presentation. Across all sections, memberships must partition every deck slide exactly once and in deck order.
- `nativeId` (string) — Optional preserved brace-delimited GUID for native p14:section/@id; new source-free sections receive a deterministic GUID.

**Schema returns:**

- `section` (PresentationSection) — Appended native PowerPoint p14:sectionLst entry. Source-free authoring owns the complete ordered slide partition. Canonical imported sections keep count, order, facade identity, and native GUID fixed; only names and contiguous partition boundaries may change. Extension-bearing or irregular section graphs remain opaque and fail closed.

#### `presentation.sections.getItem`

Resolve a source-free or canonical imported PowerPoint section by zero-based index, stable facade ID, or exact name.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `idOrNameOrIndex` (string|number) required — Stable section ID, exact name, or zero-based collection index.

**Schema returns:**

- `section` (PresentationSection|undefined) — Matching PowerPoint-section facade or undefined.

#### `presentation.slides.add`

Append an editable core slide with optional hidden slideshow state, a bounded source-free layout, direct ECMA-376 base transition, solid/style-reference background, and plain-text speaker notes. A supplied layout is resolved and materialized transactionally; effective imported Layout/Master inheritance is never flattened.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `name` (string) — Inspectable slide name.
- `hidden` (boolean) — Whether the slide is skipped by the ordinary slide show. Source-free hidden slides write p:sld/@show=0; visible slides omit the default-valued attribute.
- `layout` (string|object) — Optional bounded layout name/ID/facade. slides.add resolves it transactionally and materializes its text placeholders; an unknown or cross-presentation layout leaves no slide behind.
- `background` (string|object) — Optional direct slide background: RGB/theme color or { fill, mode: 'solid'|'reference', index? }. Gradient, pattern, image, transform, and effect-bearing backgrounds are preview-only/source-preserved and fail closed on canonical mutation.
- `transition` (object) — Optional direct ECMA-376 base transition. effect is one of blinds/checker/circle/comb/cover/cut/diamond/dissolve/fade/newsflash/plus/pull/push/random/randomBar/split/strips/wedge/wheel/wipe/zoom. Effect-specific fields are direction, orientation, throughBlack, or spokes (1..8); common fields are slow/medium/fast speed, durationMs 0..86400000, advanceOnClick, and advanceAfterMs 0..86400000. durationMs controls transition playback; advanceAfterMs controls slide advancement.
- `notes` (string|PresentationParagraph[]) — Optional speaker notes authored into the canonical PresentationML notes graph. A paragraph has runs plus ordinary direct paragraph/run styling; note-local links, fields, picture bullets, list styles, and body layout are rejected.

**Schema returns:**

- `slide` (Slide) — Appended editable slide. A supplied bounded source-free layout is bound and materialized immediately.

#### `presentation.slides.insert`

Insert a source-free slide after an existing Slide or 0-based index, or at the beginning with after: null. It uses the same hidden-state, transactional layout, direct base-transition, notes, and background profile as slides.add; imported additions fail closed, while slide.duplicate and slide.delete each have their own narrow source-preserving OPC profiles.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `after` (Slide|number|null) — Existing slide facade or 0-based index to insert after; null inserts first. Omit to append.
- `name` (string) — Inspectable slide name.
- `hidden` (boolean) — Whether the new source-free slide is skipped by the ordinary slide show.
- `layout` (string|object) — Optional bounded layout name/ID/facade. The new source-free slide is created and materialized transactionally.
- `background` (string|object) — Optional direct slide background: RGB/theme color or { fill, mode: 'solid'|'reference', index? }.
- `transition` (object) — Optional direct transition with the same complete ECMA-376 base-effect, speed, click, and timer profile as presentation.slides.add.
- `notes` (string|PresentationParagraph[]) — Optional speaker notes authored into the canonical PresentationML notes graph. A paragraph has runs plus ordinary direct paragraph/run styling; note-local links, fields, picture bullets, list styles, and body layout are rejected.

**Schema returns:**

- `slide` (Slide) — Inserted source-free slide. Unknown insertion targets or layouts leave the collection unchanged; imported additions remain fail-closed. See slide.duplicate for the separate bounded source-preserving clone profile.

#### `presentation.slideSize`

Read or set the deck canvas in pixels. On a trusted imported PPTX, a changed size is a deliberately canvas-only source-bound operation: OfficeKit updates only ppt/presentation.xml p:sldSz, clears an old preset type, and leaves slide, layout, master, chart, and shape coordinates unchanged. It never silently rescales or reflows content; callers must make any layout edits explicitly.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `width` (number) required — Finite non-negative canvas width in pixels; a changed imported canvas must resolve to a positive signed 32-bit EMU value.
- `height` (number) required — Finite non-negative canvas height in pixels; a changed imported canvas must resolve to a positive signed 32-bit EMU value.

**Schema returns:**

- `slideSize` ({ width: number, height: number }) — Current deck canvas. A trusted imported PPTX may change only this p:sldSz canvas; existing slide, layout, master, chart, and shape coordinates are preserved exactly, and callers must explicitly recompose any affected layout.

#### `presentation.textRange`

Inspect or resolve stable textRange anchors such as shapeId/text for editable slide text frames.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `id` (string) required — Stable shape text-range ID ending in /text.

**Schema returns:**

- `textRange` (TextRange|undefined) — Editable slide text-range facade or undefined.

#### `presentation.theme`

Inspect the model theme and theme inheritance. Custom source-free themes are not authored by OfficeKit 0.2, and imported themes are source-bound and read-only.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `name` (string) — Model theme name. Source-free customization is rejected; imported theme metadata is read-only.
- `colors` (object) — Complete tx1/bg1/tx2/bg2, accent1-accent6, hlink, and folHlink color scheme; dk1/lt1/dk2/lt2 aliases are accepted.
- `fonts` (object) — Major/minor Latin plus optional East-Asian and complex-script font families.
- `textStyles` (object) — Slide Master title/body/other defaults with fontSize, bold, italic, color, fontFamily, and alignment.
- `colorMap` (object) — Slide Master semantic color mapping for bg1/tx1/bg2/tx2, accents, and hyperlinks.

**Schema returns:**

- `theme` (PresentationTheme) — Inspectable model theme; canonical export accepts only the default source-free theme and preserves imported themes read-only.

#### `presentation.validateLayout`

Detect layout QA issues across slides, including off-canvas elements, geometry overlaps, and basic text overflow. Explicit text-free accessibility.decorative objects are excluded from overlap and partial-bleed errors; confirm their crop in the rendered slide.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `minOverlapArea` (number) — Minimum overlap area in square pixels before reporting.
- `boundsPadding` (number) — Allowed padding outside the slide bounds.
- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — Layout QA result with ok, issues, ndjson, and truncated. Text-free objects explicitly marked accessibility.decorative are excluded from overlap and partial-bleed errors; a fully invisible or meaningful object remains an offCanvas error.

#### `presentation.verify`

Return QA issues for layout validation, missing master/layout references, placeholder fidelity, chart/data consistency, table shape, image data, and dangling comments.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/review-deliver.md#evidence

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `minOverlapArea` (number) — Minimum overlap area for layout QA.
- `boundsPadding` (number) — Allowed padding outside slide bounds.
- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — Presentation semantic/layout QA result.

#### `presentation.view`

Control local editor gridline/guide visibility and inspect imported PowerPoint grid spacing, snap settings, and guides. Visibility is local model state; a separately capability-gated fixed-topology source-bound edit profile may change only already-present grid/snap values and guide positions in viewProps.xml.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `view` (PresentationView) — Local gridlinesVisible/guidesVisible state with show/hide/toggle methods, imported grid/snap/guide getters, and capability-aware source-bound editing. Local visibility is never persisted. Imported viewProps.xml may change only through setSourceProperties when capability.editable is true; that narrow profile retains field presence, guide count/order/orientation, relationships, extensions, and every non-editable XML residual.

#### `presentation.view.capability`

Return defensive sourceBound, partPresent, editable, existing-field, and guide-count evidence for the imported PPTX view-properties part. It is preflight evidence only; export re-proves hashes, topology, and the non-editable XML residual.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Defensive { sourceBound, partPresent, editable, gridSpacingCxEmuPresent, gridSpacingCyEmuPresent, slideViewSnapToGridPresent, slideViewSnapToObjectsPresent, guideCount } evidence. editable is true only for a relationship-free imported fixed-topology p:viewPr profile. It is preflight evidence, not mutable authority; export independently re-proves the source part, binding hashes, topology, and residual XML.

#### `presentation.view.setSourceProperties`

Change already-present imported grid spacing, snap flags, and existing guide positions only when view.capability.editable is true. It cannot create viewProps.xml, add/remove/reorient guides, write showGuides, or reconstruct extensions/relationships; unsupported profiles fail closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `gridSpacingCxEmu` (number) — Optional positive signed 32-bit EMU value. The imported p:gridSpacing/@cx attribute must already be present.
- `gridSpacingCyEmu` (number) — Optional positive signed 32-bit EMU value. The imported p:gridSpacing/@cy attribute must already be present.
- `slideViewSnapToGrid` (boolean) — Optional replacement for an already-present p:cSldViewPr/@snapToGrid attribute.
- `slideViewSnapToObjects` (boolean) — Optional replacement for an already-present p:cSldViewPr/@snapToObjects attribute.
- `slideGuides` ({ orientation: 'horizontal'|'vertical', position: integer }[]) — Optional complete existing guide-position list. Count, order, and horizontal/vertical orientation must exactly match the imported list; only positions may change.

**Schema returns:**

- `view` (PresentationView) — Returns the same view after a local requested source-bound patch. The method requires at least one field and view.capability.editable. It never writes local gridline/guide visibility or source p:cSldViewPr/@showGuides, creates no view-properties part, and rejects topology/relationship/extension changes; OfficeKit re-proves those constraints at export.

#### `PresentationFile.exportPptx`

Serialize PPTX through the single bundled OfficeKit codec. Only limits is accepted; legacy codec and lossy-fallback options fail explicitly.

**Adoption tier:** `compatibility`

**Use when:**

- A package-level or legacy interoperability operation is explicitly required.
- The caller can provide source-bound evidence and perform a second import.

**Avoid when:**

- Do not use as the default authoring route; use the typed Presentation facade first.
- Do not infer that an opaque or unsupported object became editable.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- Re-import the output and compare package/source evidence.
- Report unsupported or preserved content explicitly.

**Recipes:**

- skills/presentations/skills/presentations/tasks/review-deliver.md#evidence

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `presentation` (Presentation) required — Presentation facade to serialize.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — Native OOXML PPTX package bytes.

#### `PresentationFile.importPptx`

Import PPTX through the single bundled OfficeKit codec with bounded free-positioned p:sp lines including direct line ends/caps/joins, source-bound opaque preservation, speaker-notes edit/add capability evidence, bounded text-only edits for recognized local SlidePart placeholders and canonical SmartArt nodes whose fixed direct paragraphs retain optional empty paragraphs and contain between one and 256 total plain runs plus canonical fixed a:br leaves, eligible OLE XLSX payload access/replacement plus uniquely bound DOCX Office-package access/replacement, and fail-closed unsupported edits.

**Adoption tier:** `compatibility`

**Use when:**

- A package-level or legacy interoperability operation is explicitly required.
- The caller can provide source-bound evidence and perform a second import.

**Avoid when:**

- Do not use as the default authoring route; use the typed Presentation facade first.
- Do not infer that an opaque or unsupported object became editable.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- Re-import the output and compare package/source evidence.
- Report unsupported or preserved content explicitly.

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `pptx` (FileBlob|Uint8Array) required — PPTX package bytes.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `presentation` (Presentation) — Imported presentation facade with editable core objects, bounded free-positioned p:sp lines with the shared six-style/line-end/cap/join profile, bounded text-only replacement for recognized owner-local SlidePart placeholders, recognized direct slide backgrounds, canonical fixed-topology recursive groups, literal bar/line/pie/standard-area/fixed-doughnut/marker-scatter/2D-bubble charts plus the clustered bar+line combo profile with either shared primary axes or a secondary line pair, legacy text-only speaker notes plus fixed-topology relationship-free rich notes and explicit edit/add capability evidence, bounded legacy slide-level comments (unchanged-only), bounded Office 2021 modern root/direct-reply threads (text/status editable), and payload-only replacement for eligible source-bound OLE workbooks plus the uniquely bound DOCX Office-package profile. A notes-absent slide can add a canonical NotesSlide only when the source NotesMaster/SlideMaster Theme graph is re-proven safe. Chart formulas/external data and advanced plot topology remain source-bound. Compound/theme/custom-dash/effect/extension line outlines, placeholder identity/geometry/formatting and inherited Master/Layout graphs, complex backgrounds/groups, field/link/picture-bullet/layout-bearing notes, mixed line groups, secondary bars, irregular comment anchors/reactions/task fields, themes, arbitrary OLE, other native objects, and unsupported package graphs remain source-bound.

#### `PresentationFile.inspectPptx`

Inspect bounded PPTX parts, content types, the required presentation/root officeDocument relationship, namespace-aware source XML references, legacy notes/comments evidence, and Office 2021 modern author/thread/anchor semantics after raw-input, part-count, decompression, and optional compression-ratio budgets; verifyCrc32 additionally checks ZIP entry CRCs.

**Adoption tier:** `compatibility`

**Use when:**

- A package-level or legacy interoperability operation is explicitly required.
- The caller can provide source-bound evidence and perform a second import.

**Avoid when:**

- Do not use as the default authoring route; use the typed Presentation facade first.
- Do not infer that an opaque or unsupported object became editable.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- Re-import the output and compare package/source evidence.
- Report unsupported or preserved content explicitly.

**Recipes:**

- skills/presentations/skills/presentations/tasks/review-deliver.md#evidence

**Example paths:**

- examples/create-pptx-compose.mjs

**Examples:**

- await PresentationFile.inspectPptx(pptx, { includeText: true, maxChars: 12000 })

**Schema parameters:**

- `pptx` (FileBlob|Uint8Array) required — PPTX package bytes.
- `includeText` (boolean) — Include bounded XML, relationship, and JSON text previews.
- `maxPreviewChars` (number) — Maximum preview characters per textual package part.
- `maxInputBytes` (number) — Maximum compressed input bytes checked before JSZip parses the package.
- `maxParts` (number) — Maximum package part count.
- `maxPartBytes` (number) — Maximum uncompressed bytes per part.
- `maxTotalBytes` (number) — Maximum total uncompressed package bytes.
- `maxCompressionRatio` (number) — Optional maximum declared uncompressed/compressed ZIP-entry ratio; zero or omitted disables this extra check.
- `verifyCrc32` (boolean) — Verify every ZIP entry CRC32 before inspecting package structure; use for untrusted retained inputs.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `package` (object) — PPTX package result with ok, issues, parts, records, bounded NDJSON, and notes/comments semantic validation evidence.

#### `PresentationFile.patchPptx`

Apply path-validated PPTX part patches, including safe slide/master/layout ID lists and slide image/chart DrawingML mutations, and atomically reject dangling package references or invalid notes/comments semantics.

**Adoption tier:** `compatibility`

**Use when:**

- A package-level or legacy interoperability operation is explicitly required.
- The caller can provide source-bound evidence and perform a second import.

**Avoid when:**

- Do not use as the default authoring route; use the typed Presentation facade first.
- Do not infer that an opaque or unsupported object became editable.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- Re-import the output and compare package/source evidence.
- Report unsupported or preserved content explicitly.

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `pptx` (FileBlob|Uint8Array) required — PPTX package bytes.
- `patches` (array|object) required — Safe part edits with text, xml, json, bytes, content, remove, or delete.
- `maxInputBytes` (number) — Maximum compressed input bytes checked before JSZip parses the package.
- `maxPatchBytes` (number) — Maximum bytes per replacement part.
- `maxParts` (number) — Maximum resulting package part count.
- `maxPartBytes` (number) — Maximum uncompressed bytes per source or resulting part.
- `maxTotalBytes` (number) — Maximum total uncompressed source or resulting package bytes.
- `maxCompressionRatio` (number) — Optional maximum declared uncompressed/compressed ZIP-entry ratio; zero or omitted disables this extra check.
- `syncContentTypes` (boolean) — Synchronize inferred or explicit content-type declarations; defaults to true.
- `syncRelationships` (boolean) — Remove relationships to deleted parts and apply relationship recipes; defaults to true.
- `syncSourceReferences` (boolean) — Apply opt-in standard sourceReference XML mutations for supported semantic recipes; defaults to true.
- `validateResult` (boolean) — Validate final content types, relationships, and PPTX notes/comments semantics atomically; defaults to true. Set false only for deliberate invalid-package fixtures.
- `recipe` (string|object) — Standard OOXML part recipe with optional source/id/target and sourceReference fields; PPTX supports slide/master/layout ID lists plus image/chart objects in a slide shape tree.
- `sourceReference` (boolean|object) — Opt-in semantic XML mutation. Image/chart objects require explicit pixel position { left, top, width, height }, validate generated or explicit non-visual objectId, and clean matching slide objects on deletion.
- `relationship` (object) — Per-patch source/id/type/target/targetMode relationship recipe; explicit ID collisions require replaceExisting:true. relationships accepts an array.

**Schema returns:**

- `blob` (FileBlob) — Patched PPTX FileBlob with part/relationship/content-type/source-reference update counts and validation metadata.

#### `shape.accessibilityCapability`

Report sourceBound/editable/addable preflight for ordinary-shape p:cNvPr title/description/decorative metadata; export re-proves it.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, editable, addable } preflight; export revalidates p:cNvPr.

#### `shape.delete`

Explicitly remove a source-free shape or one capability-proven imported top-level ordinary shape. Relationship-owning shapes, connector/comment/timing/extension identity graphs, nested children, and raw collection mutation fail closed; pictures, connectors, tables, and charts expose their own typed deletion capability.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `shape` (Shape) — The removed Shape facade. Source-free deletion checks current connector/comment references. Imported deletion additionally requires shape.deletionCapability.supported, records explicit deletion intent, and export independently re-proves the source. Native objects, nested group children, relationship-owning shapes, and direct array splicing remain unsupported; groups, pictures, connectors, tables, and charts use their typed delete methods.

#### `shape.deletionCapability`

Report whether one imported top-level ordinary shape is inside the bounded element-deletion profile, with a package-local native ID used for post-write absence proof. Export recomputes the capability from source bytes.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, known, supported, blockedReason, nativeId } preflight. nativeId is package-local p:cNvPr identity evidence, not a cross-file artifact ID. Imported export ignores caller claims and re-proves the direct ShapeTree parent, element hash, unique native ID, relationship-free subtree, and absence of connector/comment/timing/extension identity consumers.

#### `shape.setAccessibilityMetadata`

Transactionally add, change, or clear non-visible ordinary-shape title/description/decorative metadata. Imported irregular p:cNvPr graphs fail closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `update` (object) required — { title?, description?, decorative? }; null clears a field, strings require 1-1,024 XML-safe characters, decorative requires a boolean, and a classification change plus its text clears/additions must be one transaction.

**Schema returns:**

- `shape` (Shape) — Same Shape. Source-free and canonical imported metadata is editable; unknown, duplicate, or malformed decorative extension graphs fail closed.

#### `shape.text.set`

Set plain or structured text with ordered text, field, and line-break inlines; bounded run formatting; character, picture-bullet, or auto-numbered lists; levels, indents, spacing; and external URI, internal-slide, relative-action, or existing custom-show hyperlinks. Missing, opaque, malformed, relationship-bearing, or dangling custom-show targets and unmodeled text graphs fail closed in canonical PPTX export.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `text` (string|string[]|object|object[]) required — Plain text, paragraph strings, inline arrays, or paragraph objects. Canonical OfficeKit export supports ordered text, fields, styled line breaks, bounded run/paragraph formatting, character and picture bullets, auto-numbering, levels, indents, spacing, tab stops, and one absolute uri, target slideId, relative action (nextSlide, previousSlide, firstSlide, lastSlide, endShow), or existing customShow name per link. customShow may include returnToSlide and survives the bounded slide clone as the same relationship-free stable-identity action without adding the clone to show membership; missing, opaque, malformed, relationship-bearing, or dangling targets fail closed.

**Schema returns:**

- `textFrame` (TextFrame) — The same live text frame with normalized paragraphs and a backward-compatible flattened value.

#### `shape.useBackgroundFill`

Read the presence-aware imported PresentationML p:sp useBgFill flag. It affects preview paint but remains source-bound and read-only; source-free authoring or wire mutation fails closed.

**Adoption tier:** `compatibility`

**Use when:**

- A package-level or legacy interoperability operation is explicitly required.
- The caller can provide source-bound evidence and perform a second import.

**Avoid when:**

- Do not use as the default authoring route; use the typed Presentation facade first.
- Do not infer that an opaque or unsupported object became editable.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- Re-import the output and compare package/source evidence.
- Report unsupported or preserved content explicitly.

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `useBackgroundFill` (boolean|undefined) — True/false only when the native attribute was present; otherwise undefined.

#### `slide.addNotes`

Set speaker notes as text or relationship-free paragraph/run data for inspect, preview, and canonical PPTX output. OfficeKit authors source-free notes, preserves the legacy text-only edit path, and edits a fixed imported rich paragraph/run topology; fields, hyperlinks, picture bullets, notes-body list styles/layout, and unsafe NotesMaster graphs remain source-bound and fail closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `text` (string|PresentationParagraph[]) required — Speaker notes text or paragraph/run data. Each structured paragraph follows the presentation text subset; note-local hyperlinks, fields, picture bullets, list styles, and body properties are rejected.

**Schema returns:**

- `notes` (object) — Speaker-notes record. Source-free notes and simple hash-bound imported text remain editable; an imported relationship-free rich body may edit only its fixed paragraph/inline topology. A notes-absent imported slide may add a canonical NotesSlide only when speakerNotes.capability.addable is true; export re-proves that package graph. Fields, hyperlinks, picture bullets, notes-page layout, list styles, and unsafe NotesMaster graphs remain preservation-only.

#### `slide.animations.add`

Add one bounded native object animation for fade, wipe, fly, zoom, or pulse. Use withPrevious, afterPrevious, or onClick to express speaking order; textBuild reveals whole text or paragraphs, and chartBuild reveals chart content by all-at-once, series, category, series-element, or category-element. The typed surface writes canonical PowerPoint timing and never accepts raw XML.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/motion.md#typed-surface

**Example paths:**

- skills/presentations/skills/presentations/examples/officekit-motion-workflow.mjs

**Schema parameters:**

- `target` (Shape|ImageElement|TableElement|ChartElement|Connector|GroupShape|string) required — A target on this slide or its stable target ID.
- `effect` (string) — fade, wipe, fly, zoom, or pulse.
- `phase` (string) — entrance, emphasis, or exit.
- `start` (string) — withPrevious, afterPrevious, or onClick.
- `direction` (string) — Required for wipe/fly: left, right, up, or down.
- `durationMs` (number) — Positive integer duration from 1 through 60000.
- `delayMs` (number) — Optional integer delay from 0 through 60000.
- `textBuild` (string) — Optional whole or paragraph text build.
- `chartBuild` (string) — Optional all-at-once, series, category, series-element, or category-element chart build.
- `staggerMs` (number) — Optional per-item stagger from 0 through 10000.
- `animateChartBackground` (boolean) — Whether the chart background participates in the chart build; defaults to false.

**Schema returns:**

- `animation` (object) — The normalized animation record. Imported timing must be capability-editable; source-free timing is emitted as canonical PresentationML.

#### `slide.animations.remove`

Remove one animation issued by slide.animations or identified by its stable animation ID. Imported timing must be capability-editable; opaque timing is preserved and rejects mutation.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/motion.md#typed-surface

**Example paths:**

- skills/presentations/skills/presentations/examples/officekit-motion-workflow.mjs

**Schema parameters:**

- `animation` (object|string) required — An animation record or stable animation ID returned by slide.animations.

**Schema returns:**

- `boolean` (Whether one existing animation was removed.)

#### `slide.applyLayout`

Bind a slide to a bounded source-free layout and materialize its effective direct-frame placeholder shapes. Applying the same layout is idempotent; switching a materialized layout fails closed. The resulting p:ph identities and direct frames export natively; imported Layout relationships remain preservation-only.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/continue.md#reinspect

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `layout` (string|SlideLayoutTemplate) required — Layout name/ID or layout facade.

**Schema returns:**

- `shapes` (Shape[]) — Binds the slide and materializes effective direct-frame title/body/ctrTitle/subTitle placeholder shapes for native source-free PPTX output. Reapplying the same layout is idempotent; switching an already-materialized layout fails closed.

#### `slide.autoLayout`

Place existing shapes inside a frame using horizontal or vertical flow, gap, padding, and alignment options.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `shapes` (object[]) required — Existing editable slide elements.
- `frame` (string|object) — slide, a frame object, or an element facade.
- `direction` (string) — horizontal or vertical.
- `horizontalGap` (number|string) — Horizontal gap or auto.
- `verticalGap` (number|string) — Vertical gap or auto.
- `horizontalPadding` (number) — Left/right inset.
- `verticalPadding` (number) — Top/bottom inset.
- `align` (string) — Cross-axis alignment.

**Schema returns:**

- `shapes` (object[]) — The positioned input elements.

#### `slide.charts.add`

Add a source-free literal bar, line, pie, standard area, fixed 50%-hole doughnut, marker-only scatter, bounded 2D bubble, or clustered bar+line combo chart. Category families use shared literal categories; scatter and bubble use aligned per-series numeric X/Y values, with positive area-based bubble sizes. Bar and line series, including combo members, accept up to 16 bounded native linear, exponential, logarithmic, power, polynomial, or moving-average trendlines plus one fixed/percentage/standard-deviation/standard-error/custom-literal errorBars projection. Imported trendline count and error-bar presence are fixed; unsupported labels/extensions/unknown children/complex lines remain source-owned. Supported variants retain title, legend, bounded axes, basic series styling, chart-level data labels, layout JSON, error-bar-aware SVG preview, and native ChartPart output across import/edit/re-export. Formula-backed custom error bars without an explicit embedded-workbook route, other formula/external data, advanced family geometry, topology changes, and unsupported styling fail closed rather than being flattened.

**Adoption tier:** `golden`

**Use when:**

- Use a de-defaulted evidence chart when comparison, change, distribution, or contribution is the page's primary claim.
- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- skills/presentations/skills/presentations/examples/officekit-design-decisions-workflow.mjs

**Schema parameters:**

- `chartType` (string) — bar, line, pie, standard area, fixed 50%-hole doughnut, marker-only scatter, bounded 2D bubble, or combo for canonical OfficeKit export. combo is the literal clustered bar+line profile described by series; unsupported or advanced family variants fail closed.
- `title` (string) — Chart title.
- `categories` (string[]) — Shared literal labels required by bar, line, pie, area, doughnut, and combo. Scatter and bubble reject shared categories and use per-series xValues.
- `series` (object[]) required — One or more named series. Category charts require one finite value per category. Scatter and bubble require aligned finite xValues and Y values; bubble additionally requires aligned positive bubbleSizes. Markers are limited to line and scatter, and marker-only scatter rejects a series line in favor of marker.line. Bar and line series, including combo members, accept up to 16 trendlines with type linear/exp/log/power/poly/movingAvg, optional name, type-specific order/period, half-category forward/backward forecasts, intercept, equation/R-squared flags, and bounded RGB line. They also accept one errorBars object: reference type standardError/percentage/standardDeviation/none or canonical direction x/y, type both/minus/plus, valueType fixedVal/percentage/stdDev/stdErr/cust, bounded value or exact-count custom side arrays, cap policy, and bounded RGB line. Imported trendline count and error-bar presence are fixed; labels, extensions, unknown children, malformed caches, and complex/theme lines remain source-owned. Formula-backed PPTX custom sides require a separately supported embedded-workbook route. For combo, every series declares chartType bar or line; there must be at least one primary bar and one line. Bars cannot be secondary. Lines are either all primary or all axisGroup: secondary; mixed primary/secondary line plots fail closed. Other formula sources, point overrides, per-series labels, smooth, and per-series chart types outside combo fail closed.
- `externalData` (object|FileBlob|ArrayBuffer|Uint8Array|string) — Model-only external/embedded workbook metadata. OfficeKit 0.2 source-free charts require literal categories and values and reject externalData.
- `position` (object) — Pixel left/top/width/height frame.
- `axes` (object) — Basic axis titles, number formats, intervals, bounds, and major units. Category families use a category/value pair; scatter and bubble use two numeric value axes; pie and doughnut reject axes. A combo with all lines axisGroup: secondary may also set axes.secondary.category and axes.secondary.value, written at top/right. Secondary axes are invalid for primary-line combos, mixed line groups, or secondary bars.
- `legend` (object) — Legend options.
- `dataLabels` (boolean|object) — Chart-level showValue/showCategoryName/showSeriesName, circular-only showPercent for pie/doughnut, and a supported bounded position. Per-series overrides are unsupported.
- `styleId` (number) — Model-only chart style metadata; it is not part of the bounded OfficeKit chart wire.
- `styleIndex` (number) — Model-only alias for styleId.
- `varyColors` (boolean) — Model-only varied-color preference outside the bounded OfficeKit chart wire.
- `barOptions` (object) — Model-only advanced bar layout options outside the bounded OfficeKit chart wire.
- `lineOptions` (object) — Model-only advanced line grouping/smoothing options; direct per-series marker formatting remains supported.
- `accessibility` (object) — Non-visible { title?, description?, decorative? }. Strings require 1-1,024 XML-safe characters. decorative is a presence-aware boolean: true is mutually exclusive with title/description, explicit false differs from omission, and the Office 2019+ value maps through the canonical adec:decorative extension. Maps to p:nvGraphicFramePr/p:cNvPr independently of the visible chart title and object name.

**Schema returns:**

- `chart` (ChartElement) — Appended editable native-chart facade.

#### `slide.clearBackground`

Remove the direct slide background so preview and PPTX output inherit from the preserved Layout/Master chain. Unsupported imported background graphs fail closed rather than being flattened or discarded.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `slide` (Slide) — The same slide with no direct background, inheriting from its preserved Layout/Master chain.

#### `slide.clearBackgroundImage`

Remove the image previously authored by slide.setBackgroundImage without changing the slide's solid/theme background.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/layered-composition.md#public-surface

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `slide` (Slide) — The same slide after removing its authored background-image layer, if present.

#### `slide.clearMorph`

Clear a source-free or capability-approved Morph transition. Imported unknown Morph extensions remain preserved and reject mutation.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/motion.md#typed-surface

**Example paths:**

- skills/presentations/skills/presentations/examples/officekit-motion-workflow.mjs

**Schema returns:**

- `slide` (Slide) — The same slide with no authored Morph transition.

#### `slide.clearNativeBackgroundImage`

Remove the direct native p:bg image while preserving the inherited Layout/Master background and leaving any ordinary setBackgroundImage layer untouched.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/layered-composition.md#public-surface

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `slide` (Slide) — Remove the direct native p:bg image authored or replaced by slide.setNativeBackgroundImage, restoring the preserved Layout/Master background chain. It never removes an ordinary setBackgroundImage scene-layer picture.

#### `slide.clearTransition`

Remove one canonical direct imported or source-free slide transition. A transition-absent imported slide remains a no-op until an explicit capability-approved add; timing, sound, extension, and opaque-effect graphs remain byte-preserved and reject mutation.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `slide` (Slide) — The same slide with no direct p:transition. Removing an imported transition requires the same canonical editable source profile as replacement.

#### `slide.cloneCapability`

Report whether an imported SlidePart can be copied as one ownership-checked OPC graph. The Codec copies every uniquely owned descendant, DataPart, and external relationship while rebinding proven shared layout, NotesMaster, image, and retained-slide targets. Sections, modern comments, outside-owned nodes, removed slide-jump targets, and over-budget graphs fail closed before the model changes.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Defensive { sourceBound, known, supported, blockedReason, clonedPartCount, sharedPartCount, sourceRevisionSha256 } preflight. clonedPartCount includes the SlidePart and uniquely owned OpenXmlPart descendants; sharedPartCount reports recognized resources rebound to the source package; sourceRevisionSha256 binds a reuse request to the exact imported package. Export ignores caller claims and re-analyzes the hash-bound package graph.

#### `slide.comments.addThread`

Create either a bounded legacy PPTX annotation or an Office 2021 modern thread. A comment-free imported presentation may add canonical legacy review comments only when comments.capability.addable is true; a canonical imported legacy leaf with comments.capability.editable permits only existing root-text replacement, never addThread/replies/metadata edits. Modern mode supports a top-level element/text-range/textMatch anchor, one root, direct replies, independent people/timestamps, and active/resolved/closed state; imported modern graphs permit only fixed-topology text/status edits.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `target` (undefined|string|Shape|ImageElement|TableElement|ChartElement|ConnectorElement|TextRange|object) — Legacy mode requires undefined. Modern mode accepts a top-level element/text-range ID or facade, { element }, { textRange }, or { textMatch: { element, query, occurrence? } }. Nested group-child and slide-level modern anchors remain unsupported.
- `text` (string) required — Root comment text.
- `author` (string) — Display author. Modern comments may instead provide comments[0].person with brace-delimited GUID id, name, initials, userId, and providerId.
- `position` (object) required — Explicit slide coordinate { x, y, unit?: 'px'|'emu' }. Legacy defaults to px; modern defaults to emu.
- `resolved` (boolean) — Modern root state. Legacy mode requires false.
- `created` (string) — ISO-8601 creation time for the root comment; defaults to the Unix epoch for deterministic output.
- `nativeFormat` (string) — Set modern for explicit Office 2021 authoring; Presentation.create({ commentFormat: 'modern' }) must select the same wire family.
- `comments` (object[]) — Optional root record. Modern records support nativeId/id, authorId/person, text, created, and active/resolved/closed status. Reactions, task fields, extensions, and nested replies fail closed.

**Schema returns:**

- `thread` (SlideCommentThread) — Create a bounded legacy annotation or Office 2021 modern root. A comment-free imported presentation may create canonical legacy parts only after comments.capability.addable preflight; OfficeKit re-proves the whole source graph, allocates collision-free relationships, and never mixes comment families. Recognized legacy imports remain unchanged-only. Recognized modern imports expose root/direct replies and allow only text/status edits; author/person/date identity, position, target moniker, reply topology, part paths, relationships, and source hashes remain fixed.

#### `slide.comments.capability`

Inspect defensive source-bound comment-family evidence before authoring or editing. A comment-free imported presentation may advertise legacy addability; one closed imported legacy leaf may instead advertise editable, which permits only its existing root text to change while author/time/coordinate/native identity/order/topology remain fixed. Modern graphs retain their separate fixed-topology edit contract.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Defensive { sourceBound, format, partPresent, editable, addable } evidence. For imported files, addable is true only when the complete presentation has no legacy or Office 2021 comment graph and OfficeKit can create one canonical legacy CommentAuthorsPart plus slide-local SlideCommentsPart leaves. editable is true only for an existing closed legacy leaf with one relationship-free author catalog and a re-proven fixed comment topology; then only the existing root text may change. Author, timestamp, coordinate, package-local author/index identity, order, count, relationships, and family remain fixed. This is preflight evidence, not mutable write authority; export re-proves the source bytes and fails closed on existing irregular, mixed, connected, or tampered graphs.

#### `slide.compose`

Materialize a clean-room compose tree with row, column, grid, layers, box, paragraph/text, shape, table, chart, image, and rule nodes into editable slide objects. Capture the returned elements for later edits or connector targets; compose nodes remain declarative and are not Shape facades.

**Adoption tier:** `golden`

**Use when:**

- Use free composition for an asymmetric editorial page or restrained recurring motif instead of a universal container grid.
- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- skills/presentations/skills/presentations/examples/officekit-design-decisions-workflow.mjs

**Schema parameters:**

- `node` (object) required — Compose tree rooted in row, column, grid, layers, box, paragraph/text, shape, table, chart, image, or rule.
- `frame` (object) — Pixel materialization frame; defaults to an inset slide frame.

**Schema returns:**

- `elements` (object[]) — Materialized editable slide elements. Capture this return value when a later edit or connector needs a Shape/Table/Chart/Image facade; the input compose nodes themselves are declarative and have no stable object identity.

#### `slide.connectors.add`

Legacy low-level connector authoring from explicit points or target centers. Prefer slide.shapes.connect or geometry: connector when DrawingML target-plus-site identity matters.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `from` (string|object) — Legacy start element/ID or point. Without a modeled site it uses the supplied start point or target center.
- `to` (string|object) — Legacy end element/ID or point. Without a modeled site it uses the supplied end point or target center.
- `start` (object) — Explicit start point {x,y}.
- `end` (object) — Explicit end point {x,y}.
- `connectorType` (string) — straight, elbow, or curved; defaults to straight.
- `line` (object) — Line color, width, solid/dashed/none style, and compatibility arrow metadata.
- `accessibility` (object) — Non-visible { title?, description?, decorative? }. Strings require 1-1,024 XML-safe characters. decorative is a presence-aware boolean: true is mutually exclusive with title/description, explicit false differs from omission, and the Office 2019+ value maps through the canonical adec:decorative extension. Maps to p:nvCxnSpPr/p:cNvPr.

**Schema returns:**

- `connector` (ConnectorElement) — Appended legacy low-level connector. Prefer slide.shapes.connect or direct geometry: connector for explicit target-plus-site identity.

#### `slide.continuationCapability`

Report full-authoring, pending-clone (export/reimport first), or bounded-overlay. Bounded overlay token-preserves the tree and allows one clean export of listed basic shapes/images. Separate SlidePart edits by reviewed revision.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/continue.md#reinspect

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Defensive { sourceBound, ready, profile, requiresExportReopen, oneSlideMutationPerExport?, shapeGeometries?, embeddedImage?, sourceRevisionSha256? }. pending-clone requires export/reimport. bounded-overlay permits only listed shapes/images; other native additions and mixed SlidePart edits stay blocked.

#### `slide.delete`

Remove this slide. Source-free decks may remove any non-final slide. An imported PPTX first requires deletionCapability.supported, then removes the real SlidePart and every exclusively owned descendant (including closed notes/comments/chart/OLE/diagram/media leaves) while retaining shared parts. Inbound slide references and presentation-level custom-show/section/extension identity remain fail closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `result` (undefined) — No return value. The slide must belong to a presentation with at least two slides. Imported deletion checks slide.deletionCapability before mutating the model, then export independently recomputes the graph. The actual SlidePart, its relationships, and every exclusively owned OpenXml/DataPart descendant are removed; shared layout/master/theme/image/media descendants remain. Any inbound slide reference or presentation-level custom-show/section/extension identity fails closed.

#### `slide.deletionCapability`

Report whether an imported SlidePart and its exclusively owned OPC descendant closure can be deleted. The count includes the slide plus owned OpenXml/DataPart descendants; shared layout/master/theme/media remain outside the closure. Export re-proves the graph from source bytes and aggregates all requested slide deletions into one ownership transaction.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Defensive { sourceBound, known, supported, blockedReason, ownedPartCount } per-slide preflight. Imported capability presence makes known true; export ignores caller claims and recomputes one aggregate exclusive OPC descendant closure for every requested slide deletion from the hash-bound source package.

#### `slide.duplicate`

Clone one original imported PPTX slide after slide.cloneCapability proves a bounded ownership graph. The JavaScript model copies the unchanged semantic element tree and resolves connector targets to fresh clone-local identities; the OfficeKit Codec then creates a distinct SlidePart, recursively byte-copies every uniquely owned OpenXmlPart and DataPart with exact local relationship IDs and external links, and rebinds only proven shared layout, NotesMaster, image, slide-jump, and other identity resources. Custom-show membership is unchanged. The pending clone cannot be edited, cloned twice, or lose its origin before export/reimport. Source-free slides, sections, modern comments, outside-owned unknown nodes, removed slide-jump targets, unresolved semantic elements/connectors, pending native payload replacements, and over-budget graphs fail closed.

**Adoption tier:** `golden`

**Use when:**

- The requested presentation intent is covered by this bounded, inspect-backed primitive.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- fresh presentation.inspect() evidence when editing an imported file

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/continue.md#reinspect

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `slide` (Slide) — A new adjacent Slide available only when slide.cloneCapability.supported is true and the original imported slide remains semantically unchanged. The Codec recursively copies the uniquely owned OPC descendant closure, including unknown OpenXmlParts, DataParts, relationship-bearing modeled leaves, and external relationships, while preserving relationship IDs and exact bytes. It reuses proven shared layouts, NotesMaster, images, and retained slide targets. The pending clone must stay unchanged until export/reimport. Sections, modern comments, any owned node with an outside parent, a jump to a removed slide, unresolved semantic elements or connector targets, pending native replacements, repeated pending clones, origin deletion, and graph-budget overflow fail closed before partial model mutation.

**Notes:**

- Clone eligibility is graph ownership, not a native-object type whitelist. Unknown or relationship-bearing descendants are accepted when they are uniquely owned, within budget, and unchanged; their bytes, content types, child/external relationships, and DataParts are copied recursively. A descendant with any parent outside the owned closure is not guessed safe and blocks the operation.
- Open XML SDK allocates collision-free package URIs for copied parts. Agents must use imported object IDs and inspect/resolve results rather than assuming physical names such as slide2.xml. After export and reimport, modeled edits to an independently copied chart, OLE package, SmartArt, InkML, media, notes, or comments leaf remain subject to that feature's own edit capability.

#### `slide.elements`

Read the slide's direct cross-type scene stack from back to front. Shapes, textboxes, images, tables, charts, connectors, and groups share this order; type-specific collections remain indexes over the same elements.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/layered-composition.md#public-surface

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `elements` (object) — Read-only collection facade whose items are the direct slide elements in bottom-to-top export order.

#### `slide.groups.add`

Author recursive native DrawingML p:grpSp trees with optional non-visible group title/description/decorative metadata, outer off/ext, and local chOff/chExt coordinates. The bounded profile supports modeled shapes, connectors, images, tables, charts, and nested groups; canonical imported groups allow fixed-topology semantic edits, while group-level fills/effects, locks, transforms, unknown extensions, or unsupported descendants remain opaque and read-only.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `name` (string) — Inspectable group name.
- `accessibility` (object) — Non-visible { title?, description?, decorative? }. Strings require 1-1,024 XML-safe characters. decorative is a presence-aware boolean: true is mutually exclusive with title/description, explicit false differs from omission, and the Office 2019+ value maps through the canonical adec:decorative extension. Maps to p:nvGrpSpPr/p:cNvPr.
- `position` (object) required — Group frame in parent or slide pixel coordinates.
- `childFrame` (object) — Local child coordinate rectangle mapped through DrawingML chOff/chExt; defaults to the group width/height from 0,0.
- `shapes` (object[]) — Initial child shape/textbox definitions in local coordinates.
- `connectors` (object[]) — Initial child connector definitions in local coordinates.
- `groups` (object[]) — Initial nested group definitions.
- `tables` (object[]) — Initial native DrawingML table definitions in local coordinates.
- `charts` (object[]) — Initial relationship-backed chart definitions in local coordinates.
- `images` (object[]) — Initial relationship-backed picture definitions in local coordinates.
- `children` (object[]) — Ordered mixed child definitions using kind shape, connector, groupShape, table, chart, or image.

**Schema returns:**

- `group` (GroupShape) — Appended recursive grouped-shape facade for resolve, inspect, layout, SVG preview, and native p:grpSp export. Canonical imported groups are source-bound and editable without changing child topology; complex group shells or unsupported descendants are preserved as one opaque read-only object.

#### `slide.hide`

Hide this slide from the ordinary slide show through the same source-bound p:sld/@show primitive as slide.setHidden(true).

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `slide` (Slide) — The same slide after setting hidden=true through the bounded p:sld/@show source-bound primitive.

#### `slide.images.add`

Add an embedded image with accessibility metadata, fit/crop, frame, rotation/flips, layout, preview, and PPTX output. Ready bounded-overlay accepts rectangular images in a clean export. OfficeKit writes native p:cNvPr, decorative metadata, and a:srcRect.

**Adoption tier:** `golden`

**Use when:**

- Use an image-led composition when a supplied, referenced, sourced, or generated image carries the page's context or emotion.
- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- skills/presentations/skills/presentations/examples/officekit-design-decisions-workflow.mjs

**Schema parameters:**

- `blob` (FileBlob) — Embedded image bytes loaded from a local task asset; avoids constructing a large base64 string in Agent code.
- `dataUrl` (string) — Embedded image data URL.
- `uri` (string) — External image URI metadata.
- `prompt` (string) — Generation/source prompt metadata.
- `alt` (string) — Compatibility alias for accessibility.description. Reading or writing alt reads or writes the same state; an empty string clears description.
- `accessibility` (object) — Non-visible { title?, description?, decorative? }. Strings require 1-1,024 XML-safe characters. decorative is a presence-aware boolean: true is mutually exclusive with title/description, explicit false differs from omission, and the Office 2019+ value maps through the canonical adec:decorative extension. Maps to p:nvPicPr/p:cNvPr independently of the object name and visible pixels. When this object is supplied, prompt metadata is never synthesized as alt text.
- `fit` (string) — contain, cover, or stretch. For embedded images, OfficeKit computes a bounded native a:srcRect from intrinsic dimensions; imported native source rectangles normalize to fit stretch plus explicit crop because PPTX has no fit keyword.
- `crop` (object) — Optional normalized { left, top, right, bottom } source edges in -1..1 with opposing sums below 1. Positive values crop; negative values expand for contain/letterbox semantics. Manual crop is applied before contain/cover fitting.
- `position` (object) — Pixel left/top/width/height frame.
- `transform` (object) — Optional { rotationDegrees, flipHorizontal, flipVertical } center transform. OfficeKit preserves explicit false and safely edits recognized top-level embedded pictures.

**Schema returns:**

- `image` (ImageElement) — Appended editable image facade. OfficeKit authors/imports embedded PNG/JPEG/GIF/safe-SVG rectangular pictures and permits native source-rectangle add/edit/remove plus same-format byte, name/title/description/decorative metadata, frame, and direct-transform edits; unmodeled cNvPr children remain residual-protected, while effects, external sources, complex blips, and non-rectangular geometry remain opaque.

#### `slide.moveTo`

Move this slide to an existing 0-based deck index. On an imported PPTX, OfficeKit rewrites only the retained source SlidePart order in the presentation slide-ID list; unrelated topology changes and broad graph clones remain fail-closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `index` (number) required — Existing zero-based destination index. It must be an integer from 0 through presentation.slides.items.length - 1.

**Schema returns:**

- `slide` (Slide) — The same slide at its new collection position. Imported PPTX export rewrites only p:sldIdLst for the retained source SlideParts; unrelated topology changes and broad graph clones fail closed. See slide.duplicate and slide.delete for their separate constrained source-part contracts.

#### `slide.placeholders.getItem`

Resolve a slide placeholder shape by stable ID, name, placeholder type, or numeric index. Imported placeholder.textEditable reports a verified local SlidePart text capability; identity, geometry, formatting, layout binding, and inherited Master/Layout graphs remain source-bound.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create-from-template.md#distill-and-reuse

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `idOrNameOrTypeOrIndex` (string|number) required — Placeholder stable ID, display name, type, or numeric idx.

**Schema returns:**

- `shape` (Shape|undefined) — Matching placeholder shape or undefined. Imported shape.placeholder.textEditable is true only when the source binding recognizes the concrete SlidePart's local text body. In that case text.set(...) preserves native formatting/topology while replacing characters; use text.replace(...) for an in-run edit. The capability is re-proved from source on export and cannot be granted by mutating the model flag. Identity, geometry, formatting, and layout binding remain read-only.

#### `slide.setBackground`

Set a direct slide background to a six-digit RGB/theme color solid fill or a native style reference. Recognized imported direct backgrounds are hash-bound and editable; inherited Layout/Master backgrounds remain inherited.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `background` (string|object) required — Direct RGB/theme color or { fill, mode: 'solid'|'reference', index? }; reference index must be an unsigned 32-bit integer.

**Schema returns:**

- `slide` (Slide) — The same slide with a normalized direct background; canonical PPTX export never flattens inherited Layout/Master backgrounds.

#### `slide.setBackgroundImage`

Add or replace one full-slide embedded image at the bottom of a source-free scene stack. Combine it with a translucent shape and editable foreground objects for image-led pages. Imported slides reject authored underlays because they cannot be placed beneath the complete source-bound prefix without changing native order.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/layered-composition.md#public-surface

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `blob` (FileBlob) — Embedded PNG, JPEG, GIF, or safe SVG bytes. Prefer FileBlob.load(path, { type }) over building base64 in task code.
- `dataUrl` (string) — Embedded image data URL when a FileBlob is not available.
- `fit` (string) — cover by default; contain or stretch are also accepted.
- `crop` (object) — Optional normalized source crop { left, top, right, bottom }.
- `alt` (string) — Compatibility alias for accessibility.description.
- `accessibility` (object) — Non-visible { title?, description?, decorative? }. Strings require 1-1,024 XML-safe characters. decorative is a presence-aware boolean: true is mutually exclusive with title/description, explicit false differs from omission, and the Office 2019+ value maps through the canonical adec:decorative extension.

**Schema returns:**

- `image` (ImageElement) — The authored full-slide image at stackIndex 0. Repeated calls replace the same image. Source-bound slides reject because a new image cannot be inserted beneath preserved native elements.

#### `slide.setHidden`

Set whether this slide is skipped by the ordinary slide show. OfficeKit writes only p:sld/@show, uses absence for visible and show=0 for hidden, and re-proves the source-bound SlidePart before export.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `hidden` (boolean) required — true writes the canonical native show=0 state; false clears the attribute to PresentationML's visible default.

**Schema returns:**

- `slide` (Slide) — The same slide with updated ordinary-slide-show visibility. Only p:sld/@show changes; content, layout, relationships, transitions, notes, comments, and static slide pixels remain fixed. Unknown or irregular imported lexical values fail closed.

#### `slide.setLayout`

Alias of slide.applyLayout(layout): bind and materialize a bounded source-free layout for native PPTX export.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `layout` (string|SlideLayoutTemplate) required — Layout name/ID or layout facade.

**Schema returns:**

- `slide` (Slide) — Alias of applyLayout that returns the slide.

#### `slide.setMorph`

Author a bounded cross-slide Morph transition between adjacent slides with real source and destination objects and unique named object pairs. OfficeKit gives both objects the same Selection Pane identity; unknown imported Morph extensions remain source-bound and are not reconstructed.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/motion.md#typed-surface

**Example paths:**

- skills/presentations/skills/presentations/examples/officekit-motion-workflow.mjs

**Schema parameters:**

- `morph` (object) required — { from: immediatelyPreviousSlide, durationMs?, pairs: [{ key, from: sourceObject, to: destinationObject }] }; one through 256 unique named pairs.

**Schema returns:**

- `slide` (Slide) — The same destination slide with a bounded Morph transition. Both paired objects receive the same !!key Selection Pane identity; charts, non-adjacent slides, incompatible kinds, duplicate objects, name conflicts, and conflicting transitions reject.

#### `slide.setNativeBackgroundImage`

Set a direct native p:bg/p:bgPr/a:blipFill image stretched across the slide. It stays behind all slide content and is not a reorderable or animatable scene-layer picture; use slide.setBackgroundImage when you need a movable or animated image layer.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/references/layered-composition.md#public-surface

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `blob` (FileBlob) — Embedded PNG, JPEG, GIF, or safe SVG bytes. Prefer FileBlob.load(path, { type }) over building base64 in task code.
- `dataUrl` (string) — Embedded image data URL when a FileBlob is not available.
- `assetId` (string) — Existing content-addressed presentation image asset ID, normally obtained from an imported artifact.
- `fit` (string) — Must be stretch; native p:bg does not accept crop, cover, contain, transform, effects, or external links.
- `alphaModulationFixed` (boolean) — Optional preservation flag for the source's parameterless a:alphaModFix child; omission preserves it when replacing a recognized imported native image background.

**Schema returns:**

- `slide` (Slide) — Set the direct native PresentationML background as p:bg/p:bgPr/a:blipFill with one embedded image stretched across the slide. The image is behind all slide content, remains editable through this method, and is not a scene-stack element that can be reordered or animated. Source-bound edits require a recognized direct-background profile; complex crop, tile, effect, linked, or inherited backgrounds remain opaque and fail closed.

#### `slide.setTransition`

Set one direct p:transition from the complete 21-effect ECMA-376 base vocabulary, with effect-specific direction/orientation/throughBlack/spokes plus speed, Office 2010+ durationMs, and click/timer advancement. Source-free slides may author it; imported slides may replace one canonical existing direct transition or add one only when transition.capability.addable is true. Timing, sound, Office-extension effects, non-integer-unit duration, and irregular source graphs fail closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `transition` (object) required — A complete ECMA-376 base-transition object. Effect-specific fields are direction (cardinal, corner, or in/out as applicable), orientation (horizontal/vertical), throughBlack (cut/fade), or spokes (wheel, 1..8). speed defaults to medium, advanceOnClick to true, and independent durationMs and advanceAfterMs fields accept 0..86400000.

**Schema returns:**

- `slide` (Slide) — The same slide with a normalized direct p:transition. Source-free slides may author it. An imported slide may replace exactly one canonical direct base transition, or add one only when transition.capability.addable proves the root contains only p:cSld plus optional p:clrMapOvr and has no transition, timing, or extension leaf. Opaque source graphs are not reconstructed.

#### `slide.shapes.add`

Add a shape/textbox, free-positioned p:sp line, or exact-site p:cxnSp connector with accessibility metadata. Ready bounded-overlay accepts only textbox/rect/roundRect/ellipse in a clean export. Lines support dash/ends/cap/join; custom geometry supports ordered adjustment/guide formulas, XY/polar adjustment handles, and connection sites. Only a connector retains target-plus-site identity.

**Adoption tier:** `golden`

**Use when:**

- Use editable native geometry and typography when their spatial relationship is the visual carrier.
- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- skills/presentations/skills/presentations/examples/officekit-design-decisions-workflow.mjs

**Schema parameters:**

- `name` (string) — Inspectable shape name.
- `geometry` (string) — rect, ellipse, roundRect, textbox, line, custom, or connector. line creates a free-positioned p:sp without targets; custom requires customPaths; connector creates p:cxnSp and requires from, to, fromIdx, and toIdx.
- `from` (Shape|string) — For connector geometry, the start shape facade or stable ID in this same slide/group tree.
- `to` (Shape|string) — For connector geometry, the end shape facade or stable ID in this same slide/group tree.
- `fromIdx` (number) — For connector geometry, the required unsigned DrawingML start connection-site index.
- `toIdx` (number) — For connector geometry, the required unsigned DrawingML end connection-site index.
- `kind` (string) — For connector geometry: straight, elbow/elbow2..5, or curved. Elbow aliases normalize to the canonical elbow model.
- `head` (object) — For connector geometry, optional { type, width, length } start line end.
- `tail` (object) — For connector geometry, optional { type, width, length } end line end.
- `cap` (string) — For connector geometry: flat, round, or square.
- `join` (string) — For connector geometry: round, bevel, or miter.
- `customAdjustments` (object[]) — For geometry custom, up to 256 ordered { name, formula } adjustment guides written to a:avLst. Names are bounded ASCII identifiers; formulas use the 17 ECMA-376 operators and may reference integer literals, DrawingML built-ins, or an earlier adjustment. Forward references, duplicate/reserved names, invalid arithmetic, and unsupported grammar fail closed.
- `customGuides` (object[]) — For geometry custom, up to 1,024 ordered { name, formula } calculated guides written to a:gdLst after customAdjustments. Each formula may reference integer literals, DrawingML built-ins, or any earlier adjustment/guide. Path, connection-site, handle, and text-rectangle fields share that built-in-plus-declared reference namespace.
- `customConnectionSites` (object[]) — For geometry custom, up to 1,024 ordered { angle, x, y } native a:cxnLst entries. Numeric angle is degrees; numeric x/y are shape-local pixels. Each value may instead reference one DrawingML built-in or declared adjustment/guide and must evaluate within one turn or the shape frame. Array index is connector identity: source-free shapes author it, recognized imports may edit values at existing indexes but keep the list length fixed, and connectors to custom shapes require explicit fromIdx/toIdx.
- `customAdjustmentHandles` (object[]) — For geometry custom, up to 1,024 ordered native a:ahLst entries with kind xy or polar. XY handles may control declared xAdjustment/yAdjustment names with paired min/max coordinate bounds; polar handles may control radialAdjustment with paired non-negative radius bounds and angleAdjustment with paired degree bounds. Bounds and x/y positions accept literals or DrawingML built-in/declared guide references. Bounds and current adjustment values are evaluated before export. Recognized imports may edit bounds/positions but keep each index's kind and controlled adjustment names fixed; malformed or broader handle topology remains opaque and fails closed.
- `customPaths` (object[]) — For geometry custom, 1-64 DrawingML paths with optional positive literal integer width/height and bounded moveTo, lineTo, quadraticBezTo, cubicBezTo, arcTo, and close commands. Omitted or zero extents use the shape-coordinate default independently per axis. Point coordinates and arc radii/angles accept a literal or one DrawingML built-in or declared custom adjustment/guide name. Each path may carry presence-aware fillMode (normal or none), stroke, and extrusionAllowed; omission preserves native defaults and extrusionAllowed is metadata rather than 3D authoring. arcTo radii must evaluate positive, requires a current point, and limits its evaluated non-zero sweep to one full turn. Unknown references, unsupported handle topology, and lighten/darken path-fill modes remain opaque or fail closed.
- `textRectangle` (object) — Optional { left, top, right, bottom } rectangle relative to a custom shape frame. Each edge is a finite pixel coordinate or a DrawingML built-in/declared adjustment/guide name; resolved right/bottom must exceed left/top. Numeric edges retain the deterministic private scaling-guide profile, reference edges write standard a:rect ST_AdjCoordinate values directly, and mixed rectangles round-trip. The state drives inspect, SVG origin, and overflow QA. Omission keeps the full-shape default; unknown references, malformed leaves, and invalid resolved bounds fail closed.
- `position` (object) — Pixel left/top/width/height frame. For geometry line, left/top is the start point and width/height is the non-negative endpoint delta; one extent may be zero, but both zero fail closed.
- `transform` (object) — Optional { rotationDegrees, flipHorizontal, flipVertical } center transform. Rotation is bounded to -360 through 360 degrees and flip booleans retain explicit false. OfficeKit authors/imports this direct DrawingML transform on supported shapes; complex or unknown native transform graphs remain read-only.
- `accessibility` (object) — Non-visible { title?, description?, decorative? }. Strings require 1-1,024 XML-safe characters. decorative is a presence-aware boolean: true is mutually exclusive with title/description, explicit false differs from omission, and the Office 2019+ value maps through the canonical adec:decorative extension. Maps to the native p:cNvPr of an ordinary p:sp or exact-site p:cxnSp connector and remains independent of visible text/name; irregular imports stay source-bound.
- `text` (string|string[]|object|object[]) — Plain text or structured paragraphs accepted by shape.text.set, including ordered text/field/line-break inlines, paragraph tab stops, styles, and relationship-backed hyperlinks. Run and paragraph styles accept fontFamilyEastAsia; when East Asian text has fontFamily but no explicit override, OfficeKit writes the same direct a:ea typeface so LibreOffice and PowerPoint do not rely on host font inference.
- `textBodyProperties` (object) — DrawingML text-frame layout: pixel insets; anchor/wrap/AutoFit; optional normalAutoFit { fontScale, lineSpacingReduction } percentages only with shrinkText; -360..360 degree rotation; horizontal/vertical/vertical270 text; horizontal/vertical overflow; 1-16 columns with pixel spacing and RTL flow; and upright text. Percentages retain at most three decimal places; noncanonical imported AutoFit markup remains source-bound.
- `fill` (string|object) — Shape fill.
- `line` (object) — For ordinary shapes and free lines: { style: solid|dashed|dotted|dash-dot|dash-dot-dot|none, fill, width, head?, tail?, cap?, join? }. Line ends use triangle|stealth|diamond|oval|arrow with optional sm|med|lg width/length; only geometry line accepts ends. dash/dot/dashDot/longDashDotDot and non-conflicting startArrow*/endArrow* aliases normalize canonically.
- `placeholder` (object) — Optional layout placeholder metadata. Free-positioned lines cannot be placeholders.

**Schema returns:**

- `shape` (Shape|ConnectorElement) — Appended editable shape/textbox, free-positioned p:sp line, or exact-site p:cxnSp connector. Unknown line profiles, unsupported site tables, and incomplete connector endpoints fail closed.

#### `slide.shapes.connect`

Connect two modeled shapes in the same slide/group tree by preset side or exact DrawingML connection-site index. Custom shapes require an explicit index into customConnectionSites. `head` is the from/start end and `tail` is the to/end end; use tail for a forward arrow, and bringToFront() when a background shape would hide the route. The target-plus-site pair survives import, edit, clone, and second import; moved or re-parameterized modeled targets reroute before render/export.

**Adoption tier:** `golden`

**Use when:**

- Use a relationship diagram whose connectors encode direction, causality, dependency, or handoff.
- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- skills/presentations/skills/presentations/examples/officekit-design-decisions-workflow.mjs

**Schema parameters:**

- `from` (Shape|string) required — Start shape facade or stable ID in this same slide/group tree.
- `to` (Shape|string) required — End shape facade or stable ID in this same slide/group tree.
- `kind` (string) — straight, elbow/elbow2..5, or curved; defaults to elbow.
- `fromSide` (string) — Preset rect/roundRect/textbox/ellipse top, left, bottom, or right. Mutually exclusive with fromIdx; custom shapes require fromIdx.
- `toSide` (string) — Preset rect/roundRect/textbox/ellipse top, left, bottom, or right. Mutually exclusive with toIdx; custom shapes require toIdx.
- `fromIdx` (number) — Exact unsigned DrawingML start connection-site index, including an index into a custom shape's ordered customConnectionSites.
- `toIdx` (number) — Exact unsigned DrawingML end connection-site index, including an index into a custom shape's ordered customConnectionSites.
- `line` (object) — { style: solid|dashed|none, fill, width } plus compatibility startArrow/endArrow fields.
- `head` (object) — Optional start (`from`) line end { type: none|triangle|stealth|diamond|oval|arrow, width?: sm|med|lg, length?: sm|med|lg }.
- `tail` (object) — Optional end (`to`) line end using the same bounded type/size union; use this for a usual forward arrow.
- `cap` (string) — flat, round, or square.
- `join` (string) — round, bevel, or miter.
- `accessibility` (object) — Non-visible { title?, description?, decorative? }. Strings require 1-1,024 XML-safe characters. decorative is a presence-aware boolean: true is mutually exclusive with title/description, explicit false differs from omission, and the Office 2019+ value maps through the canonical adec:decorative extension. Maps to p:nvCxnSpPr/p:cNvPr.

**Schema returns:**

- `connector` (ConnectorElement) — A source-free connector behind its nodes by default; call bringToFront() when it must remain visible above another layer. `head` is the from/start end and `tail` is the to/end end. Target movement reroutes modeled sites. Recognized imported direct connectors retain target-plus-site identity and may reorder only with an editable zOrderCapability.

#### `slide.shapes.getConnectionSiteIndex`

Resolve top/left/bottom/right to a stable bounded preset connection-site index for rect, roundRect, textbox, or ellipse. Custom shapes expose an ordered site table but require its explicit numeric index; other geometries fail closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `target` (Shape|string) required — Same-tree rect, roundRect, textbox, or ellipse shape.
- `side` (string) required — top, left, bottom, or right.

**Schema returns:**

- `siteIndex` (number) — The bounded preset connection-site index. Custom shapes require an explicit customConnectionSites index; unsupported geometry fails closed rather than guessing.

#### `slide.show`

Show this slide in the ordinary slide show by clearing the source-bound p:sld/@show leaf through slide.setHidden(false).

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `slide` (Slide) — The same slide after setting hidden=false and clearing canonical p:sld/@show. This does not add the slide to a custom show or alter custom-show membership.

#### `slide.speakerNotes.capability`

Return defensive sourceBound, partPresent, editable, and addable evidence. addable identifies an imported notes-absent slide whose source NotesMaster/SlideMaster Theme graph can safely receive a canonical NotesSlide. Export independently re-proves the package graph, so mutating model or wire data cannot grant authority.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Defensive { sourceBound, partPresent, editable, addable } evidence. addable is true only for an imported notes-absent slide whose presentation graph is safely extensible. It is Agent preflight evidence, not mutable write authority; OfficeKit independently re-proves the source package before export.

#### `slide.tables.add`

Add an inspectable table facade with rows, columns, values, cells, rectangular merges, layout JSON, SVG preview, and canonical OfficeKit plain-text PPTX output.

**Adoption tier:** `golden`

**Use when:**

- The agent is compiling or refining a presentation plan with an explicit reader outcome.
- The operation can be followed by the Presentation review and commit workflow.

**Avoid when:**

- Do not use it to bypass the active authoring plan or to edit raw package paths.
- Do not publish before semantic, structural, layout, and delivery review.

**Requires:**

- Presentation facade
- active authoring plan when the task creates a deck

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `values` (unknown[][]) required — Table cell value matrix.
- `name` (string) — Inspectable table name.
- `position` (object) — Pixel left/top/width/height frame.
- `style` (object) — Table/cell fill, margins, borders, and text style.
- `styleOptions` (object) — Optional headerRow and bandedRows booleans plus model-rendering font options. OfficeKit authors the two native flags, but keeps them immutable after source-bound import.
- `accessibility` (object) — Non-visible { title?, description?, decorative? }. Strings require 1-1,024 XML-safe characters. decorative is a presence-aware boolean: true is mutually exclusive with title/description, explicit false differs from omission, and the Office 2019+ value maps through the canonical adec:decorative extension. Maps to p:nvGraphicFramePr/p:cNvPr independently of visible cell text and the object name.

**Schema returns:**

- `table` (TableElement) — Appended editable table facade. OfficeKit accepts a non-empty rectangular 1-256-column by 1-2048-row plain-text grid with non-overlapping rectangular merges; recognized imports may change name, complete frame, and visible origin/unmerged cell text without changing merge topology or native style flags.

#### `slide.visibilityCapability`

Report whether the imported p:sld/@show state is known and editable. OfficeKit exposes the inverse Agent-facing hidden boolean; invalid native lexical values stay source-owned and fail closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Defensive { sourceBound, known, editable } preflight. known is false for an opaque or invalid native p:sld/@show value; editable is not mutable write authority because OfficeKit re-proves the source SlidePart and semantic hash at export.

#### `slideCommentThread.addReply`

Append a direct reply to a source-free Office 2021 modern comment thread. Imported reply topology is fixed: existing reply text/status may change, but adding or removing replies fails closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `text` (string) required — Direct reply text.
- `author` (string) — Reply display author.
- `person` (object) — Modern author identity with id/name/initials/userId/providerId.
- `created` (string) — ISO-8601 timestamp.
- `status` (string) — active, resolved, or closed; defaults to active.

**Schema returns:**

- `thread` (SlideCommentThread) — Append one direct source-free modern reply and return the thread. Imported topology changes fail closed.

#### `slideCommentThread.reopen`

Set the modern root comment status back to active while preserving fixed imported identity, anchor, position, and reply topology.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `thread` (SlideCommentThread) — Set resolved=false and the modern root status to active. Legacy comments cannot encode this state.

#### `slideCommentThread.resolve`

Set the modern root comment status to resolved. Imported export re-proves author/date/anchor/position/topology and source-part hashes before changing only status.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `thread` (SlideCommentThread) — Set resolved=true and the modern root status to resolved. Legacy comments cannot encode this state.

#### `table.accessibilityCapability`

Report sourceBound/editable/addable preflight for table graphic-frame p:cNvPr title/description/decorative metadata; export re-proves it.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, editable, addable } preflight; export revalidates the table graphic-frame p:cNvPr.

#### `table.delete`

Explicitly remove a source-free table or one capability-proven imported direct table p:graphicFrame. Relationship-bearing, irregular, nested, or identity-sensitive frames and raw collection mutation fail closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `table` (TableElement) — The removed TableElement facade. Imported deletion requires table.deletionCapability.supported and records explicit intent; export removes only the direct p:graphicFrame, validates native-ID absence, and rejects relationships, irregular or nested tables, identity-sensitive graphs, or direct array splicing.

#### `table.deletionCapability`

Report whether one imported top-level bounded relationship-free DrawingML table can be deleted, with a package-local native ID used for post-write absence proof. Export recomputes the source-bound capability.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, known, supported, blockedReason, nativeId } preflight. nativeId is package-local p:cNvPr evidence. Export ignores caller claims and re-proves one direct bounded DrawingML table p:graphicFrame, a relationship-free subtree, a unique native ID, and absence of connector/comment/timing/extension identity consumers.

#### `table.merge`

Merge one inclusive rectangular table range, retain the upper-left value, clear and lock covered cells, and emit canonical DrawingML merge topology.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/create.md#compose-and-review

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `range` (object) required — Inclusive zero-based { startRow, endRow, startColumn, endColumn } rectangle. It must span at least two in-bounds cells and cannot overlap an existing merge.

**Schema returns:**

- `table` (TableElement) — The same table after preserving the upper-left value, clearing covered values, and making covered cells read-only. Imported merge topology remains source-bound and cannot be changed.

#### `table.setAccessibilityMetadata`

Transactionally add, change, or clear non-visible table title/description/decorative metadata. Imported irregular graphic-frame p:cNvPr graphs fail closed.

**Adoption tier:** `advanced`

**Use when:**

- A specific advanced PresentationML capability is requested after its capability record has been inspected.
- The task can tolerate a narrower edit surface than the golden authoring routes.

**Avoid when:**

- Do not substitute it for the create, template, edit, continue, or review task route.
- Do not bypass source hashes, capability checks, or fail-closed boundaries.

**Requires:**

- Presentation facade
- capability or source evidence appropriate to the operation

**Review:**

- presentation.validateLayout and presentation.verify
- reviewArtifact with the active plan and changed page scope
- visualReview: complete, unavailable, or requires-human

**Recipes:**

- skills/presentations/skills/presentations/tasks/edit-existing.md#bounded-edit

**Example paths:**

- examples/create-pptx-compose.mjs

**Schema parameters:**

- `update` (object) required — { title?, description?, decorative? }; null clears a field, strings require 1-1,024 XML-safe characters, decorative requires a boolean, and a classification change plus its text clears/additions must be one transaction.

**Schema returns:**

- `table` (TableElement) — Same table. Source-free and canonical imported metadata is editable; unsupported graphic-frame p:cNvPr profiles fail closed without disabling unrelated supported table edits.

## shared

| Name | Kind | Summary |
| --- | --- | --- |
| `clearOfficeFontDesignMetrics` | api | Clear process-level and scoped Office font design metrics. |
| `createCanvasRenderer` | api | Create an optional node-canvas renderer adapter from office-kit/renderers/canvas for SVG/PNG/JPEG/WebP FileBlob raster conversion to PNG or JPEG. |
| `createLibreOfficeRenderer` | api | Create a LibreOffice CLI renderer adapter from office-kit/renderers/libreoffice for DOCX/XLSX/PPTX/HTML/PDF FileBlob conversion, typically to PDF. |
| `createNativeOfficeRenderer` | api | Create a native Office renderer adapter from office-kit/native/office-bridge that calls a JSON stdin/stdout sidecar command with timeout, temp-file isolation, cleanup, and structured errors. |
| `createPlaywrightRenderer` | api | Create an optional Playwright renderer adapter from office-kit/renderers/playwright for deterministic SVG/HTML to PNG, WebP, JPEG, or PDF conversion with network blocked by default. |
| `createPopplerRenderer` | api | Create a Poppler CLI renderer adapter from office-kit/renderers/poppler for application/pdf FileBlob page rasterization to PNG, PPM, or TIFF. |
| `createSharpRenderer` | api | Create an optional sharp renderer adapter from office-kit/renderers/sharp for SVG/PNG/JPEG/WebP FileBlob raster conversion to PNG, WebP, or JPEG. |
| `registerScopedOfficeFontDesignMetrics` | api | Register a last-in-first-resolved scoped font design-metric collection and return an idempotent disposer. |
| `renderArtifact` | api | Render an artifact through its render/export method, attach normalized FileBlob metadata, and optionally pass SVG output through a caller-provided renderer adapter for PNG/WebP/JPEG/PDF output. |
| `renderFileWithNativeOffice` | api | Render or convert a DOCX/XLSX/PPTX/PDF FileBlob through a configured native Office bridge command, returning a FileBlob for PDF/PNG/WebP or other requested output. |
| `resolveOfficeFontDesignMetrics` | api | Resolve the requested primary family, style, and nearest numeric weight from scoped then process-level font design metrics without silently skipping to later family fallbacks. |
| `reviewArtifact` | api | Reopen a final DOCX, XLSX, PPTX, or PDF and return one bounded post-edit report covering modeled semantics, package structure, representative render evidence, plan-bound presentation communication/narrative/cognitive/visual risks, an optional compact text reading view powered by AnyDoc, visual-review status, and delivery identity. |
| `setOfficeFontDesignMetrics` | api | Replace the process-level Office font design-metric registry with normalized public metric records used by deterministic layout integrations. |
| `skiaPaintBaselineCompensationPx` | api | Return the signed subpixel residual between a finite paint baseline and its nearest integer pixel, or zero for non-finite input. |
| `verifyArtifact` | api | Run an artifact's verify() method and return a bounded NDJSON QA report. |
| `visualQaArtifact` | api | Render an artifact, compare PNG/JPEG/WebP/PPM decoded pixels against a baseline render, optionally register small translations, and return a configurable aligned PNG diff heatmap. |

### shared details

#### `clearOfficeFontDesignMetrics`

Clear process-level and scoped Office font design metrics.

**Schema returns:**

- `result` (undefined) — All registered metrics are removed synchronously.

#### `createCanvasRenderer`

Create an optional node-canvas renderer adapter from office-kit/renderers/canvas for SVG/PNG/JPEG/WebP FileBlob raster conversion to PNG or JPEG.

**Examples:**

- const renderer = createCanvasRenderer({ width: 1200, height: 800, background: 'white' })

**Schema parameters:**

- `canvas` (object) — Injected node-canvas compatible module.
- `width` (number) — Output width override.
- `height` (number) — Output height override.
- `background` (string) — Canvas background color.
- `outputOptions` (object) — node-canvas encoder options.

**Schema returns:**

- `renderer` (function) — SVG/PNG/JPEG/WebP to PNG/JPEG renderer adapter.

#### `createLibreOfficeRenderer`

Create a LibreOffice CLI renderer adapter from office-kit/renderers/libreoffice for DOCX/XLSX/PPTX/HTML/PDF FileBlob conversion, typically to PDF.

**Examples:**

- const renderer = createLibreOfficeRenderer({ command: 'soffice', timeoutMs: 60000 })

**Schema parameters:**

- `command` (string) — soffice/LibreOffice executable path or command name.
- `format` (string) — Default target format, normally pdf.
- `convertTo` (string) — Explicit LibreOffice --convert-to filter value.
- `timeoutMs` (number) — CLI timeout.
- `tempRoot` (string) — Temporary directory root.
- `argsBuilder` (function) — Custom LibreOffice argument builder.
- `keepTemp` (boolean) — Keep temporary files for diagnostics.

**Schema returns:**

- `renderer` (function) — Office/HTML conversion renderer adapter.

#### `createNativeOfficeRenderer`

Create a native Office renderer adapter from office-kit/native/office-bridge that calls a JSON stdin/stdout sidecar command with timeout, temp-file isolation, cleanup, and structured errors.

**Examples:**

- const renderer = createNativeOfficeRenderer({ command: 'dotnet', args: ['OfficeBridge.dll'], timeoutMs: 60000 })

**Schema parameters:**

- `command` (string) — Native Office bridge executable.
- `args` (string[]) — Arguments passed before the bridge reads its JSON request from stdin.
- `timeoutMs` (number) — Bridge request timeout.
- `format` (string) — Default requested output format.
- `inputType` (string) — Default input MIME type.
- `outputType` (string) — Default output MIME type.
- `nativeOptions` (object) — Operation-specific native Office options.

**Schema returns:**

- `renderer` (function) — DOCX/XLSX/PPTX/PDF native Office renderer adapter.

#### `createPlaywrightRenderer`

Create an optional Playwright renderer adapter from office-kit/renderers/playwright for deterministic SVG/HTML to PNG, WebP, JPEG, or PDF conversion with network blocked by default.

**Examples:**

- const renderer = createPlaywrightRenderer({ viewport: { width: 900, height: 1200 }, deviceScaleFactor: 1 })

**Options:**

- viewport
- deviceScaleFactor
- allowNetwork
- timeoutMs
- format

**Schema parameters:**

- `viewport` (object) — Chromium viewport width and height; SVG geometry is inferred when omitted.
- `deviceScaleFactor` (number) — Chromium device scale factor.
- `allowNetwork` (boolean) — Permit network requests; disabled by default for deterministic rendering.
- `timeoutMs` (number) — Navigation and rendering timeout.
- `background` (string) — Page background CSS color.
- `chromium` (object) — Injected Playwright Chromium launcher for tests or custom runtimes.

**Schema returns:**

- `renderer` (function) — SVG/HTML to PNG/WebP/JPEG/PDF renderer adapter.

**Returns:**

renderer adapter function for renderArtifact(...)

#### `createPopplerRenderer`

Create a Poppler CLI renderer adapter from office-kit/renderers/poppler for application/pdf FileBlob page rasterization to PNG, PPM, or TIFF.

**Examples:**

- const renderer = createPopplerRenderer({ command: 'pdftoppm', dpi: 150 })

**Schema parameters:**

- `command` (string) — pdftoppm executable path or command name.
- `dpi` (number) — Raster resolution.
- `page` (number) — One-based PDF page number; pageIndex is the zero-based alias.
- `timeoutMs` (number) — CLI timeout.
- `tempRoot` (string) — Temporary directory root.
- `argsBuilder` (function) — Custom pdftoppm argument builder.
- `keepTemp` (boolean) — Keep temporary input/output files for diagnostics.

**Schema returns:**

- `renderer` (function) — PDF to PNG/PPM/TIFF page renderer adapter.

#### `createSharpRenderer`

Create an optional sharp renderer adapter from office-kit/renderers/sharp for SVG/PNG/JPEG/WebP FileBlob raster conversion to PNG, WebP, or JPEG.

**Examples:**

- const renderer = createSharpRenderer({ resize: { width: 1200 }, flatten: true })

**Schema parameters:**

- `sharp` (function) — Injected sharp factory; otherwise the optional peer dependency is loaded.
- `resize` (object) — sharp resize options.
- `flatten` (boolean|object) — Flatten transparency using background options.
- `background` (string|object) — Flatten background color.
- `pngOptions` (object) — sharp PNG encoder options.
- `webpOptions` (object) — sharp WebP encoder options.
- `jpegOptions` (object) — sharp JPEG encoder options.

**Schema returns:**

- `renderer` (function) — SVG/PNG/JPEG/WebP raster renderer adapter.

#### `registerScopedOfficeFontDesignMetrics`

Register a last-in-first-resolved scoped font design-metric collection and return an idempotent disposer.

**Schema parameters:**

- `entries` (object[]) required — Iterable normalized font design-metric candidates.

**Schema returns:**

- `dispose` (function) — Idempotently removes only this scoped registration.

#### `renderArtifact`

Render an artifact through its render/export method, attach normalized FileBlob metadata, and optionally pass SVG output through a caller-provided renderer adapter for PNG/WebP/JPEG/PDF output.

**Examples:**

- await renderArtifact(document, { format: 'png', renderer: createPlaywrightRenderer() })

**Options:**

- format
- renderer/rasterRenderer/renderAdapter
- page/pageIndex
- slide
- sheetName
- range

**Schema parameters:**

- `artifact` (Workbook|Presentation|DocumentModel|PdfArtifact) required — Artifact facade to render through its native preview/export path.
- `format` (string) — svg, png, webp, jpeg, pdf, layout, or an output MIME type.
- `renderer` (function) — Optional pluggable renderer adapter for raster/PDF conversion.
- `source` (string) — Optional native source such as docx or pdf for renderer gates.

**Schema returns:**

- `blob` (FileBlob) — Rendered output with normalized metadata.

**Returns:**

FileBlob with normalized render metadata

#### `renderFileWithNativeOffice`

Render or convert a DOCX/XLSX/PPTX/PDF FileBlob through a configured native Office bridge command, returning a FileBlob for PDF/PNG/WebP or other requested output.

**Examples:**

- await renderFileWithNativeOffice(docx, { command, format: 'pdf', artifactKind: 'document' })

**Schema parameters:**

- `input` (FileBlob|Uint8Array) required — Office/PDF input bytes.
- `command` (string) required — Native Office bridge executable.
- `args` (string[]) — Arguments passed to the bridge executable.
- `operation` (string) — Bridge operation, defaulting to render.
- `format` (string) — Requested output format.
- `artifactKind` (string) — document, workbook, presentation, or pdf.
- `timeoutMs` (number) — Bridge request timeout.
- `nativeOptions` (object) — Operation-specific native Office options.
- `keepTemp` (boolean) — Keep temporary files for diagnostics.

**Schema returns:**

- `blob` (FileBlob) — Native Office bridge output bytes and renderer metadata.

#### `resolveOfficeFontDesignMetrics`

Resolve the requested primary family, style, and nearest numeric weight from scoped then process-level font design metrics without silently skipping to later family fallbacks.

**Schema parameters:**

- `request` (object) required — { family: string[], weight?, style? }; the first family is the explicit lookup target.

**Schema returns:**

- `metric` (object|undefined) — A defensive normalized metric record or undefined.

#### `reviewArtifact`

Reopen a final DOCX, XLSX, PPTX, or PDF and return one bounded post-edit report covering modeled semantics, package structure, representative render evidence, plan-bound presentation communication/narrative/cognitive/visual risks, an optional compact text reading view powered by AnyDoc, visual-review status, and delivery identity.

**Examples:**

- await reviewArtifact('/absolute/path/output.pptx', { authoringPlan, changedPageIds: ['page-04'], playbackEvidence: 'structural', visualReview: 'unavailable' })
- await reviewArtifact('/absolute/path/output.pptx', { source: '/absolute/path/input.pptx', contentView: 'anydoc', visualReview: 'unavailable' })

**Options:**

- format/kind
- outputPath
- source
- baseline
- authoringPlan
- changedPageIds
- playbackEvidence
- contentView
- visualReview
- layout
- renderOptions
- maxBytes
- maxContentChars
- maxInspectChars
- maxSummaryChars

**Schema parameters:**

- `input` (string|FileBlob|Uint8Array|Blob|Workbook|Presentation|DocumentModel|PdfArtifact) required — Final artifact path, bytes, or model. Modeled input is exported and reopened before review.
- `format` (string) — Required only when raw bytes do not carry a supported MIME type; docx, xlsx, pptx, or pdf.
- `source` (string|FileBlob|Uint8Array|Blob) — Optional read-only source used for SHA-256 and canonical input/output collision evidence.
- `baseline` (string|FileBlob|Uint8Array|Blob|Workbook|Presentation|DocumentModel|PdfArtifact) — Optional pre-edit artifact. Exact matching semantic/layout issues are marked preexisting warnings; structural package failures and new errors still fail the review.
- `outputPath` (string) — Absolute or working-directory-relative final path when reviewing an in-memory model.
- `authoringPlan` (object) — Optional office-kit/presentation-authoring-plan/v1 plan. Presentation review checks its communication strategy, narrative/cognitive/visual risks, design grammar, and motion intent against the candidate.
- `changedPageIds` (string[]) — Optional stable plan page IDs for a local edit. Non-target page changes become design-scope errors.
- `playbackEvidence` (string) — structural, keynote, or powerpoint. Host values require actual playback evidence and are not inferred from XML.
- `contentView` (string|boolean) — Set to anydoc or true to request the bounded text reading view. Omitted, none, or false does not initialize its AnyDoc parser.
- `visualReview` (string) — Caller-attested complete, unavailable, or requires-human. Text reading/OCR output never qualifies as complete.
- `layout` (boolean) — Set false only when a separate render review is already recorded; otherwise a representative render check runs.
- `renderOptions` (object) — Existing visualQaArtifact render/baseline options. PDF defaults to one native MuPDF PNG page.
- `maxBytes` (number) — Positive input/source byte budget checked before parser work.
- `maxContentChars` (number) — Positive character budget for the text reading view's Markdown.
- `maxInspectChars` (number) — Positive semantic/structural/layout evidence character budget.
- `maxSummaryChars` (number) — Positive combined review-summary character budget.

**Schema returns:**

- `report` (object) — Schema-v1 post-edit report. Verdict is passed, passed-with-limitations, or failed; Presentation design.strategy and design.layers remain bounded risk evidence, while factual, visual, and playback judgments remain explicit and separate.

**Returns:**

{ verdict, semantic, structural, layout, design, motion, playbackEvidence, contentView, visualReview, delivery, baseline, summary }

**Notes:**

- The text reading view is runtime-lazy and optional; AnyDoc is its parser backend. It is not a structural authority, render validator, OCR route, or substitute for direct pixel/aesthetic review.
- Do not request the text reading view routinely. Use contentView='anydoc' only when it can close an identified text or table content-coverage gap; it does not resolve OCR, layout, image, formula, or metadata-provenance gaps.
- playbackEvidence='structural' proves only timing targets and package structure. Use keynote or powerpoint only after actual host playback.
- For a Presentation authoring plan, design.strategy records the communication job, scenario, direction, medium fit, delivery mode, and after-use; design.layers groups deterministic communication, narrative, cognitive, and visual-risk findings. These signals do not verify factual truth or aesthetic quality.

#### `setOfficeFontDesignMetrics`

Replace the process-level Office font design-metric registry with normalized public metric records used by deterministic layout integrations.

**Schema parameters:**

- `entries` (object[]) required — Iterable records with family, weight, unitsPerEm, ascent, non-negative descent, and optional lineGap/style/width.

**Schema returns:**

- `result` (undefined) — Registry replacement is synchronous.

#### `skiaPaintBaselineCompensationPx`

Return the signed subpixel residual between a finite paint baseline and its nearest integer pixel, or zero for non-finite input.

**Schema parameters:**

- `value` (number) required — Baseline coordinate in CSS pixels.

**Schema returns:**

- `compensation` (number) — A finite residual in the interval [-0.5, 0.5).

#### `verifyArtifact`

Run an artifact's verify() method and return a bounded NDJSON QA report.

**Examples:**

- verifyArtifact(workbook, { maxChars: 12000 })

**Options:**

- maxChars

**Schema parameters:**

- `artifact` (Workbook|Presentation|DocumentModel|PdfArtifact) required — Artifact exposing a verify() method.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `report` (object) — Semantic QA result with artifactKind, ok, issues, ndjson, and truncated.

**Returns:**

{ artifactKind, ok, issues, ndjson, truncated }

#### `visualQaArtifact`

Render an artifact, compare PNG/JPEG/WebP/PPM decoded pixels against a baseline render, optionally register small translations, and return a configurable aligned PNG diff heatmap.

**Examples:**

- await visualQaArtifact(document, { baseline, pixelDiff: true, minBytes: 100 })

**Options:**

- baseline/expected/baselineBlob
- pixelDiff
- diffImage
- diffPalette
- diffAlignment
- pixelRegistration
- PNG/JPEG/WebP/PPM raster pixel comparison
- allowChange
- minBytes
- maxBytes
- maxChars

**Schema parameters:**

- `artifact` (Workbook|Presentation|DocumentModel|PdfArtifact) required — Artifact to render and compare.
- `format` (string) — Requested render format such as svg, png, ppm, jpeg, webp, or pdf.
- `renderer` (function) — Optional renderer adapter used for format conversion.
- `baseline` (FileBlob|Uint8Array) — Expected render bytes; expected and baselineBlob are aliases.
- `pixelDiff` (boolean|object) — Enable PNG/JPEG/WebP/PPM pixel comparison, optional channel thresholds, and decoded-pixel limits.
- `diffImage` (boolean) — Set false to disable PNG heatmap generation for changed raster baselines.
- `diffPalette` (object) — Optional changed/unchanged RGB colors and alpha values for the PNG heatmap.
- `diffAlignment` (string) — Dimension-mismatch behavior: strict (no heatmap), top-left, or center alignment on a union canvas.
- `pixelRegistration` (boolean|number|object) — Optionally search a bounded baseline translation (up to 8 pixels) before comparison; records sampled and exact before/after metrics plus ignored edge pixels.
- `allowChange` (boolean) — Allow baseline byte/pixel changes without emitting issues.
- `minBytes` (number) — Warn when the render is smaller than this byte count.
- `maxBytes` (number) — Warn when the render exceeds this byte count.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `report` (object) — Visual QA result with ok, blob, optional diffBlob PNG heatmap, summary, issues, ndjson, and truncation metadata.

**Returns:**

{ ok, blob, diffBlob, summary, issues, ndjson }

## workbook

| Name | Kind | Summary |
| --- | --- | --- |
| `exportXlsxWithOfficeKit` | api | Export the bounded Workbook model through the bundled C# Open XML SDK NativeAOT codec: cells, formulas, styles, merges, dimensions, freezes, ordinary tables, PNG/JPEG pictures, validation, conditional formatting, threaded-comment roots with direct replies, bar/line/pie/area/doughnut charts, marker-only numeric-X/Y scatter charts, bounded numeric-X/Y/positive-Size bubble charts, standard Office 2010 line/column/stacked sparklines, canonical one-variable or two-variable What-If data tables, native PivotTables with exact item or absolute whole-day date filters, and one source-free XLDAPR dynamic-array metadata profile. Imported QueryTables permit only source-bound one-way refresh hardening through table.setQueryRefreshPolicy; an imported connection may only change explicit refreshOnLoad=true to false through workbook.disableConnectionRefreshOnLoad; a uniquely owned imported Pivot cache may make that same one-way change through pivot.disableRefreshOnLoad; imported dynamic-array topology, commands, fields, sorts, other Pivot configuration/data/output, and unsupported extension graphs are preservation-only or fail closed. |
| `fx.ABS` | formula | Return the absolute value of a number. |
| `fx.ACOS` | formula | Return the inverse cosine for a finite input in the closed interval [-1,1]; other inputs fail as #NUM!. |
| `fx.ACOSH` | formula | Return the inverse hyperbolic cosine for finite inputs at least one; smaller inputs fail as #NUM!. |
| `fx.ADDRESS` | formula | Return one bounded worksheet address as text from 1-based row and column numbers, reference mode 1 through 4, A1 or R1C1 style, and optional Excel-quoted sheet text. Coordinates outside XFD1048576, invalid modes, nonlogical style selectors, and nontext sheet names fail as #VALUE!. |
| `fx.AND` | formula | Return TRUE when all conditions are true. |
| `fx.ASIN` | formula | Return the inverse sine for a finite input in the closed interval [-1,1]; other inputs fail as #NUM!. |
| `fx.ASINH` | formula | Return the inverse hyperbolic sine of a finite number. |
| `fx.ATAN` | formula | Return the inverse tangent of a finite number. |
| `fx.ATAN2` | formula | Return the quadrant-aware angle for x and y coordinates; the origin returns #DIV/0!. |
| `fx.ATANH` | formula | Return the inverse hyperbolic tangent for finite inputs strictly between -1 and 1; boundary values fail as #NUM!. |
| `fx.AVERAGE` | formula | Average numeric values across arguments and ranges in the clean-room formula engine. |
| `fx.AVERAGEIF` | formula | Average values whose corresponding entries match case-insensitive comparison or wildcard criteria. |
| `fx.AVERAGEIFS` | formula | Average values where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria. |
| `fx.CEILING` | formula | Round a number up to the nearest significance. |
| `fx.CHOOSE` | formula | Select one scalar result from up to 254 ordered choices using a truncated 1-based index; invalid indexes and unsupported arity return #VALUE!. |
| `fx.CHOOSECOLS` | formula | Select and reorder one or more 1-based or negative column indexes from an array. |
| `fx.CHOOSEROWS` | formula | Select and reorder one or more 1-based or negative row indexes from an array. |
| `fx.CLEAN` | formula | Remove ASCII C0 control characters from one bounded scalar text value while preserving Unicode text and other controls; overlong, error, or multi-cell inputs fail closed. |
| `fx.COLUMN` | formula | Return the 1-based column of the current formula cell or one explicit single-cell reference; ranges, spills, computed matrices, and invalid arity fail closed as #VALUE!. |
| `fx.COLUMNS` | formula | Return the column count of one bounded rectangular reference or dynamic spill. |
| `fx.COMBIN` | formula | Return the number of combinations for two non-negative bounded integer arguments. |
| `fx.COMBINA` | formula | Return combinations with repetition for two non-negative bounded integer arguments. |
| `fx.CONCAT` | formula | Concatenate text values and ranges. |
| `fx.CORREL` | formula | Return the Pearson correlation coefficient for two same-length bounded sources, ignoring positions where either value is nonnumeric; mismatched lengths return #N/A and empty or zero-variance pairs return #DIV/0!. |
| `fx.COS` | formula | Return the cosine of a finite radian value. |
| `fx.COSH` | formula | Return the hyperbolic cosine of a finite number; overflow fails as #NUM!. |
| `fx.COUNT` | formula | Count numeric values across arguments and ranges. |
| `fx.COUNTA` | formula | Count non-empty values across arguments and ranges, including text, logical values, errors, and empty-text formula results. |
| `fx.COUNTBLANK` | formula | Count blank cells and formula results that are empty text in one range. |
| `fx.COUNTIF` | formula | Count values using case-insensitive numeric/text criteria and Excel ?, *, and ~ wildcard semantics. |
| `fx.COUNTIFS` | formula | Count rows where multiple criteria ranges of the same size match case-insensitive comparison or wildcard criteria. |
| `fx.COVARIANCE.P` | formula | Calculate population covariance for two same-length bounded sources with pairwise numeric filtering; mismatched lengths return #N/A and no numeric pairs returns #DIV/0!. |
| `fx.COVARIANCE.S` | formula | Estimate sample covariance for two same-length bounded sources with pairwise numeric filtering; mismatched lengths return #N/A and fewer than two numeric pairs returns #DIV/0!. |
| `fx.CUMIPMT` | formula | Calculate cumulative interest paid across a bounded inclusive range of constant-payment loan periods. |
| `fx.CUMPRINC` | formula | Calculate cumulative principal paid across a bounded inclusive range of constant-payment loan periods. |
| `fx.DATE` | formula | Return an Excel serial in the workbook's 1900 or 1904 date system, with overflow and 1900 serial-60 compatibility. |
| `fx.DATEVALUE` | formula | Convert deterministic ISO or English month-name date text to a serial in the workbook's 1900 or 1904 date system; ambiguous locale-numeric dates return #VALUE!. |
| `fx.DAY` | formula | Return the day component of a serial in the workbook's date system, including 1900 compatibility serial 60. |
| `fx.DAYS` | formula | Return the whole-day difference between two Excel date serials. |
| `fx.DAYS360` | formula | Return the accounting day count between two valid Excel date serials using the U.S. NASD 30/360 method by default or the European 30E/360 method when the optional logical method is TRUE. Invalid dates, method text, and arity fail explicitly. |
| `fx.DB` | formula | Calculate one fixed-declining-balance depreciation period with an optional first-year month count. |
| `fx.DDB` | formula | Calculate one double-declining-balance depreciation period with an optional positive factor. |
| `fx.DEGREES` | formula | Convert finite radians to degrees with an explicit non-finite-result guard. |
| `fx.DROP` | formula | Drop rows and optional columns from the start or end of an array and spill the remainder. |
| `fx.EDATE` | formula | Shift a serial date by whole months and clamp the day to the target month end. |
| `fx.EOMONTH` | formula | Return the final date serial of a month offset from a start date. |
| `fx.EVEN` | formula | Round a finite number away from zero to the next even integer. |
| `fx.EXACT` | formula | Compare two bounded scalar text values with case-sensitive equality; multi-cell sources and overlong values fail closed. |
| `fx.EXP` | formula | Return e raised to a finite number; overflow fails as #NUM! instead of leaking Infinity. |
| `fx.EXPAND` | formula | Expand an array to requested row and column dimensions with optional padding. |
| `fx.FACT` | formula | Return the factorial of a non-negative integer through the finite 170! boundary. |
| `fx.FACTDOUBLE` | formula | Return the double factorial of a non-negative integer through the bounded finite range. |
| `fx.FALSE` | formula | Return the logical value FALSE with no arguments; supplied arguments fail as #VALUE!. |
| `fx.FILTER` | formula | Filter rows from a source range with a boolean or comparison include array and spill the matching rows. |
| `fx.FIND` | formula | Return the 1-based position of a case-sensitive literal text sequence. |
| `fx.FLOOR` | formula | Round a number down to the nearest significance. |
| `fx.FORECAST.LINEAR` | formula | Predict one y value from one bounded scalar x and aligned known-y/known-x sources using the shared stable linear fit; nonnumeric x returns #VALUE!, source mismatch or no pairs returns #N/A, and zero x variance returns #DIV/0!. |
| `fx.FORMULATEXT` | formula | Return the stored formula text for one explicit single-cell reference, #N/A when that cell has no formula, and #VALUE! for ranges, computed matrices, spills, or invalid input. |
| `fx.FV` | formula | Calculate the future value of a finite constant-payment stream from rate, term, payment, optional present value, and payment timing. |
| `fx.GCD` | formula | Return the greatest common divisor of bounded integer arguments and ranges; unsafe integer results fail closed as #NUM!. |
| `fx.GROWTH` | formula | Return a bounded single-variable exponential prediction dynamic array for y=b*m^x with the same row or column shape as new-x. Known-y must be positive; x arguments may be omitted, const may force b=1, and constant known-x is removed. Overflow, multivariable or two-dimensional inputs, nonnumeric new-x positions, and mismatched known source shapes fail closed. |
| `fx.HLOOKUP` | formula | Look up one scalar in the first row of a nonempty rectangular range of at most 10,000 cells; FALSE/0 performs an exact, wildcard-aware lookup, while TRUE/1 or omission requires a proven ascending homogeneous numeric or text key row and returns the greatest matching-or-lower key. Invalid table/mode/index inputs and unproven ordering return #VALUE!, while an out-of-range return-row index returns #REF!. |
| `fx.HOUR` | formula | Return the 0 through 23 hour component from a nonnegative serial or supported time text. |
| `fx.HSTACK` | formula | Append arrays horizontally, padding shorter arrays with #N/A to the maximum row count. |
| `fx.IF` | formula | Return one value when a condition is true and another when false. |
| `fx.IFERROR` | formula | Return a fallback value when an expression evaluates to a formula error. |
| `fx.IFNA` | formula | Return a fallback only when an expression evaluates to #N/A; preserve every other result or error. |
| `fx.IFS` | formula | Evaluate condition/value pairs in order and return the first matching value, or #N/A when no condition matches. |
| `fx.INDEX` | formula | Select one value from a nonempty rectangular range of at most 10,000 cells with host-compatible row and optional column selectors, preserving an error-valued selector such as a failed MATCH. Only the documented 2- or 3-argument array/range form is modeled; missing or extra selectors and oversized ranges return #VALUE!, while a missing or out-of-range source cell returns #REF!. |
| `fx.INT` | formula | Round a number down to the nearest integer. |
| `fx.INTERCEPT` | formula | Return the y-axis intercept for the same bounded source-aware linear regression profile as SLOPE; empty or mismatched sources return #N/A and zero x variance returns #DIV/0!. |
| `fx.IPMT` | formula | Calculate the interest component of one constant-payment loan period from finite rate, period, term, present value, optional future value, and payment-timing inputs. |
| `fx.IRR` | formula | Return a bounded-convergence periodic return rate for a finite cash-flow vector. |
| `fx.ISBLANK` | formula | Return TRUE when a referenced value is empty. |
| `fx.ISERR` | formula | Return TRUE for recognized formula errors other than #N/A. |
| `fx.ISERROR` | formula | Return TRUE when a value is any recognized formula error. |
| `fx.ISFORMULA` | formula | Return TRUE when one explicit single-cell reference contains a formula, FALSE when the cell is not formula-backed, and #VALUE! for ranges, computed matrices, spills, or invalid input. |
| `fx.ISLOGICAL` | formula | Return TRUE when a value is a logical TRUE or FALSE. |
| `fx.ISNA` | formula | Return TRUE only when a value is the #N/A error. |
| `fx.ISNONTEXT` | formula | Return TRUE when a value is not text, including blank, logical, numeric, and error values. |
| `fx.ISNUMBER` | formula | Return TRUE when a value is numeric. |
| `fx.ISOWEEKNUM` | formula | Return the ISO 8601 week number for one valid Excel date serial in the workbook's 1900 or 1904 date system. OfficeKit Codec owns the required _xlfn.ISOWEEKNUM package spelling. |
| `fx.ISREF` | formula | Return TRUE only for a direct A1, defined-name, or spill reference expression; computed values and functions return FALSE, while invalid arity fails closed as #VALUE!. |
| `fx.ISTEXT` | formula | Return TRUE when a value is text and not a formula error. |
| `fx.LARGE` | formula | Return the k-th largest numeric value in an array or range. |
| `fx.LCM` | formula | Return the least common multiple of bounded integer arguments and ranges; zero inputs return zero and unsafe overflow returns #NUM!. |
| `fx.LEFT` | formula | Return up to 32,767 Unicode characters from the start of one bounded scalar text value; num_chars defaults to 1 and invalid or multi-cell inputs fail closed. |
| `fx.LEN` | formula | Return the Unicode code-point length of one bounded scalar text value; overlong, error, or multi-cell inputs fail closed. |
| `fx.LET` | formula | Bind up to 16 scalar local names from left to right and evaluate a final scalar expression; write the public formula exactly as Excel displays it, while OfficeKit Codec owns scoped _xlfn.LET/_xlpm package spelling. Invalid names, array-valued bindings, and missing arguments fail closed as #VALUE!. |
| `fx.LINEST` | formula | Return a bounded single-variable least-squares dynamic array: 1x2 slope/intercept by default or the documented 5x2 coefficient, error, R-squared, F/df, and regression/residual statistics matrix when stats is TRUE. Known-x may be omitted, const may force a zero intercept, constant known-x is removed, and mismatched shapes return #N/A; multivariable inputs and array constants remain unsupported. |
| `fx.LN` | formula | Return the natural logarithm of a positive finite number; non-positive inputs fail as #NUM!. |
| `fx.LOG` | formula | Return a logarithm for a positive number and positive base other than one; the base defaults to 10 and invalid domains fail as #NUM!. |
| `fx.LOG10` | formula | Return the base-10 logarithm of a positive finite number. |
| `fx.LOGEST` | formula | Return a bounded single-variable exponential regression dynamic array for y=b*m^x: 1x2 multiplier/base by default or a 5x2 matrix whose remaining diagnostics describe the natural-log regression when stats is TRUE. Known-y must be positive; known-x may be omitted, const may force b=1, and constant known-x is removed. Mismatched shapes, multivariable inputs, and array constants fail closed. |
| `fx.LOOKUP` | formula | Return the result aligned with the greatest ascending homogeneous numeric or text key less than or equal to one scalar. The bounded vector form accepts one optional same-length result vector; the array form searches its first column when square or taller and its first row when wider, then returns from the last column or row. Unproven ordering, mixed keys, mismatched vectors, two-dimensional vector arguments, and sources above 10,000 cells fail as #VALUE!. |
| `fx.LOWER` | formula | Convert text to lowercase. |
| `fx.MATCH` | formula | Return a 1-based lookup position in one row or column vector of 1 through 10,000 cells. Exact 0 matching is wildcard-aware; default/1 approximate matching requires a proven ascending homogeneous numeric or text vector and returns the greatest matching-or-lower key, while -1 requires proven descending order and returns the smallest matching-or-higher key. Two-dimensional, oversized, mixed, unordered, or invalid-mode inputs return #VALUE!. |
| `fx.MAX` | formula | Return the maximum numeric value across arguments and ranges. |
| `fx.MAXIFS` | formula | Return the largest numeric value where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria. |
| `fx.MEDIAN` | formula | Return the middle numeric value, or the average of the two middle values, across arguments and ranges. |
| `fx.MID` | formula | Return a bounded Unicode slice from one scalar text value using a 1-based start and non-negative character count; invalid or multi-cell inputs fail closed. |
| `fx.MIN` | formula | Return the minimum numeric value across arguments and ranges. |
| `fx.MINIFS` | formula | Return the smallest numeric value where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria. |
| `fx.MINUTE` | formula | Return the 0 through 59 minute component from a nonnegative serial or supported time text. |
| `fx.MIRR` | formula | Calculate a modified periodic internal rate of return using distinct finance and reinvestment rates for a finite cash-flow vector. |
| `fx.MOD` | formula | Return the remainder after division, preserving the divisor sign and returning #DIV/0! for a zero divisor. |
| `fx.MODE.MULT` | formula | Return every numeric value tied for the highest frequency as an ascending vertical spill; if no value repeats, return #N/A instead of synthesizing modes. |
| `fx.MODE.SNGL` | formula | Return the most frequently occurring numeric value, or #N/A when no value repeats. |
| `fx.MONTH` | formula | Return the month component of a serial in the workbook's 1900 or 1904 date system. |
| `fx.MROUND` | formula | Round a finite number to the nearest multiple with explicit zero-multiple and sign checks. |
| `fx.N` | formula | Return a bounded numeric coercion: numbers and date serials unchanged, TRUE/FALSE as 1/0, text or blank as 0, and formula errors propagated; multi-cell or matrix input fails closed as #VALUE!. |
| `fx.NA` | formula | Return the #N/A error value to mark unavailable data explicitly. |
| `fx.NETWORKDAYS` | formula | Count Monday-through-Friday dates inclusively between two serial dates, excluding optional holidays. |
| `fx.NETWORKDAYS.INTL` | formula | Count inclusive workdays with a numbered or Monday-first seven-character custom weekend and optional holidays. |
| `fx.NOT` | formula | Reverse the truth value of a condition. |
| `fx.NPER` | formula | Solve the finite payment-period count from rate, payment, present value, optional future value, and payment timing. |
| `fx.NPV` | formula | Discount a finite periodic cash-flow vector beginning one period after the present value date. |
| `fx.ODD` | formula | Round a finite number away from zero to the next odd integer. |
| `fx.OR` | formula | Return TRUE when any condition is true. |
| `fx.PERCENTILE.EXC` | formula | Return an exclusive percentile from a bounded numeric range using rank k*(n+1); k must be strictly between 0 and 1, and endpoints that cannot be interpolated return #NUM!. |
| `fx.PERCENTILE.INC` | formula | Return an inclusive percentile from a bounded array or range; k must be from 0 through 1 and the result uses linear interpolation, while nonnumeric reference values are ignored, formula errors propagate, and an empty numeric set fails as #NUM!. |
| `fx.PI` | formula | Return the deterministic mathematical constant π; arguments are rejected rather than ignored. |
| `fx.PMT` | formula | Calculate a constant-period loan payment from finite rate, term, present value, optional future value, and payment-timing inputs. |
| `fx.POWER` | formula | Raise a finite base to a finite exponent; non-finite results fail as #NUM! rather than leaking JavaScript Infinity or NaN. |
| `fx.PPMT` | formula | Calculate the principal component of one constant-payment loan period using the same bounded inputs as IPMT. |
| `fx.PRODUCT` | formula | Multiply numeric values across arguments and bounded ranges; formula errors propagate and empty invocation returns #VALUE!. |
| `fx.PV` | formula | Calculate the present value of a finite constant-payment stream from rate, term, payment, optional future value, and payment timing. |
| `fx.QUARTILE.EXC` | formula | Return an exclusive first, second, or third quartile from a bounded numeric range; the selector is truncated and indexes outside 1 through 3 return #NUM!. |
| `fx.QUARTILE.INC` | formula | Return an inclusive quartile from a bounded array or range; the quartile index must be an integer from 0 through 4 and the result uses linear interpolation, while nonnumeric reference values are ignored, formula errors propagate, and an empty numeric set fails as #NUM!. |
| `fx.QUOTIENT` | formula | Return the integer portion of a division result, truncating toward zero and returning #DIV/0! for a zero divisor. |
| `fx.RADIANS` | formula | Convert finite degrees to radians with an explicit non-finite-result guard. |
| `fx.RANK.AVG` | formula | Return a number's rank in a bounded numeric range and average the occupied positions when values tie; a number absent from the range returns #N/A. |
| `fx.RANK.EQ` | formula | Return a number's equal rank in a numeric range, descending by default or ascending when order is nonzero. |
| `fx.RATE` | formula | Solve a bounded periodic interest rate from an integer payment term, payment, present value, optional future value, payment timing, and optional guess. |
| `fx.REPLACE` | formula | Replace a bounded scalar text span using 1-based character and non-negative length arguments; invalid positions, matrices, and overlong results fail closed. |
| `fx.REPT` | formula | Repeat one bounded scalar text value an integer number of times, with a 32,767-character result budget. |
| `fx.RIGHT` | formula | Return up to 32,767 Unicode characters from the end of one bounded scalar text value; num_chars defaults to 1 and invalid or multi-cell inputs fail closed. |
| `fx.ROUND` | formula | Round a numeric value to decimal places or, with negative digits, positions left of the decimal point. |
| `fx.ROUNDDOWN` | formula | Round a numeric value toward zero at the requested positive or negative digit position. |
| `fx.ROUNDUP` | formula | Round a numeric value away from zero at the requested positive or negative digit position. |
| `fx.ROW` | formula | Return the 1-based row of the current formula cell or one explicit single-cell reference; ranges, spills, computed matrices, and invalid arity fail closed as #VALUE!. |
| `fx.ROWS` | formula | Return the row count of one bounded rectangular reference or dynamic spill. |
| `fx.RSQ` | formula | Return the square of Pearson correlation for aligned known-y and known-x sources; positions are pairwise filtered, length mismatch or no pairs returns #N/A, and fewer than two or zero-variance pairs returns #DIV/0!. |
| `fx.SEARCH` | formula | Return the 1-based position of case-insensitive text, supporting Excel ?, *, and ~ wildcard syntax. |
| `fx.SECOND` | formula | Return the 0 through 59 second component from a nonnegative serial or supported time text. |
| `fx.SEQUENCE` | formula | Return a dynamic array sequence that spills into neighboring cells in the clean-room formula engine. |
| `fx.SHEET` | formula | Return the 1-based OfficeKit worksheet position for the current sheet or one validated single-sheet cell/range, workbook defined name, table, or sheet-name string. Missing sheet-name strings return #N/A; invalid references, nonreference values, 3D spans, and extra arguments fail explicitly. Chart, macro, and dialog sheets are outside the OfficeKit workbook model. |
| `fx.SHEETS` | formula | Return the total number of OfficeKit worksheets, including hidden worksheets, or 1 for one validated single-sheet cell/range, workbook defined name, table, or sheet-name string. Invalid references, nonreference values, 3D spans, and extra arguments fail explicitly; chart, macro, and dialog sheets are not modeled. |
| `fx.SIGN` | formula | Return -1, 0, or 1 according to the sign of a finite numeric value. |
| `fx.SIN` | formula | Return the sine of a finite radian value. |
| `fx.SINH` | formula | Return the hyperbolic sine of a finite number; overflow fails as #NUM!. |
| `fx.SLN` | formula | Calculate straight-line depreciation from cost, salvage value, and useful life. |
| `fx.SLOPE` | formula | Return the least-squares slope for aligned known-y and known-x sources using stable pair moments; nonnumeric reference positions are ignored together, mismatched lengths return #N/A, and zero x variance returns #DIV/0!. |
| `fx.SMALL` | formula | Return the k-th smallest numeric value in an array or range. |
| `fx.SORT` | formula | Sort a range by a 1-based column index and spill the sorted rows. |
| `fx.SQRT` | formula | Return the non-negative square root of a finite number; negative inputs return #NUM!. |
| `fx.STDEV.P` | formula | Calculate population standard deviation with a numerically stable bounded calculation; references ignore text, logical, blank, and error cells, while direct logical and numeric-text arguments are counted, direct errors propagate, and an empty numeric set returns #DIV/0!. |
| `fx.STDEV.S` | formula | Estimate sample standard deviation with a numerically stable bounded calculation; references ignore text, logical, blank, and error cells, while direct logical and numeric-text arguments are counted, direct errors propagate, and fewer than two numbers returns #DIV/0!. |
| `fx.STEYX` | formula | Return the standard error of predicted y values for a bounded linear regression; pairwise source semantics match SLOPE, fewer than three numeric pairs returns #DIV/0!, and mismatched source lengths return #N/A. |
| `fx.SUBSTITUTE` | formula | Replace all or one 1-based occurrence of a literal substring in bounded scalar text; matching is case-sensitive and empty search text fails closed. |
| `fx.SUM` | formula | Sum numeric values across arguments and ranges. |
| `fx.SUMIF` | formula | Sum corresponding values using case-insensitive numeric/text criteria and Excel ?, *, and ~ wildcards. |
| `fx.SUMIFS` | formula | Sum values where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria. |
| `fx.SUMPRODUCT` | formula | Multiply corresponding numeric values in equally sized arrays and return the sum of those products; bounded same-shape direct-range predicate factors support comparisons, unary signs, and scalar arithmetic within SUMPRODUCT. |
| `fx.SUMSQ` | formula | Sum the squares of numeric values across arguments and bounded ranges; overflow returns #NUM! and formula errors propagate. |
| `fx.SWITCH` | formula | Match an expression against ordered value/result pairs and return an optional default or #N/A when no value matches. |
| `fx.SYD` | formula | Calculate sum-of-years-digits depreciation for one bounded useful-life period. |
| `fx.T` | formula | Return text unchanged, convert non-text scalars to empty text, and propagate formula errors; multi-cell or matrix input fails closed as #VALUE!. |
| `fx.TAKE` | formula | Take rows and optional columns from the start or end of an array and spill the result. |
| `fx.TAN` | formula | Return the tangent of a finite radian value. |
| `fx.TANH` | formula | Return the hyperbolic tangent of a finite number. |
| `fx.TEXT` | formula | Format an Excel serial date as text with the bounded yyyy, yy, m/mm/mmm/mmmm, and d/dd token profile and literal separators. |
| `fx.TEXTAFTER` | formula | Return the text after a delimiter occurrence, with bounded positive/negative instance selection, case mode, end matching, and an explicit not-found result. |
| `fx.TEXTBEFORE` | formula | Return the text before a delimiter occurrence, with bounded positive/negative instance selection, case mode, end matching, and an explicit not-found result. |
| `fx.TEXTJOIN` | formula | Join text values with a delimiter and optional empty-value skipping. |
| `fx.TEXTSPLIT` | formula | Split one scalar text value into a bounded spilled matrix by column and optional row delimiters, with empty-item skipping, case mode, and padding; multi-cell sources, empty delimiters, and oversized results fail closed. |
| `fx.TIME` | formula | Return a time fraction from hour, minute, and second values from 0 through 32767, carrying overflow and wrapping at 24 hours. |
| `fx.TIMEVALUE` | formula | Convert deterministic 12-hour or 24-hour time text, optionally following date text, to a fraction of one day. |
| `fx.TOCOL` | formula | Flatten an array into one spilled column, optionally ignoring blanks or errors and scanning by column. |
| `fx.TOROW` | formula | Flatten an array into one spilled row, optionally ignoring blanks or errors and scanning by column. |
| `fx.TRANSPOSE` | formula | Transpose a source range into a spilled dynamic array with spillRange/spillValues inspect metadata. |
| `fx.TREND` | formula | Return a bounded single-variable linear prediction dynamic array with the same row or column shape as new-x. Known-x and new-x may be omitted, const may force a zero intercept, and a constant known-x column is removed; multivariable or two-dimensional inputs, nonnumeric new-x positions, and mismatched known source shapes fail closed. |
| `fx.TRIM` | formula | Trim leading/trailing whitespace and collapse internal whitespace. |
| `fx.TRIMMEAN` | formula | Average a bounded numeric range after removing an even number of observations symmetrically from both tails; the requested percentage must be from 0 through 1. |
| `fx.TRUE` | formula | Return the logical value TRUE with no arguments; supplied arguments fail as #VALUE!. |
| `fx.TRUNC` | formula | Truncate a finite number toward zero at an optional decimal position without rounding. |
| `fx.TYPE` | formula | Return Excel type codes 1 for numbers or blank, 2 for text, 4 for logical, 16 for errors, or 64 for arrays and multi-cell references; bounded spill/reference detection is explicit and invalid arity fails closed. |
| `fx.UNICHAR` | formula | Return one Unicode scalar character for an integer from 1 through 1,114,111; surrogate values, invalid ranges, errors, and multi-cell inputs fail closed. |
| `fx.UNICODE` | formula | Return the Unicode code point of the first character in one bounded scalar text value; empty, overlong, error, or multi-cell inputs fail closed. |
| `fx.UNIQUE` | formula | Return unique rows from a range as a spilled dynamic array. |
| `fx.UPPER` | formula | Convert text to uppercase. |
| `fx.VALUE` | formula | Convert deterministic ASCII numeric text with optional grouping, scientific notation, accounting parentheses, or percent suffix to a number. |
| `fx.VAR.P` | formula | Calculate population variance with a numerically stable bounded calculation; references ignore text, logical, blank, and error cells, while direct logical and numeric-text arguments are counted, direct errors propagate, and an empty numeric set returns #DIV/0!. |
| `fx.VAR.S` | formula | Estimate sample variance with a numerically stable bounded calculation; references ignore text, logical, blank, and error cells, while direct logical and numeric-text arguments are counted, direct errors propagate, and fewer than two numbers returns #DIV/0!. |
| `fx.VLOOKUP` | formula | Look up one scalar in the first column of a nonempty rectangular range of at most 10,000 cells; FALSE/0 performs an exact, wildcard-aware lookup, while TRUE/1 or omission requires a proven ascending homogeneous numeric or text key column and returns the greatest matching-or-lower key. Invalid table/mode/index inputs and unproven ordering return #VALUE!, while an out-of-range return-column index returns #REF!. |
| `fx.VSTACK` | formula | Append arrays vertically, padding narrower arrays with #N/A to the maximum column count. |
| `fx.WEEKDAY` | formula | Return a weekday number for Excel return types 1, 2, 3, and 11 through 17. |
| `fx.WEEKNUM` | formula | Return a calendar week number under Excel system 1 for return types 1, 2, and 11 through 17, or the ISO 8601 week number for return type 21; invalid dates and return types fail explicitly. |
| `fx.WORKDAY` | formula | Move forward or backward by working days while skipping weekends and optional holidays. |
| `fx.WORKDAY.INTL` | formula | Move by workdays using a numbered or Monday-first seven-character custom weekend and optional holidays. |
| `fx.WRAPCOLS` | formula | Wrap a one-dimensional vector into columns of a requested height, padding the final column when needed. |
| `fx.WRAPROWS` | formula | Wrap a one-dimensional vector into rows of a requested width, padding the final row when needed. |
| `fx.XIRR` | formula | Return a bounded-convergence annualized return rate for date-aligned finite cash flows using a 365-day year. |
| `fx.XLOOKUP` | formula | Look up one scalar in same-shaped one-dimensional row or column vectors of 1 through 10,000 cells; exact, next-smaller, next-larger, wildcard, and first/last linear search modes are modeled, while binary-search modes and mismatched or two-dimensional ranges fail as #VALUE!. |
| `fx.XMATCH` | formula | Return a 1-based lookup position in one row or column vector of 1 through 10,000 cells, with exact, next-smaller, next-larger, wildcard, and forward or reverse linear search modes; two-dimensional, oversized, and binary-search inputs fail as #VALUE!. |
| `fx.XNPV` | formula | Discount date-aligned finite cash flows by actual day offsets from the first date using a 365-day year. |
| `fx.XOR` | formula | Return TRUE when an odd number of up to 255 scalar conditions are true; array-valued logical arguments remain outside the bounded evaluator. |
| `fx.YEAR` | formula | Return the year component of a serial in the workbook's 1900 or 1904 date system. |
| `importXlsxWithOfficeKit` | api | Import XLSX bytes through OfficeKit with editable core cells, formulas, styles, ordinary tables, PNG/JPEG pictures, validation, conditional formatting, threaded-comment roots with direct replies, bar/line/pie/area/doughnut charts, marker-only numeric-X/Y scatter charts, bounded numeric-X/Y/positive-Size bubble charts, and recognized PivotTables with exact item or absolute whole-day date filters. Imported data-table and dynamic-array topology is source-bound and read-only. A recognized source-bound QueryTable can only disable automatic refresh through table.setQueryRefreshPolicy; a recognized connection can only disable an explicit on-load refresh through workbook.disableConnectionRefreshOnLoad; a recognized uniquely owned Pivot cache can only disable an explicit on-load refresh through pivot.disableRefreshOnLoad; commands, fields, sorts, topology, non-marker scatter styles, noncanonical bubble profiles, nested/branched replies, mentions, other Pivot configuration/data/output, non-reversible sparkline graphs, and other advanced package content remain source-bound and read-only. |
| `invokeOfficeKit` | api | Advanced experimental byte-boundary API for invoking the public OfficeKit codec protocol with generated wire-message objects. |
| `officeKitStatus` | api | Lazily initialize the bundled OfficeKit NativeAOT codec and report its backend, target, transport, protocol, assembly, and integrity manifest. |
| `pivot.disableRefreshOnLoad` | api | On one recognized imported PivotTable with a uniquely owned cache and explicit refreshOnLoad=true, set only that cache root switch to false while preserving the complete Pivot graph and every other cache attribute. |
| `pivot.sourceCapabilities` | api | Inspect whether a PivotTable is source-bound and whether its uniquely owned imported cache can receive the one-way refreshOnLoad hardening operation. |
| `range.clear` | api | Clear range contents, formats, or both without silently changing validations, dimensions, or other package graphs. |
| `range.conditionalFormats.add` | api | Add a conditional formatting rule; cellIs/expression/containsText/colorScale plus standard dataBar/iconSet rules cross the public model and OfficeKit, with computedStyle inspect records, layout JSON visuals, SVG preview, and native XLSX rendering. |
| `range.copyFrom` | api | Copy values, formulas, or complete cells from an equally sized or evenly tiling source range with relative A1 translation. |
| `range.copyTo` | api | Copy this range to an equally sized or evenly tiled destination range. |
| `range.dataValidation` | api | Assign a list, whole, decimal, date, time, text-length, or custom-formula validation rule to a range, including bounded input prompts, error alerts, blank policy, and intuitive list-arrow visibility; use sheet.dataValidations.add({ range, rule }) for the collection form. |
| `range.displayFormulas` | api | Read displayed A1 formulas, including the anchor formula projected across non-editable dynamic-array or legacy-array result cells. |
| `range.fillDown` | api | Copy top-row contents and formatting down the range while translating relative A1 formula references. |
| `range.fillRight` | api | Copy left-column contents and formatting right across the range while translating relative A1 formula references. |
| `range.format` | api | Assign cell styles, symbolic theme/tint/indexed colors, patterned fills, native dimensions, pixel sizing, and hidden axes through a live range format facade. |
| `range.format.autofitColumns` | api | Measure displayed range values deterministically and set native best-fit widths on each selected column. |
| `range.format.autofitRows` | api | Measure explicit/wrapped range text deterministically and set native custom heights on each selected row. |
| `range.formulaInfos` | api | Read per-cell stored/projected formula metadata with editability, spill/array source, anchor, and reference evidence. |
| `range.formulasR1C1` | api | Read or assign R1C1 formulas relative to each target cell while storing canonical A1 formulas. |
| `range.getCell` | api | Select one zero-based cell relative to the current range. |
| `range.getColumn` | api | Select one zero-based column relative to the current range. |
| `range.getCurrentRegion` | api | Expand to the contiguous data region bounded by fully blank rows and columns. |
| `range.getRangeByIndexes` | api | Select a bounded zero-based subrange relative to the current range. |
| `range.getRow` | api | Select one zero-based row relative to the current range. |
| `range.merge` | api | Merge the target range as one region or as separate row-wise regions when across=true. |
| `range.offset` | api | Return an equally sized range shifted by zero-based row and column offsets, rejecting worksheet overflow. |
| `range.resize` | api | Return a range at the same upper-left cell with explicit positive row and column counts. |
| `range.setNumberFormat` | api | Assign one number format or an evenly tiling matrix of Excel-invariant number-format codes. |
| `range.unmerge` | api | Remove merged regions intersecting the target range. |
| `range.write` | api | Write a mixed matrix or one explicit values/formulas/formulasR1C1 payload from the range anchor and return the actual written range. |
| `range.writeValues` | api | Write a one- or two-dimensional value matrix from the range anchor. |
| `sheet.charts.add` | api | Create an inspectable worksheet chart from a range or config; setData(range) infers category series, scatter per-series numeric xValues/xFormula plus y values/formula, or one exact X/Y/positive-Size bubble series. series.fill sets an explicit #RRGGBB solid color, series.line sets bounded RGB color/dash/width (series.stroke is an alias), line/scatter markers set direct symbol/size/RGB fill/bounded outline semantics, lineOptions controls standard/stacked/percent-stacked grouping, smooth interpolation, and direct vary-colors behavior, dataLabels controls plot-level value/category/series-name visibility and bounded position, and xAxis/yAxis configure primary titles, formats, intervals, and linear value bounds. Bar and line series accept up to 16 bounded native linear, exponential, logarithmic, power, polynomial, or moving-average trendlines plus one bounded native errorBars projection with fixed/percentage/standard-deviation/standard-error/custom semantics, one-/two-sided values, optional XLSX formula caches, cap policy, and bounded RGB line. Imported trendline count and error-bar presence are fixed; unsupported labels/extensions/unknown children/complex lines remain source-owned. Marker-only scatter rejects series.line/stroke and writes an explicit native no-fill series outline; use marker.line for marker borders. Bubble charts use two numeric axes and reject ambiguous range shortcuts or nonpositive sizes. |
| `sheet.dataTables.__getDefinitions` | api | Return defensive inspectable definitions for the worksheet's canonical What-If data tables, including result range, native anchor, inputs, orientation, and display formula. |
| `sheet.dataTables.add` | api | Create a canonical native Excel What-If data table from a rectangular formula/input grid and one row input, one column input, or both. Excel or another compatible host calculates the result values; the JavaScript evaluator does not simulate TABLE. |
| `sheet.images.add` | api | Create an inspectable worksheet image from a data URL, URI, or prompt with one-cell, two-cell, or absolute pixel geometry plus optional percentage crop, bounded grayscale/luminance/opacity effects, rotation, and horizontal/vertical flips. |
| `sheet.pivotTables.add` | api | Create a native bounded XLSX PivotTable with derived cached output, cache records, exact axis-item filters, and absolute whole-day date conditions. Relative-clock and sub-day filters remain model-only. Recognized imports are hash-bound and read-only except the separately verified refresh-on-load hardening primitive. |
| `sheet.sparklineGroups.add` | api | Create standard Office 2010 line/column/stacked sparkline groups for inspect, SVG preview, and OfficeKit XLSX export. Source-free groups use reversible one-dimensional target/source mappings; recognized imported groups support fixed-topology semantic edits while unsupported native graphs remain source-bound. |
| `sheet.tables.add` | api | Create an ordinary worksheet table over an A1 range with headers, columns, totals metadata, style, and bounded filtering/sorting. QueryTable bindings cannot be authored; recognized imported bindings expose only table.setQueryRefreshPolicy for one-way automatic-refresh hardening, while all other QueryTable edits fail closed. |
| `SpreadsheetFile.exportCsv` | api | Export one worksheet or range as UTF-8 CSV, using calculated values unless formula output is explicitly requested. |
| `SpreadsheetFile.exportDelimited` | api | Serialize one workbook sheet/range as bounded CSV/TSV text with calculated-value defaults and RFC-style quoting. |
| `SpreadsheetFile.exportTsv` | api | Export one worksheet or range as UTF-8 tab-separated text with RFC-style quoting where needed. |
| `SpreadsheetFile.exportXlsx` | api | Serialize a Workbook facade through the single bundled OfficeKit codec. |
| `SpreadsheetFile.importCsv` | api | Import UTF-8 CSV bytes into an editable Workbook through the bounded delimited parser. |
| `SpreadsheetFile.importDelimited` | api | Parse bounded RFC-style CSV/TSV bytes into an editable Workbook, including quoted delimiters, escaped quotes, and embedded newlines. |
| `SpreadsheetFile.importTsv` | api | Import UTF-8 tab-separated bytes into an editable Workbook through the bounded delimited parser. |
| `SpreadsheetFile.importXlsx` | api | Load XLSX through the single bundled OfficeKit codec into an editable Workbook facade. |
| `SpreadsheetFile.inspectDelimited` | api | Inspect bounded CSV/TSV bytes as file/row records with dimensions, delimiter, quoting, and formula-like cell evidence. |
| `SpreadsheetFile.inspectXlsx` | api | Inspect bounded XLSX parts, content types, the required workbook/root officeDocument relationship, and namespace-aware source XML r:id/r:embed/r:link references after raw-input, part-count, decompression, and optional compression-ratio budgets; verifyCrc32 additionally checks ZIP entry CRCs. |
| `SpreadsheetFile.patchXlsx` | api | Apply path-validated XLSX part patches, build worksheet/table/drawing/image/chart/pivot source references, and atomically reject dangling content types or relationships. |
| `table.setQueryRefreshPolicy` | api | On one recognized imported QueryTable, monotonically disable automatic refresh without changing its connection, command, fields, sort, refresh history, or topology. |
| `thread.addReply` | api | Append a direct reply to an Office threaded-comment root with independent author/person/date/done metadata. Nested or branched reply graphs and mentions fail closed. |
| `workbook.auditAccessibility` | api | Audit worksheet images and charts for explicit meaningful/decorative classification and non-visible xdr:cNvPr title/description coverage. Native reading order and broader worksheet semantics remain manual checks; the report never claims Excel Accessibility Checker, WCAG, or PDF conformance. |
| `workbook.comments.addThread` | api | Create one root Office threaded comment per thread with GUID/person metadata, date, and resolved state; attach bounded direct replies with thread.addReply(). |
| `workbook.connections` | api | Inspect bounded non-secret metadata for imported database connections. Connections are source-bound; the sole mutation is workbook.disableConnectionRefreshOnLoad(connectionId) for an explicit imported refreshOnLoad=true value. |
| `Workbook.create` | api | Create an empty workbook with an explicit date system and optional native SpreadsheetML theme colors. |
| `workbook.definedNames.add` | api | Create a workbook or sheet-scoped defined name over an A1 range; exported as native workbook.xml definedName and usable in formulas such as SUM(RevenueData). |
| `workbook.disableConnectionRefreshOnLoad` | api | On one recognized imported connection with explicit refreshOnLoad=true, set that sole root switch to false without changing its command, credentials, topology, or any other connection state. |
| `workbook.fontFamilies` | api | Return a fresh sorted, case-insensitively deduplicated list of workbook default and explicit cell font families. |
| `workbook.formulaGraph` | api | Return a bounded dependency graph of formula nodes, edges, dependents, cycles, formula errors, and syntax-input/reference-budget refusals for workbook QA. |
| `workbook.inspect` | api | Emit bounded NDJSON records for workbook, connections, sheets, worksheet protections, tables, formulas, matches, comments, validations, conditional formats, and drawings; narrow with search/target anchors and shape fields with include/exclude. |
| `workbook.layoutJson` | api | Return workbook/worksheet layout JSON with cell, table, chart, image, sparkline, rule bounding boxes, and target/search context slicing. |
| `workbook.recalculate` | api | Recalculate bounded workbook formulas and dynamic-array spills, with dependency edges, cycles, errors, and syntax-input/reference-budget refusals. |
| `workbook.render` | api | Return a lightweight SVG preview for a sheet/range or layout JSON when called with { format: 'layout' }. |
| `workbook.resolve` | api | Resolve stable workbook, source-bound connection, worksheet, table, pivot, chart, image, sparkline, rule, comment, and defined-name IDs. |
| `workbook.setCalculation` | api | Set bounded workbook-level SpreadsheetML calculation mode, on-save/full-recalculation flags, iterative-calculation limits, and full-precision policy. |
| `workbook.setDateSystem` | api | Select the Excel 1900 or 1904 serial-date system for formula calculation and native workbookPr export. |
| `workbook.sharedArrayFormulas` | formula | Import and export bounded shared, legacy-array, and source-free XLDAPR dynamic-array formula metadata. Imported dynamic-array anchors remain source-bound and read-only; malformed or topology-changing edits fail closed. |
| `workbook.spillReferences` | formula | Use a direct or defined-name A1# reference to consume only an anchor's current, unblocked dynamic spill matrix. Supported range consumers and a direct re-spill read the verified matrix; scalar/general-vector coercion returns #VALUE!, non-spilling anchors return #REF!, and graph/trace record one spillReference edge to the anchor. |
| `workbook.structuredReferences` | formula | Evaluate Excel table references including sections, column ranges/unions, space intersections, escaped special-character headers, unqualified calculated-column references, and @/#This Row context while expanding exact table-cell precedents. |
| `workbook.trace` | api | Return a formula precedent tree and bounded NDJSON trace for a target cell, with circular references and syntax-input/reference-budget refusals flagged. |
| `workbook.verify` | api | Return bounded QA issues for source-bound connections, sheets, formulas (including syntax-input and reference-budget refusals), tables, charts, and comments. |
| `workbook.windows` | api | Access the ordered workbook-window collection; window 0 is the primary view used by legacy worksheet-selection APIs. |
| `workbook.windows.add` | api | Append an additional workbook window with its own active worksheet and selected tab group. |
| `workbook.worksheets.add` | api | Append an editable visible, hidden, or very-hidden worksheet with a stable name and ID. |
| `workbook.worksheets.getSelectedWorksheets` | api | Return the visible worksheet-tab group selected in the primary workbook window, in workbook order. |
| `workbook.worksheets.setActiveWorksheet` | api | Select the visible worksheet opened by default and used by workbook operations that omit an explicit sheet. |
| `workbook.worksheets.setSelectedWorksheets` | api | Select one or more visible worksheet tabs in the primary workbook window while retaining exactly one active worksheet. |
| `workbook.xlsxFormulaSyntax` | formula | Write formulas with the names and spill syntax shown in Excel, such as STDEV.S(A1:A10), FILTER(A1:A10,A1:A10>0), and SUM(E1#). OfficeKit Codec maps modeled future functions plus A1# to their required _xlfn/_xlws/ANCHORARRAY XLSX storage forms, returns public formulas without those package prefixes, and preserves an unchanged imported cell formula's original storage spelling. |
| `workbookWindow.getActiveWorksheet` | api | Return the visible active worksheet for one workbook window. |
| `workbookWindow.getSelectedWorksheets` | api | Return one window's visible selected worksheet tabs in workbook order. |
| `workbookWindow.setActiveWorksheet` | api | Set one window's active worksheet and collapse that window's selected tab group to it. |
| `workbookWindow.setSelectedWorksheets` | api | Set one window's non-empty visible selected tab group, which must include its active worksheet. |
| `worksheet.freezePanes.freezeColumns` | api | Freeze a leading column count in the worksheet view while preserving any frozen rows. |
| `worksheet.freezePanes.freezeRows` | api | Freeze a leading row count in the worksheet view while preserving any frozen columns. |
| `worksheet.freezePanes.unfreeze` | api | Remove all frozen worksheet panes and restore a single scrollable view. |
| `worksheet.getRange` | api | Select an A1 range for values, formulas, formatting, merge, fill, and copy operations. |
| `worksheet.getUsedRange` | api | Return the worksheet used rectangle, optionally excluding formatting-only cells with valuesOnly=true. |
| `worksheet.mergeCells` | api | Merge an A1 range as one region or merge each row separately with across=true, retaining only upper-left content. |
| `worksheet.protection` | api | Author, inspect, edit, or remove one passwordless worksheet editing restriction with an intuitive allowed-operation list. Cell locked/hidden styles become effective only while protection is active. This is not encryption or access control; password/hash variants remain source-owned and fail closed on replacement. |
| `worksheet.sortState` | api | Get or set bounded worksheet-level row/column sorting; columnSort=true uses unique single-row conditions across the sort range. |
| `worksheet.unmergeCells` | api | Remove every merged region intersecting an A1 range without discarding the retained upper-left content. |
| `worksheet.visibility` | api | Read or assign native worksheet visibility as visible, hidden, or veryHidden; at least one sheet must remain visible. |
| `worksheetChart.accessibilityCapability` | api | Report sourceBound/editable/addable preflight for a worksheet chart graphic-frame xdr:cNvPr title/description/decorative leaf independently of ChartSpace editability. |
| `worksheetChart.setAccessibilityMetadata` | api | Transactionally add, change, or clear a worksheet chart's non-visible title/description/decorative metadata without changing its visible chart title. Ambiguous imported extension graphs fail closed. |
| `worksheetImage.accessibilityCapability` | api | Report sourceBound/editable/addable preflight for worksheet picture xdr:cNvPr title/description/decorative metadata. |
| `worksheetImage.setAccessibilityMetadata` | api | Transactionally add, change, or clear worksheet picture title/description/decorative metadata. image.alt is the same description state and is never inferred from the object or file name. |

### workbook details

#### `exportXlsxWithOfficeKit`

Export the bounded Workbook model through the bundled C# Open XML SDK NativeAOT codec: cells, formulas, styles, merges, dimensions, freezes, ordinary tables, PNG/JPEG pictures, validation, conditional formatting, threaded-comment roots with direct replies, bar/line/pie/area/doughnut charts, marker-only numeric-X/Y scatter charts, bounded numeric-X/Y/positive-Size bubble charts, standard Office 2010 line/column/stacked sparklines, canonical one-variable or two-variable What-If data tables, native PivotTables with exact item or absolute whole-day date filters, and one source-free XLDAPR dynamic-array metadata profile. Imported QueryTables permit only source-bound one-way refresh hardening through table.setQueryRefreshPolicy; an imported connection may only change explicit refreshOnLoad=true to false through workbook.disableConnectionRefreshOnLoad; a uniquely owned imported Pivot cache may make that same one-way change through pivot.disableRefreshOnLoad; imported dynamic-array topology, commands, fields, sorts, other Pivot configuration/data/output, and unsupported extension graphs are preservation-only or fail closed.

**Schema parameters:**

- `workbook` (Workbook) required — Workbook facade within the core cell/formula/style/merge/dimension/freeze/ordinary-table/image/validation/conditional-format/root-plus-direct-reply-comment/bar-line-pie-chart/standard-sparkline/bounded-source-free-XLDAPR boundary. A recognized imported QueryTable may only receive one-way automatic-refresh hardening through table.setQueryRefreshPolicy, a recognized imported connection may only turn explicit refreshOnLoad=true off through workbook.disableConnectionRefreshOnLoad, and a recognized uniquely owned Pivot cache may only turn explicit refreshOnLoad=true off through pivot.disableRefreshOnLoad; imported dynamic-array topology, commands, fields, sorts, topology, nested reply graphs, mentions, other Pivot configuration/data/output, non-reversible sparkline graphs, and other advanced package graphs must remain unchanged or fail closed.
- `recalculate` (boolean) — Recalculate formulas before serialization; defaults to true.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — XLSX bytes produced by the bundled Open XML SDK NativeAOT codec, with codec diagnostics in metadata.

#### `fx.ABS`

Return the absolute value of a number.

**Examples:**

- =ABS(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ABS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ACOS`

Return the inverse cosine for a finite input in the closed interval [-1,1]; other inputs fail as #NUM!.

**Examples:**

- =ACOS(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ACOS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ACOSH`

Return the inverse hyperbolic cosine for finite inputs at least one; smaller inputs fail as #NUM!.

**Examples:**

- =ACOSH(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ACOSH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ADDRESS`

Return one bounded worksheet address as text from 1-based row and column numbers, reference mode 1 through 4, A1 or R1C1 style, and optional Excel-quoted sheet text. Coordinates outside XFD1048576, invalid modes, nonlogical style selectors, and nontext sheet names fail as #VALUE!.

**Examples:**

- =ADDRESS(2,3)
- =ADDRESS(2,3,2,FALSE)
- =ADDRESS(2,3,1,TRUE,"EXCEL SHEET")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ADDRESS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.AND`

Return TRUE when all conditions are true.

**Examples:**

- =AND(A1>0,B1>0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =AND(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.ASIN`

Return the inverse sine for a finite input in the closed interval [-1,1]; other inputs fail as #NUM!.

**Examples:**

- =ASIN(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ASIN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ASINH`

Return the inverse hyperbolic sine of a finite number.

**Examples:**

- =ASINH(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ASINH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ATAN`

Return the inverse tangent of a finite number.

**Examples:**

- =ATAN(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ATAN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ATAN2`

Return the quadrant-aware angle for x and y coordinates; the origin returns #DIV/0!.

**Examples:**

- =ATAN2(A1,B1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ATAN2(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ATANH`

Return the inverse hyperbolic tangent for finite inputs strictly between -1 and 1; boundary values fail as #NUM!.

**Examples:**

- =ATANH(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ATANH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.AVERAGE`

Average numeric values across arguments and ranges in the clean-room formula engine.

**Examples:**

- =AVERAGE(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =AVERAGE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.AVERAGEIF`

Average values whose corresponding entries match case-insensitive comparison or wildcard criteria.

**Examples:**

- =AVERAGEIF(A1:A10,"East*",B1:B10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =AVERAGEIF(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.AVERAGEIFS`

Average values where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria.

**Examples:**

- =AVERAGEIFS(C1:C10,A1:A10,"East*",B1:B10,">=10")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =AVERAGEIFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.CEILING`

Round a number up to the nearest significance.

**Examples:**

- =CEILING(A1,5)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CEILING(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.CHOOSE`

Select one scalar result from up to 254 ordered choices using a truncated 1-based index; invalid indexes and unsupported arity return #VALUE!.

**Examples:**

- =CHOOSE(A1,"Low","Medium","High")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CHOOSE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.CHOOSECOLS`

Select and reorder one or more 1-based or negative column indexes from an array.

**Examples:**

- =CHOOSECOLS(A2:C10,3,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CHOOSECOLS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.CHOOSEROWS`

Select and reorder one or more 1-based or negative row indexes from an array.

**Examples:**

- =CHOOSEROWS(A2:C10,3,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CHOOSEROWS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.CLEAN`

Remove ASCII C0 control characters from one bounded scalar text value while preserving Unicode text and other controls; overlong, error, or multi-cell inputs fail closed.

**Examples:**

- =CLEAN(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CLEAN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.COLUMN`

Return the 1-based column of the current formula cell or one explicit single-cell reference; ranges, spills, computed matrices, and invalid arity fail closed as #VALUE!.

**Examples:**

- =COLUMN()
- =COLUMN(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COLUMN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COLUMNS`

Return the column count of one bounded rectangular reference or dynamic spill.

**Examples:**

- =COLUMNS(A1:C10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COLUMNS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COMBIN`

Return the number of combinations for two non-negative bounded integer arguments.

**Examples:**

- =COMBIN(10,3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COMBIN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COMBINA`

Return combinations with repetition for two non-negative bounded integer arguments.

**Examples:**

- =COMBINA(5,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COMBINA(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.CONCAT`

Concatenate text values and ranges.

**Examples:**

- =CONCAT(A1,"-",B1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CONCAT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.CORREL`

Return the Pearson correlation coefficient for two same-length bounded sources, ignoring positions where either value is nonnumeric; mismatched lengths return #N/A and empty or zero-variance pairs return #DIV/0!.

**Examples:**

- =CORREL(A1:A10,B1:B10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CORREL(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COS`

Return the cosine of a finite radian value.

**Examples:**

- =COS(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COSH`

Return the hyperbolic cosine of a finite number; overflow fails as #NUM!.

**Examples:**

- =COSH(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COSH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COUNT`

Count numeric values across arguments and ranges.

**Examples:**

- =COUNT(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COUNT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COUNTA`

Count non-empty values across arguments and ranges, including text, logical values, errors, and empty-text formula results.

**Examples:**

- =COUNTA(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COUNTA(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COUNTBLANK`

Count blank cells and formula results that are empty text in one range.

**Examples:**

- =COUNTBLANK(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COUNTBLANK(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COUNTIF`

Count values using case-insensitive numeric/text criteria and Excel ?, *, and ~ wildcard semantics.

**Examples:**

- =COUNTIF(A1:A10,"East*")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COUNTIF(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COUNTIFS`

Count rows where multiple criteria ranges of the same size match case-insensitive comparison or wildcard criteria.

**Examples:**

- =COUNTIFS(A1:A10,"East*",B1:B10,">=10")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COUNTIFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COVARIANCE.P`

Calculate population covariance for two same-length bounded sources with pairwise numeric filtering; mismatched lengths return #N/A and no numeric pairs returns #DIV/0!.

**Examples:**

- =COVARIANCE.P(A1:A10,B1:B10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COVARIANCE.P(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COVARIANCE.S`

Estimate sample covariance for two same-length bounded sources with pairwise numeric filtering; mismatched lengths return #N/A and fewer than two numeric pairs returns #DIV/0!.

**Examples:**

- =COVARIANCE.S(A1:A10,B1:B10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COVARIANCE.S(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.CUMIPMT`

Calculate cumulative interest paid across a bounded inclusive range of constant-payment loan periods.

**Examples:**

- =CUMIPMT(B1,B2,B3,1,12,0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CUMIPMT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- All six arguments are required. The bounded evaluator requires positive rate and present value, payment type 0 or 1, and integer start/end periods ordered from 1 through the term; the ending period is capped at 9,999. Invalid inputs return #VALUE! or #NUM! rather than coercing a range.

#### `fx.CUMPRINC`

Calculate cumulative principal paid across a bounded inclusive range of constant-payment loan periods.

**Examples:**

- =CUMPRINC(B1,B2,B3,1,12,0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CUMPRINC(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- All six arguments are required. The bounded evaluator shares CUMIPMT's positive-rate, positive-present-value, integer-period, bounded-end, and payment-timing contract; it returns the signed principal cash flow.

#### `fx.DATE`

Return an Excel serial in the workbook's 1900 or 1904 date system, with overflow and 1900 serial-60 compatibility.

**Examples:**

- =DATE(2026,7,12)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DATE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.DATEVALUE`

Convert deterministic ISO or English month-name date text to a serial in the workbook's 1900 or 1904 date system; ambiguous locale-numeric dates return #VALUE!.

**Examples:**

- =DATEVALUE("2026-07-13")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DATEVALUE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.DAY`

Return the day component of a serial in the workbook's date system, including 1900 compatibility serial 60.

**Examples:**

- =DAY(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DAY(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.DAYS`

Return the whole-day difference between two Excel date serials.

**Examples:**

- =DAYS(B1,A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DAYS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.DAYS360`

Return the accounting day count between two valid Excel date serials using the U.S. NASD 30/360 method by default or the European 30E/360 method when the optional logical method is TRUE. Invalid dates, method text, and arity fail explicitly.

**Examples:**

- =DAYS360(A1,B1)
- =DAYS360(A1,B1,TRUE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DAYS360(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.DB`

Calculate one fixed-declining-balance depreciation period with an optional first-year month count.

**Examples:**

- =DB(B1,B2,B3,A2)
- =DB(B1,B2,B3,A2,6)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DB(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires nonnegative cost and salvage, salvage no greater than cost, integer life and period from 1 through 9,999, and an integer month from 1 through 12. A partial first year permits one prorated final period; the native three-decimal declining rate is not silently switched to straight-line.

#### `fx.DDB`

Calculate one double-declining-balance depreciation period with an optional positive factor.

**Examples:**

- =DDB(B1,B2,B3,A2)
- =DDB(B1,B2,B3,A2,1.5)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DDB(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires nonnegative cost and salvage, salvage no greater than cost, and integer life and period from 1 through 9,999. The factor defaults to 2, must be positive, and depreciation is capped at the remaining amount above salvage without a silent straight-line switch.

#### `fx.DEGREES`

Convert finite radians to degrees with an explicit non-finite-result guard.

**Examples:**

- =DEGREES(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DEGREES(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.DROP`

Drop rows and optional columns from the start or end of an array and spill the remainder.

**Examples:**

- =DROP(A2:C10,1,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DROP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.EDATE`

Shift a serial date by whole months and clamp the day to the target month end.

**Examples:**

- =EDATE(A1,3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =EDATE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.EOMONTH`

Return the final date serial of a month offset from a start date.

**Examples:**

- =EOMONTH(A1,0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =EOMONTH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.EVEN`

Round a finite number away from zero to the next even integer.

**Examples:**

- =EVEN(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =EVEN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.EXACT`

Compare two bounded scalar text values with case-sensitive equality; multi-cell sources and overlong values fail closed.

**Examples:**

- =EXACT(A1,"Approved")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =EXACT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.EXP`

Return e raised to a finite number; overflow fails as #NUM! instead of leaking Infinity.

**Examples:**

- =EXP(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =EXP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.EXPAND`

Expand an array to requested row and column dimensions with optional padding.

**Examples:**

- =EXPAND(A2:B3,4,3,"n/a")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =EXPAND(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.FACT`

Return the factorial of a non-negative integer through the finite 170! boundary.

**Examples:**

- =FACT(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FACT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.FACTDOUBLE`

Return the double factorial of a non-negative integer through the bounded finite range.

**Examples:**

- =FACTDOUBLE(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FACTDOUBLE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.FALSE`

Return the logical value FALSE with no arguments; supplied arguments fail as #VALUE!.

**Examples:**

- =FALSE()

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FALSE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.FILTER`

Filter rows from a source range with a boolean or comparison include array and spill the matching rows.

**Examples:**

- =FILTER(A2:C10,B2:B10="East")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FILTER(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.FIND`

Return the 1-based position of a case-sensitive literal text sequence.

**Examples:**

- =FIND("Review",A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FIND(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.FLOOR`

Round a number down to the nearest significance.

**Examples:**

- =FLOOR(A1,5)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FLOOR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.FORECAST.LINEAR`

Predict one y value from one bounded scalar x and aligned known-y/known-x sources using the shared stable linear fit; nonnumeric x returns #VALUE!, source mismatch or no pairs returns #N/A, and zero x variance returns #DIV/0!.

**Examples:**

- =FORECAST.LINEAR(D2,B2:B10,A2:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FORECAST.LINEAR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.FORMULATEXT`

Return the stored formula text for one explicit single-cell reference, #N/A when that cell has no formula, and #VALUE! for ranges, computed matrices, spills, or invalid input.

**Examples:**

- =FORMULATEXT(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FORMULATEXT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.FV`

Calculate the future value of a finite constant-payment stream from rate, term, payment, optional present value, and payment timing.

**Examples:**

- =FV(B1,B2,B3)
- =FV(B1,B2,B3,B4,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FV(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires rate > -1, a positive finite term, and payment type 0 or 1. It uses the same cash-flow equation as PMT and PV, including the zero-rate case.

#### `fx.GCD`

Return the greatest common divisor of bounded integer arguments and ranges; unsafe integer results fail closed as #NUM!.

**Examples:**

- =GCD(A1:A3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =GCD(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.GROWTH`

Return a bounded single-variable exponential prediction dynamic array for y=b*m^x with the same row or column shape as new-x. Known-y must be positive; x arguments may be omitted, const may force b=1, and constant known-x is removed. Overflow, multivariable or two-dimensional inputs, nonnumeric new-x positions, and mismatched known source shapes fail closed.

**Examples:**

- =GROWTH(B2:B10,A2:A10,D2:D4,TRUE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =GROWTH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.HLOOKUP`

Look up one scalar in the first row of a nonempty rectangular range of at most 10,000 cells; FALSE/0 performs an exact, wildcard-aware lookup, while TRUE/1 or omission requires a proven ascending homogeneous numeric or text key row and returns the greatest matching-or-lower key. Invalid table/mode/index inputs and unproven ordering return #VALUE!, while an out-of-range return-row index returns #REF!.

**Examples:**

- =HLOOKUP("Revenue",A1:D4,3,FALSE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =HLOOKUP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.HOUR`

Return the 0 through 23 hour component from a nonnegative serial or supported time text.

**Examples:**

- =HOUR(TIMEVALUE("6:45 PM"))

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =HOUR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.HSTACK`

Append arrays horizontally, padding shorter arrays with #N/A to the maximum row count.

**Examples:**

- =HSTACK(A2:B4,D2:E3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =HSTACK(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.IF`

Return one value when a condition is true and another when false.

**Examples:**

- =IF(A1>0,"ok","bad")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =IF(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.IFERROR`

Return a fallback value when an expression evaluates to a formula error.

**Examples:**

- =IFERROR(XLOOKUP("missing",A1:A10,B1:B10),"not found")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =IFERROR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.IFNA`

Return a fallback only when an expression evaluates to #N/A; preserve every other result or error.

**Examples:**

- =IFNA(XLOOKUP("missing",A1:A10,B1:B10),"not found")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =IFNA(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.IFS`

Evaluate condition/value pairs in order and return the first matching value, or #N/A when no condition matches.

**Examples:**

- =IFS(A1>=90,"A",A1>=80,"B",TRUE,"C")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =IFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.INDEX`

Select one value from a nonempty rectangular range of at most 10,000 cells with host-compatible row and optional column selectors, preserving an error-valued selector such as a failed MATCH. Only the documented 2- or 3-argument array/range form is modeled; missing or extra selectors and oversized ranges return #VALUE!, while a missing or out-of-range source cell returns #REF!.

**Examples:**

- =INDEX(A2:C4,2,3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =INDEX(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.INT`

Round a number down to the nearest integer.

**Examples:**

- =INT(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =INT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.INTERCEPT`

Return the y-axis intercept for the same bounded source-aware linear regression profile as SLOPE; empty or mismatched sources return #N/A and zero x variance returns #DIV/0!.

**Examples:**

- =INTERCEPT(B2:B10,A2:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =INTERCEPT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.IPMT`

Calculate the interest component of one constant-payment loan period from finite rate, period, term, present value, optional future value, and payment-timing inputs.

**Examples:**

- =IPMT(B1,A2,B2,B3)
- =IPMT(B1,A2,B2,B3,B4,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =IPMT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires rate > -1, a positive term, an integer period from 1 through the term, and payment type 0 or 1. Period-one interest is zero for payment type 1; invalid inputs return #VALUE! or #NUM!.

#### `fx.IRR`

Return a bounded-convergence periodic return rate for a finite cash-flow vector.

**Examples:**

- =IRR(B2:B8)
- =IRR(B2:B8,0.15)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =IRR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- Cash flows must contain both a positive and a negative finite number. The optional finite guess defaults to 0.1; no converged valid root or an invalid rate returns #NUM! rather than a guessed value.

#### `fx.ISBLANK`

Return TRUE when a referenced value is empty.

**Examples:**

- =ISBLANK(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISBLANK(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISERR`

Return TRUE for recognized formula errors other than #N/A.

**Examples:**

- =ISERR(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISERR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISERROR`

Return TRUE when a value is any recognized formula error.

**Examples:**

- =ISERROR(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISERROR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISFORMULA`

Return TRUE when one explicit single-cell reference contains a formula, FALSE when the cell is not formula-backed, and #VALUE! for ranges, computed matrices, spills, or invalid input.

**Examples:**

- =ISFORMULA(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISFORMULA(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISLOGICAL`

Return TRUE when a value is a logical TRUE or FALSE.

**Examples:**

- =ISLOGICAL(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISLOGICAL(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISNA`

Return TRUE only when a value is the #N/A error.

**Examples:**

- =ISNA(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISNA(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISNONTEXT`

Return TRUE when a value is not text, including blank, logical, numeric, and error values.

**Examples:**

- =ISNONTEXT(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISNONTEXT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISNUMBER`

Return TRUE when a value is numeric.

**Examples:**

- =ISNUMBER(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISNUMBER(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISOWEEKNUM`

Return the ISO 8601 week number for one valid Excel date serial in the workbook's 1900 or 1904 date system. OfficeKit Codec owns the required _xlfn.ISOWEEKNUM package spelling.

**Examples:**

- =ISOWEEKNUM(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISOWEEKNUM(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISREF`

Return TRUE only for a direct A1, defined-name, or spill reference expression; computed values and functions return FALSE, while invalid arity fails closed as #VALUE!.

**Examples:**

- =ISREF(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISREF(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISTEXT`

Return TRUE when a value is text and not a formula error.

**Examples:**

- =ISTEXT(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISTEXT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.LARGE`

Return the k-th largest numeric value in an array or range.

**Examples:**

- =LARGE(A1:A10,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LARGE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.LCM`

Return the least common multiple of bounded integer arguments and ranges; zero inputs return zero and unsafe overflow returns #NUM!.

**Examples:**

- =LCM(A1:A3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LCM(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.LEFT`

Return up to 32,767 Unicode characters from the start of one bounded scalar text value; num_chars defaults to 1 and invalid or multi-cell inputs fail closed.

**Examples:**

- =LEFT(A1,3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LEFT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.LEN`

Return the Unicode code-point length of one bounded scalar text value; overlong, error, or multi-cell inputs fail closed.

**Examples:**

- =LEN(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LEN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.LET`

Bind up to 16 scalar local names from left to right and evaluate a final scalar expression; write the public formula exactly as Excel displays it, while OfficeKit Codec owns scoped _xlfn.LET/_xlpm package spelling. Invalid names, array-valued bindings, and missing arguments fail closed as #VALUE!.

**Examples:**

- =LET(rate,0.1,principal,1000,principal*(1+rate))

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LET(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.LINEST`

Return a bounded single-variable least-squares dynamic array: 1x2 slope/intercept by default or the documented 5x2 coefficient, error, R-squared, F/df, and regression/residual statistics matrix when stats is TRUE. Known-x may be omitted, const may force a zero intercept, constant known-x is removed, and mismatched shapes return #N/A; multivariable inputs and array constants remain unsupported.

**Examples:**

- =LINEST(B2:B10,A2:A10,TRUE,TRUE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LINEST(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.LN`

Return the natural logarithm of a positive finite number; non-positive inputs fail as #NUM!.

**Examples:**

- =LN(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.LOG`

Return a logarithm for a positive number and positive base other than one; the base defaults to 10 and invalid domains fail as #NUM!.

**Examples:**

- =LOG(A1)
- =LOG(A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LOG(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.LOG10`

Return the base-10 logarithm of a positive finite number.

**Examples:**

- =LOG10(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LOG10(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.LOGEST`

Return a bounded single-variable exponential regression dynamic array for y=b*m^x: 1x2 multiplier/base by default or a 5x2 matrix whose remaining diagnostics describe the natural-log regression when stats is TRUE. Known-y must be positive; known-x may be omitted, const may force b=1, and constant known-x is removed. Mismatched shapes, multivariable inputs, and array constants fail closed.

**Examples:**

- =LOGEST(B2:B10,A2:A10,TRUE,TRUE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LOGEST(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.LOOKUP`

Return the result aligned with the greatest ascending homogeneous numeric or text key less than or equal to one scalar. The bounded vector form accepts one optional same-length result vector; the array form searches its first column when square or taller and its first row when wider, then returns from the last column or row. Unproven ordering, mixed keys, mismatched vectors, two-dimensional vector arguments, and sources above 10,000 cells fail as #VALUE!.

**Examples:**

- =LOOKUP(5.75,A2:A6,B2:B6)
- =LOOKUP(25,A1:D2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LOOKUP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.LOWER`

Convert text to lowercase.

**Examples:**

- =LOWER(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LOWER(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.MATCH`

Return a 1-based lookup position in one row or column vector of 1 through 10,000 cells. Exact 0 matching is wildcard-aware; default/1 approximate matching requires a proven ascending homogeneous numeric or text vector and returns the greatest matching-or-lower key, while -1 requires proven descending order and returns the smallest matching-or-higher key. Two-dimensional, oversized, mixed, unordered, or invalid-mode inputs return #VALUE!.

**Examples:**

- =MATCH("Beta*",A2:A10,0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MATCH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MAX`

Return the maximum numeric value across arguments and ranges.

**Examples:**

- =MAX(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MAX(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MAXIFS`

Return the largest numeric value where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria.

**Examples:**

- =MAXIFS(C1:C10,A1:A10,"East*",B1:B10,">=10")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MAXIFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MEDIAN`

Return the middle numeric value, or the average of the two middle values, across arguments and ranges.

**Examples:**

- =MEDIAN(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MEDIAN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MID`

Return a bounded Unicode slice from one scalar text value using a 1-based start and non-negative character count; invalid or multi-cell inputs fail closed.

**Examples:**

- =MID(A1,2,3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MID(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.MIN`

Return the minimum numeric value across arguments and ranges.

**Examples:**

- =MIN(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MIN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MINIFS`

Return the smallest numeric value where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria.

**Examples:**

- =MINIFS(C1:C10,A1:A10,"East*",B1:B10,">=10")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MINIFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MINUTE`

Return the 0 through 59 minute component from a nonnegative serial or supported time text.

**Examples:**

- =MINUTE(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MINUTE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MIRR`

Calculate a modified periodic internal rate of return using distinct finance and reinvestment rates for a finite cash-flow vector.

**Examples:**

- =MIRR(B2:B6,B7,B8)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MIRR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator accepts 2 through 10,000 finite cash flows containing both signs. Finance and reinvestment rates must be greater than -1; negative flows are discounted at the finance rate, positive flows compound at the reinvestment rate, and invalid profiles return #VALUE! or #NUM! rather than choosing an implied rate.

#### `fx.MOD`

Return the remainder after division, preserving the divisor sign and returning #DIV/0! for a zero divisor.

**Examples:**

- =MOD(A1,7)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MOD(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MODE.MULT`

Return every numeric value tied for the highest frequency as an ascending vertical spill; if no value repeats, return #N/A instead of synthesizing modes.

**Examples:**

- =MODE.MULT(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MODE.MULT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.MODE.SNGL`

Return the most frequently occurring numeric value, or #N/A when no value repeats.

**Examples:**

- =MODE.SNGL(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MODE.SNGL(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MONTH`

Return the month component of a serial in the workbook's 1900 or 1904 date system.

**Examples:**

- =MONTH(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MONTH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MROUND`

Round a finite number to the nearest multiple with explicit zero-multiple and sign checks.

**Examples:**

- =MROUND(A1,5)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MROUND(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.N`

Return a bounded numeric coercion: numbers and date serials unchanged, TRUE/FALSE as 1/0, text or blank as 0, and formula errors propagated; multi-cell or matrix input fails closed as #VALUE!.

**Examples:**

- =N(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =N(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.NA`

Return the #N/A error value to mark unavailable data explicitly.

**Examples:**

- =NA()

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =NA(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.NETWORKDAYS`

Count Monday-through-Friday dates inclusively between two serial dates, excluding optional holidays.

**Examples:**

- =NETWORKDAYS(A1,B1,Holidays)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =NETWORKDAYS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.NETWORKDAYS.INTL`

Count inclusive workdays with a numbered or Monday-first seven-character custom weekend and optional holidays.

**Examples:**

- =NETWORKDAYS.INTL(A1,B1,7,Holidays)
- =NETWORKDAYS.INTL(A1,B1,"0000011")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =NETWORKDAYS.INTL(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.NOT`

Reverse the truth value of a condition.

**Examples:**

- =NOT(A1>0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =NOT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.NPER`

Solve the finite payment-period count from rate, payment, present value, optional future value, and payment timing.

**Examples:**

- =NPER(B1,B2,B3)
- =NPER(B1,B2,B3,B4,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =NPER(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires rate > -1 and payment type 0 or 1. It returns a closed-form finite period count, which may be zero or negative for the supplied cash-flow signs; a zero payment at zero rate or an invalid real solution returns #NUM!.

#### `fx.NPV`

Discount a finite periodic cash-flow vector beginning one period after the present value date.

**Examples:**

- =NPV(B1,B2:B6)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =NPV(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- Rate must be finite and greater than -1. The bounded evaluator accepts at most 10,000 finite numeric cash flows and returns #VALUE! or #NUM! rather than coercing malformed inputs.

#### `fx.ODD`

Round a finite number away from zero to the next odd integer.

**Examples:**

- =ODD(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ODD(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.OR`

Return TRUE when any condition is true.

**Examples:**

- =OR(A1>0,B1>0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =OR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.PERCENTILE.EXC`

Return an exclusive percentile from a bounded numeric range using rank k*(n+1); k must be strictly between 0 and 1, and endpoints that cannot be interpolated return #NUM!.

**Examples:**

- =PERCENTILE.EXC(A1:A10,0.9)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =PERCENTILE.EXC(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.PERCENTILE.INC`

Return an inclusive percentile from a bounded array or range; k must be from 0 through 1 and the result uses linear interpolation, while nonnumeric reference values are ignored, formula errors propagate, and an empty numeric set fails as #NUM!.

**Examples:**

- =PERCENTILE.INC(A1:A10,0.9)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =PERCENTILE.INC(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.PI`

Return the deterministic mathematical constant π; arguments are rejected rather than ignored.

**Examples:**

- =PI()

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =PI(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.PMT`

Calculate a constant-period loan payment from finite rate, term, present value, optional future value, and payment-timing inputs.

**Examples:**

- =PMT(B1,B2,B3)
- =PMT(B1,B2,B3,B4,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =PMT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires rate > -1, a positive term, and payment type 0 or 1. Invalid numeric inputs return #VALUE! or #NUM!.

#### `fx.POWER`

Raise a finite base to a finite exponent; non-finite results fail as #NUM! rather than leaking JavaScript Infinity or NaN.

**Examples:**

- =POWER(A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =POWER(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.PPMT`

Calculate the principal component of one constant-payment loan period using the same bounded inputs as IPMT.

**Examples:**

- =PPMT(B1,A2,B2,B3)
- =PPMT(B1,A2,B2,B3,B4,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =PPMT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- For every supported period, PMT equals IPMT plus PPMT. The evaluator rejects an out-of-range or non-integer period and invalid payment timing with #NUM! rather than coercing them.

#### `fx.PRODUCT`

Multiply numeric values across arguments and bounded ranges; formula errors propagate and empty invocation returns #VALUE!.

**Examples:**

- =PRODUCT(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =PRODUCT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.PV`

Calculate the present value of a finite constant-payment stream from rate, term, payment, optional future value, and payment timing.

**Examples:**

- =PV(B1,B2,B3)
- =PV(B1,B2,B3,B4,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =PV(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires rate > -1, a positive finite term, and payment type 0 or 1. It preserves standard cash-flow signs and returns #VALUE! or #NUM! for invalid inputs rather than coercing them.

#### `fx.QUARTILE.EXC`

Return an exclusive first, second, or third quartile from a bounded numeric range; the selector is truncated and indexes outside 1 through 3 return #NUM!.

**Examples:**

- =QUARTILE.EXC(A1:A10,3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =QUARTILE.EXC(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.QUARTILE.INC`

Return an inclusive quartile from a bounded array or range; the quartile index must be an integer from 0 through 4 and the result uses linear interpolation, while nonnumeric reference values are ignored, formula errors propagate, and an empty numeric set fails as #NUM!.

**Examples:**

- =QUARTILE.INC(A1:A10,3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =QUARTILE.INC(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.QUOTIENT`

Return the integer portion of a division result, truncating toward zero and returning #DIV/0! for a zero divisor.

**Examples:**

- =QUOTIENT(A1,7)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =QUOTIENT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.RADIANS`

Convert finite degrees to radians with an explicit non-finite-result guard.

**Examples:**

- =RADIANS(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =RADIANS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.RANK.AVG`

Return a number's rank in a bounded numeric range and average the occupied positions when values tie; a number absent from the range returns #N/A.

**Examples:**

- =RANK.AVG(A1,A1:A10,0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =RANK.AVG(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.RANK.EQ`

Return a number's equal rank in a numeric range, descending by default or ascending when order is nonzero.

**Examples:**

- =RANK.EQ(A1,A1:A10,0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =RANK.EQ(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.RATE`

Solve a bounded periodic interest rate from an integer payment term, payment, present value, optional future value, payment timing, and optional guess.

**Examples:**

- =RATE(B1,B2,B3)
- =RATE(B1,B2,B3,B4,1,0.1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =RATE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The evaluator accepts an integer term from 1 through 9,999, type 0 or 1, and a finite guess greater than -1 (default 0.1). It finds a converged rate greater than -1 nearest the guess or returns #NUM! rather than inventing a rate.

#### `fx.REPLACE`

Replace a bounded scalar text span using 1-based character and non-negative length arguments; invalid positions, matrices, and overlong results fail closed.

**Examples:**

- =REPLACE(A1,1,5,"Draft")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =REPLACE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.REPT`

Repeat one bounded scalar text value an integer number of times, with a 32,767-character result budget.

**Examples:**

- =REPT("-",10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =REPT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.RIGHT`

Return up to 32,767 Unicode characters from the end of one bounded scalar text value; num_chars defaults to 1 and invalid or multi-cell inputs fail closed.

**Examples:**

- =RIGHT(A1,3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =RIGHT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.ROUND`

Round a numeric value to decimal places or, with negative digits, positions left of the decimal point.

**Examples:**

- =ROUND(A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ROUND(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ROUNDDOWN`

Round a numeric value toward zero at the requested positive or negative digit position.

**Examples:**

- =ROUNDDOWN(A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ROUNDDOWN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ROUNDUP`

Round a numeric value away from zero at the requested positive or negative digit position.

**Examples:**

- =ROUNDUP(A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ROUNDUP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ROW`

Return the 1-based row of the current formula cell or one explicit single-cell reference; ranges, spills, computed matrices, and invalid arity fail closed as #VALUE!.

**Examples:**

- =ROW()
- =ROW(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ROW(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ROWS`

Return the row count of one bounded rectangular reference or dynamic spill.

**Examples:**

- =ROWS(A1:C10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ROWS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.RSQ`

Return the square of Pearson correlation for aligned known-y and known-x sources; positions are pairwise filtered, length mismatch or no pairs returns #N/A, and fewer than two or zero-variance pairs returns #DIV/0!.

**Examples:**

- =RSQ(B2:B10,A2:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =RSQ(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SEARCH`

Return the 1-based position of case-insensitive text, supporting Excel ?, *, and ~ wildcard syntax.

**Examples:**

- =SEARCH("review",A1)
- =SEARCH("Re*W",A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SEARCH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SECOND`

Return the 0 through 59 second component from a nonnegative serial or supported time text.

**Examples:**

- =SECOND(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SECOND(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SEQUENCE`

Return a dynamic array sequence that spills into neighboring cells in the clean-room formula engine.

**Examples:**

- =SEQUENCE(2,3,10,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SEQUENCE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.SHEET`

Return the 1-based OfficeKit worksheet position for the current sheet or one validated single-sheet cell/range, workbook defined name, table, or sheet-name string. Missing sheet-name strings return #N/A; invalid references, nonreference values, 3D spans, and extra arguments fail explicitly. Chart, macro, and dialog sheets are outside the OfficeKit workbook model.

**Examples:**

- =SHEET()
- =SHEET('Source Data'!A1)
- =SHEET("Summary")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SHEET(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SHEETS`

Return the total number of OfficeKit worksheets, including hidden worksheets, or 1 for one validated single-sheet cell/range, workbook defined name, table, or sheet-name string. Invalid references, nonreference values, 3D spans, and extra arguments fail explicitly; chart, macro, and dialog sheets are not modeled.

**Examples:**

- =SHEETS()
- =SHEETS('Source Data'!A1:C10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SHEETS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SIGN`

Return -1, 0, or 1 according to the sign of a finite numeric value.

**Examples:**

- =SIGN(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SIGN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SIN`

Return the sine of a finite radian value.

**Examples:**

- =SIN(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SIN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SINH`

Return the hyperbolic sine of a finite number; overflow fails as #NUM!.

**Examples:**

- =SINH(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SINH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SLN`

Calculate straight-line depreciation from cost, salvage value, and useful life.

**Examples:**

- =SLN(B1,B2,B3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SLN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The evaluator accepts finite numeric inputs and returns #DIV/0! for zero life rather than coercing a rate. This is the direct per-period expense, not a declining-balance schedule.

#### `fx.SLOPE`

Return the least-squares slope for aligned known-y and known-x sources using stable pair moments; nonnumeric reference positions are ignored together, mismatched lengths return #N/A, and zero x variance returns #DIV/0!.

**Examples:**

- =SLOPE(B2:B10,A2:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SLOPE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SMALL`

Return the k-th smallest numeric value in an array or range.

**Examples:**

- =SMALL(A1:A10,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SMALL(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SORT`

Sort a range by a 1-based column index and spill the sorted rows.

**Examples:**

- =SORT(A2:C10,3,-1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SORT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.SQRT`

Return the non-negative square root of a finite number; negative inputs return #NUM!.

**Examples:**

- =SQRT(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SQRT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.STDEV.P`

Calculate population standard deviation with a numerically stable bounded calculation; references ignore text, logical, blank, and error cells, while direct logical and numeric-text arguments are counted, direct errors propagate, and an empty numeric set returns #DIV/0!.

**Examples:**

- =STDEV.P(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =STDEV.P(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.STDEV.S`

Estimate sample standard deviation with a numerically stable bounded calculation; references ignore text, logical, blank, and error cells, while direct logical and numeric-text arguments are counted, direct errors propagate, and fewer than two numbers returns #DIV/0!.

**Examples:**

- =STDEV.S(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =STDEV.S(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.STEYX`

Return the standard error of predicted y values for a bounded linear regression; pairwise source semantics match SLOPE, fewer than three numeric pairs returns #DIV/0!, and mismatched source lengths return #N/A.

**Examples:**

- =STEYX(B2:B10,A2:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =STEYX(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SUBSTITUTE`

Replace all or one 1-based occurrence of a literal substring in bounded scalar text; matching is case-sensitive and empty search text fails closed.

**Examples:**

- =SUBSTITUTE(A1,"-","/")
- =SUBSTITUTE(A1,"-","/",2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SUBSTITUTE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.SUM`

Sum numeric values across arguments and ranges.

**Examples:**

- =SUM(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SUM(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SUMIF`

Sum corresponding values using case-insensitive numeric/text criteria and Excel ?, *, and ~ wildcards.

**Examples:**

- =SUMIF(A1:A10,"East*",B1:B10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SUMIF(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SUMIFS`

Sum values where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria.

**Examples:**

- =SUMIFS(C1:C10,A1:A10,"East*",B1:B10,">=10")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SUMIFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SUMPRODUCT`

Multiply corresponding numeric values in equally sized arrays and return the sum of those products; bounded same-shape direct-range predicate factors support comparisons, unary signs, and scalar arithmetic within SUMPRODUCT.

**Examples:**

- =SUMPRODUCT(A1:A10,B1:B10)
- =SUMPRODUCT(C1:C10,--(A1:A10="Open"))

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SUMPRODUCT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SUMSQ`

Sum the squares of numeric values across arguments and bounded ranges; overflow returns #NUM! and formula errors propagate.

**Examples:**

- =SUMSQ(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SUMSQ(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SWITCH`

Match an expression against ordered value/result pairs and return an optional default or #N/A when no value matches.

**Examples:**

- =SWITCH(A1,"East",1,"West",2,0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SWITCH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.SYD`

Calculate sum-of-years-digits depreciation for one bounded useful-life period.

**Examples:**

- =SYD(B1,B2,B3,A2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SYD(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires nonnegative cost and salvage, salvage no greater than cost, and integer life and period from 1 through 9,999. The period cannot exceed life; invalid inputs return #VALUE! or #NUM! rather than being coerced.

#### `fx.T`

Return text unchanged, convert non-text scalars to empty text, and propagate formula errors; multi-cell or matrix input fails closed as #VALUE!.

**Examples:**

- =T(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =T(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.TAKE`

Take rows and optional columns from the start or end of an array and spill the result.

**Examples:**

- =TAKE(A2:C10,3,-2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TAKE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.TAN`

Return the tangent of a finite radian value.

**Examples:**

- =TAN(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TAN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.TANH`

Return the hyperbolic tangent of a finite number.

**Examples:**

- =TANH(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TANH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.TEXT`

Format an Excel serial date as text with the bounded yyyy, yy, m/mm/mmm/mmmm, and d/dd token profile and literal separators.

**Examples:**

- =TEXT(DATE(2026,7,12),"yyyymmdd")
- =TEXT(A1,"mmm yyyy")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TEXT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.TEXTAFTER`

Return the text after a delimiter occurrence, with bounded positive/negative instance selection, case mode, end matching, and an explicit not-found result.

**Examples:**

- =TEXTAFTER(A1,"::")
- =TEXTAFTER(A1,"/",-1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TEXTAFTER(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.TEXTBEFORE`

Return the text before a delimiter occurrence, with bounded positive/negative instance selection, case mode, end matching, and an explicit not-found result.

**Examples:**

- =TEXTBEFORE(A1,"::")
- =TEXTBEFORE(A1,"/",-1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TEXTBEFORE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.TEXTJOIN`

Join text values with a delimiter and optional empty-value skipping.

**Examples:**

- =TEXTJOIN("/",TRUE,A1:A3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TEXTJOIN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.TEXTSPLIT`

Split one scalar text value into a bounded spilled matrix by column and optional row delimiters, with empty-item skipping, case mode, and padding; multi-cell sources, empty delimiters, and oversized results fail closed.

**Examples:**

- =TEXTSPLIT(A1,"|")
- =TEXTSPLIT(A1,"=",";",TRUE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TEXTSPLIT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.TIME`

Return a time fraction from hour, minute, and second values from 0 through 32767, carrying overflow and wrapping at 24 hours.

**Examples:**

- =TIME(16,48,10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TIME(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.TIMEVALUE`

Convert deterministic 12-hour or 24-hour time text, optionally following date text, to a fraction of one day.

**Examples:**

- =TIMEVALUE("6:45 PM")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TIMEVALUE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.TOCOL`

Flatten an array into one spilled column, optionally ignoring blanks or errors and scanning by column.

**Examples:**

- =TOCOL(A2:C10,1,TRUE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TOCOL(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.TOROW`

Flatten an array into one spilled row, optionally ignoring blanks or errors and scanning by column.

**Examples:**

- =TOROW(A2:C10,1,TRUE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TOROW(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.TRANSPOSE`

Transpose a source range into a spilled dynamic array with spillRange/spillValues inspect metadata.

**Examples:**

- =TRANSPOSE(A1:C2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TRANSPOSE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.TREND`

Return a bounded single-variable linear prediction dynamic array with the same row or column shape as new-x. Known-x and new-x may be omitted, const may force a zero intercept, and a constant known-x column is removed; multivariable or two-dimensional inputs, nonnumeric new-x positions, and mismatched known source shapes fail closed.

**Examples:**

- =TREND(B2:B10,A2:A10,D2:D4,TRUE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TREND(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.TRIM`

Trim leading/trailing whitespace and collapse internal whitespace.

**Examples:**

- =TRIM(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TRIM(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.TRIMMEAN`

Average a bounded numeric range after removing an even number of observations symmetrically from both tails; the requested percentage must be from 0 through 1.

**Examples:**

- =TRIMMEAN(A1:A20,0.1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TRIMMEAN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.TRUE`

Return the logical value TRUE with no arguments; supplied arguments fail as #VALUE!.

**Examples:**

- =TRUE()

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TRUE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.TRUNC`

Truncate a finite number toward zero at an optional decimal position without rounding.

**Examples:**

- =TRUNC(A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TRUNC(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.TYPE`

Return Excel type codes 1 for numbers or blank, 2 for text, 4 for logical, 16 for errors, or 64 for arrays and multi-cell references; bounded spill/reference detection is explicit and invalid arity fails closed.

**Examples:**

- =TYPE(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TYPE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.UNICHAR`

Return one Unicode scalar character for an integer from 1 through 1,114,111; surrogate values, invalid ranges, errors, and multi-cell inputs fail closed.

**Examples:**

- =UNICHAR(128512)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =UNICHAR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.UNICODE`

Return the Unicode code point of the first character in one bounded scalar text value; empty, overlong, error, or multi-cell inputs fail closed.

**Examples:**

- =UNICODE(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =UNICODE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.UNIQUE`

Return unique rows from a range as a spilled dynamic array.

**Examples:**

- =UNIQUE(A2:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =UNIQUE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.UPPER`

Convert text to uppercase.

**Examples:**

- =UPPER(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =UPPER(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.VALUE`

Convert deterministic ASCII numeric text with optional grouping, scientific notation, accounting parentheses, or percent suffix to a number.

**Examples:**

- =VALUE("1,234.50")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =VALUE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.VAR.P`

Calculate population variance with a numerically stable bounded calculation; references ignore text, logical, blank, and error cells, while direct logical and numeric-text arguments are counted, direct errors propagate, and an empty numeric set returns #DIV/0!.

**Examples:**

- =VAR.P(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =VAR.P(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.VAR.S`

Estimate sample variance with a numerically stable bounded calculation; references ignore text, logical, blank, and error cells, while direct logical and numeric-text arguments are counted, direct errors propagate, and fewer than two numbers returns #DIV/0!.

**Examples:**

- =VAR.S(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =VAR.S(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.VLOOKUP`

Look up one scalar in the first column of a nonempty rectangular range of at most 10,000 cells; FALSE/0 performs an exact, wildcard-aware lookup, while TRUE/1 or omission requires a proven ascending homogeneous numeric or text key column and returns the greatest matching-or-lower key. Invalid table/mode/index inputs and unproven ordering return #VALUE!, while an out-of-range return-column index returns #REF!.

**Examples:**

- =VLOOKUP("Beta",A2:B4,2,FALSE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =VLOOKUP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.VSTACK`

Append arrays vertically, padding narrower arrays with #N/A to the maximum column count.

**Examples:**

- =VSTACK(A2:B4,A7:A9)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =VSTACK(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.WEEKDAY`

Return a weekday number for Excel return types 1, 2, 3, and 11 through 17.

**Examples:**

- =WEEKDAY(A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =WEEKDAY(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.WEEKNUM`

Return a calendar week number under Excel system 1 for return types 1, 2, and 11 through 17, or the ISO 8601 week number for return type 21; invalid dates and return types fail explicitly.

**Examples:**

- =WEEKNUM(A1,2)
- =WEEKNUM(A1,21)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =WEEKNUM(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.WORKDAY`

Move forward or backward by working days while skipping weekends and optional holidays.

**Examples:**

- =WORKDAY(A1,10,Holidays)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =WORKDAY(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.WORKDAY.INTL`

Move by workdays using a numbered or Monday-first seven-character custom weekend and optional holidays.

**Examples:**

- =WORKDAY.INTL(A1,10,11,Holidays)
- =WORKDAY.INTL(A1,10,"0000011")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =WORKDAY.INTL(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.WRAPCOLS`

Wrap a one-dimensional vector into columns of a requested height, padding the final column when needed.

**Examples:**

- =WRAPCOLS(A2:A10,3,"n/a")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =WRAPCOLS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.WRAPROWS`

Wrap a one-dimensional vector into rows of a requested width, padding the final row when needed.

**Examples:**

- =WRAPROWS(A2:A10,3,"n/a")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =WRAPROWS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result or an Excel-style formula error string.

#### `fx.XIRR`

Return a bounded-convergence annualized return rate for date-aligned finite cash flows using a 365-day year.

**Examples:**

- =XIRR(B2:B8,C2:C8)
- =XIRR(B2:B8,C2:C8,0.15)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =XIRR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- Values and dates must have the same nonzero count, dates must be valid, and cash flows must contain both signs. The optional finite guess defaults to 0.1; invalid or unconverged cases return #VALUE! or #NUM!.

#### `fx.XLOOKUP`

Look up one scalar in same-shaped one-dimensional row or column vectors of 1 through 10,000 cells; exact, next-smaller, next-larger, wildcard, and first/last linear search modes are modeled, while binary-search modes and mismatched or two-dimensional ranges fail as #VALUE!.

**Examples:**

- =XLOOKUP("Gamma",A2:A4,B2:B4,"missing")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =XLOOKUP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.XMATCH`

Return a 1-based lookup position in one row or column vector of 1 through 10,000 cells, with exact, next-smaller, next-larger, wildcard, and forward or reverse linear search modes; two-dimensional, oversized, and binary-search inputs fail as #VALUE!.

**Examples:**

- =XMATCH("Beta*",A2:A10,2,-1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =XMATCH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.XNPV`

Discount date-aligned finite cash flows by actual day offsets from the first date using a 365-day year.

**Examples:**

- =XNPV(B1,B2:B6,C2:C6)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =XNPV(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- Values and dates must have the same nonzero count; each date must be valid in the workbook date system. Rate must be greater than -1 and the vector is bounded to 10,000 entries.

#### `fx.XOR`

Return TRUE when an odd number of up to 255 scalar conditions are true; array-valued logical arguments remain outside the bounded evaluator.

**Examples:**

- =XOR(A1>0,B1>0,C1>0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =XOR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.YEAR`

Return the year component of a serial in the workbook's 1900 or 1904 date system.

**Examples:**

- =YEAR(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =YEAR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `importXlsxWithOfficeKit`

Import XLSX bytes through OfficeKit with editable core cells, formulas, styles, ordinary tables, PNG/JPEG pictures, validation, conditional formatting, threaded-comment roots with direct replies, bar/line/pie/area/doughnut charts, marker-only numeric-X/Y scatter charts, bounded numeric-X/Y/positive-Size bubble charts, and recognized PivotTables with exact item or absolute whole-day date filters. Imported data-table and dynamic-array topology is source-bound and read-only. A recognized source-bound QueryTable can only disable automatic refresh through table.setQueryRefreshPolicy; a recognized connection can only disable an explicit on-load refresh through workbook.disableConnectionRefreshOnLoad; a recognized uniquely owned Pivot cache can only disable an explicit on-load refresh through pivot.disableRefreshOnLoad; commands, fields, sorts, topology, non-marker scatter styles, noncanonical bubble profiles, nested/branched replies, mentions, other Pivot configuration/data/output, non-reversible sparkline graphs, and other advanced package content remain source-bound and read-only.

**Schema parameters:**

- `input` (FileBlob|Uint8Array|ArrayBuffer) required — XLSX package bytes.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `workbook` (Workbook) — Imported bounded workbook facade with editable core objects, canonical Office 2010 sparkline groups, bounded dynamic-array metadata, and source/opaque package evidence. A recognized QueryTable permits only table.setQueryRefreshPolicy automatic-refresh hardening; a recognized connection permits only workbook.disableConnectionRefreshOnLoad from explicit true to false; a recognized uniquely owned Pivot cache permits only pivot.disableRefreshOnLoad from explicit true to false; imported dynamic-array topology, commands, fields, sorts, QueryTable topology, other Pivot configuration/data/output, non-reversible sparkline graphs, and unsupported package graphs are exposed only for inspection or preserved unchanged.

#### `invokeOfficeKit`

Advanced experimental byte-boundary API for invoking the public OfficeKit codec protocol with generated wire-message objects.

**Schema parameters:**

- `request` (object) required — Generated public CodecRequest wire-message initializer. Prefer the typed XLSX helpers unless implementing codec infrastructure.

**Schema returns:**

- `response` (object) — Decoded public CodecResponse wire message; structured codec failures throw OfficeKitCodecError.

#### `officeKitStatus`

Lazily initialize the bundled OfficeKit NativeAOT codec and report its backend, target, transport, protocol, assembly, and integrity manifest.

**Schema returns:**

- `status` (object) — Bundled OfficeKit NativeAOT codec status with backend, target, transportVersion, protocolVersion, assemblyName, and integrity manifest.

#### `pivot.disableRefreshOnLoad`

On one recognized imported PivotTable with a uniquely owned cache and explicit refreshOnLoad=true, set only that cache root switch to false while preserving the complete Pivot graph and every other cache attribute.

**Schema returns:**

- `pivot` (WorksheetPivotTable) — One-way source-bound safety operation. It accepts only a recognized imported PivotTable with sourceCapabilities.refreshOnLoadHardenable=true and explicit refreshOnLoad=true, then changes only its cache definition root attribute to false. The cache residual, PivotTable XML, cache records, source values, cached worksheet output, cache ownership, relationships, identities, and every other policy field are re-proven; absent/already-false/shared/ambiguous/changed inputs fail closed. It does not run a refresh or prevent manual, macro, external-data, or other host-triggered refreshes.

#### `pivot.sourceCapabilities`

Inspect whether a PivotTable is source-bound and whether its uniquely owned imported cache can receive the one-way refreshOnLoad hardening operation.

**Schema returns:**

- `capabilities` (object) — Fresh { sourceBound, refreshOnLoadHardenable } evidence. refreshOnLoadHardenable is true only for a recognized imported PivotTable whose cache is uniquely owned, whose cache root explicitly has refreshOnLoad=true, and whose package binding is still eligible for the narrow hardening operation. It does not make the PivotTable generally editable.

#### `range.clear`

Clear range contents, formats, or both without silently changing validations, dimensions, or other package graphs.

**Schema parameters:**

- `applyTo` (string) — contents, formats, or all (default).

**Schema returns:**

- `result` (undefined) — No return value. Formula/spill topology is detached when contents are cleared.

#### `range.conditionalFormats.add`

Add a conditional formatting rule; cellIs/expression/containsText/colorScale plus standard dataBar/iconSet rules cross the public model and OfficeKit, with computedStyle inspect records, layout JSON visuals, SVG preview, and native XLSX rendering.

**Examples:**

- range.conditionalFormats.add('cellIs', { operator: 'greaterThan', formula: 10, format: { fill: 'green' } })
- range.conditionalFormats.add('dataBar', { color: '#2563eb', thresholds: ['min', 'max'], showValue: true })
- range.conditionalFormats.add('iconSet', { iconSet: '3TrafficLights1', thresholds: [0, '50%', '80%'], reverse: false })
- range.conditionalFormats.addColorScale({ colors: ['#fee2e2', '#fef3c7', '#22c55e'] })

**Schema parameters:**

- `ruleType` (string) required — cellIs, expression, containsText, colorScale, dataBar, or iconSet.
- `formula` (string|number) — Rule formula or scalar threshold. Omit for containsText; the range facade derives the required relative SEARCH formula.
- `text` (string) — Required search text for containsText rules.
- `operator` (string) — Comparison operator for cellIs rules.
- `format` (object) — Style patch applied when the rule matches.
- `colors` (string[]) — Two or three colors for colorScale rules.
- `color` (string|object) — RGB or symbolic Spreadsheet color for a standard gradient dataBar.
- `thresholds` (Array<string|number|object>) — Typed min/max/num/percent/percentile cfvo thresholds: exactly two for dataBar and one per icon for iconSet.
- `iconSet` (string) — One of the 17 base SpreadsheetML icon-set names. Office 2010 x14-only 3Triangles, 3Stars, and 5Boxes fail closed.
- `showValue` (boolean) — Show the formatted cell value beside the data bar or icon; false renders only the visual.
- `reverse` (boolean) — Reverse a built-in icon set's visual order.
- `gradient` (boolean) — Standard data bars are gradient-filled. false requires x14 and fails closed in this profile.

**Schema returns:**

- `conditionalFormat` (object) — Inspectable conditional-format rule with stable id.

#### `range.copyFrom`

Copy values, formulas, or complete cells from an equally sized or evenly tiling source range with relative A1 translation.

**Schema parameters:**

- `sourceRange` (Range) required — Source range whose row/column dimensions must evenly tile the destination.
- `mode` (string) — values, formulas, or all (default). Relative A1 formulas translate per destination cell.

**Schema returns:**

- `result` (undefined) — No return value; the destination range is updated transactionally in memory.

#### `range.copyTo`

Copy this range to an equally sized or evenly tiled destination range.

**Schema parameters:**

- `destinationRange` (Range) required — Destination range evenly tiled by this source range.
- `mode` (string) — values, formulas, or all (default).

**Schema returns:**

- `result` (undefined) — No return value; equivalent to destinationRange.copyFrom(sourceRange, mode).

#### `range.dataValidation`

Assign a list, whole, decimal, date, time, text-length, or custom-formula validation rule to a range, including bounded input prompts, error alerts, blank policy, and intuitive list-arrow visibility; use sheet.dataValidations.add({ range, rule }) for the collection form.

**Schema parameters:**

- `type` (string) required — Validation type: list, whole, decimal, date, time, textLength, or custom.
- `values` (unknown[]) — One through 256 non-empty, comma-free, control-safe inline list values whose quoted SpreadsheetML formula is at most 255 characters; list rules may use formula1 instead.
- `formula1` (string|number) — Primary validation formula/value.
- `formula2` (string|number) — Secondary formula/value for between rules.
- `operator` (string) — between, notBetween, equal, notEqual, lessThan, lessThanOrEqual, greaterThan, or greaterThanOrEqual.
- `allowBlank` (boolean) — Whether blank cells pass validation. Omission keeps the source-free compatibility default true.
- `showInputMessage` (boolean) — Show the bounded prompt when the cell is selected.
- `promptTitle` (string) — Input-prompt title, at most 32 characters.
- `prompt` (string) — Input-prompt message, at most 255 characters.
- `showErrorMessage` (boolean) — Show an error alert when entered data fails the rule.
- `errorTitle` (string) — Error-alert title, at most 32 characters.
- `error` (string) — Error-alert message, at most 255 characters.
- `errorStyle` (string) — stop, warning, or information.
- `showDropdown` (boolean) — For list rules, true means the in-cell arrow is visible. This deliberately hides SpreadsheetML's inverted showDropDown encoding.

**Schema returns:**

- `validation` (object) — Inspectable/editable bounded data-validation rule anchored to one contiguous range. Imported unsupported extension or multi-area graphs remain source-bound and unchanged.

#### `range.displayFormulas`

Read displayed A1 formulas, including the anchor formula projected across non-editable dynamic-array or legacy-array result cells.

**Schema returns:**

- `formulas` (string[][]) — A1 display-formula matrix, projecting spill/array anchors into non-editable result cells.

#### `range.fillDown`

Copy top-row contents and formatting down the range while translating relative A1 formula references.

**Schema returns:**

- `range` (Range) — The same range after top-row contents/formats are filled down with relative formula translation.

#### `range.fillRight`

Copy left-column contents and formatting right across the range while translating relative A1 formula references.

**Schema returns:**

- `range` (Range) — The same range after left-column contents/formats are filled right with relative formula translation.

#### `range.format`

Assign cell styles, symbolic theme/tint/indexed colors, patterned fills, native dimensions, pixel sizing, and hidden axes through a live range format facade.

**Examples:**

- sheet.getRange('A1:D1').format = { fill: '#0f172a', font: { bold: true }, columnWidth: 18, rowHeight: 24 }
- sheet.getRange('A1:D20').format.columnWidthPx = 120

**Schema parameters:**

- `fill` (string|object) — Solid color or { patternType, foreground, background }; colors accept RGB strings or { theme|indexed|auto, tint } references.
- `font` (object) — Font properties: bold, italic, underline, strike, color, size, and name. Color accepts RGB or symbolic SpreadsheetML references.
- `numberFormat` (string) — Excel number format code.
- `alignment` (object) — horizontal, vertical, wrapText, textRotation, indent, shrinkToFit, and readingOrder options.
- `border` (object) — A shared { style, color } border or per-edge records; colors accept RGB or theme/tint/indexed/auto references.
- `protection` (object) — Cell locked and hidden flags preserved through SpreadsheetML style records.
- `columnWidth` (number) — Column width in Excel character units for every column intersecting the range.
- `columnWidthPx` (number) — Column width in CSS pixels, converted with the public SpreadsheetML maximum-digit-width formula.
- `rowHeight` (number) — Row height in points for every row intersecting the range.
- `rowHeightPx` (number) — Row height in CSS pixels, converted at 96 DPI.
- `columnHidden` (boolean) — Hide or show every column intersecting the range.
- `rowHidden` (boolean) — Hide or show every row intersecting the range.

**Schema returns:**

- `range` (Range) — The formatted range facade.

#### `range.format.autofitColumns`

Measure displayed range values deterministically and set native best-fit widths on each selected column.

**Schema returns:**

- `range` (Range) — The same range after deterministic native best-fit column widths are applied.

#### `range.format.autofitRows`

Measure explicit/wrapped range text deterministically and set native custom heights on each selected row.

**Schema returns:**

- `range` (Range) — The same range after deterministic custom row heights are applied.

#### `range.formulaInfos`

Read per-cell stored/projected formula metadata with editability, spill/array source, anchor, and reference evidence.

**Schema returns:**

- `formulaInfos` (Array<Array<object|null>>) — Stored or projected per-cell formula evidence with kind, display, editability, source, anchor, and ref where applicable.

#### `range.formulasR1C1`

Read or assign R1C1 formulas relative to each target cell while storing canonical A1 formulas.

**Schema parameters:**

- `formulas` (string[][]) — R1C1 formulas relative to each target cell; blank strings clear formulas.

**Schema returns:**

- `formulas` (string[][]) — R1C1 formula matrix; stored formulas remain canonical A1.

#### `range.getCell`

Select one zero-based cell relative to the current range.

**Schema parameters:**

- `row` (number) required — Zero-based row offset within the current range.
- `column` (number) required — Zero-based column offset within the current range.

**Schema returns:**

- `range` (Range) — One-cell relative range.

#### `range.getColumn`

Select one zero-based column relative to the current range.

**Schema parameters:**

- `column` (number) required — Zero-based column offset within the current range.

**Schema returns:**

- `range` (Range) — One-column relative range spanning the current rows.

#### `range.getCurrentRegion`

Expand to the contiguous data region bounded by fully blank rows and columns.

**Schema returns:**

- `range` (Range) — Contiguous region bounded by fully blank rows and columns.

#### `range.getRangeByIndexes`

Select a bounded zero-based subrange relative to the current range.

**Schema parameters:**

- `startRow` (number) required — Zero-based row offset within the current range.
- `startColumn` (number) required — Zero-based column offset within the current range.
- `rowCount` (number) required — Positive subrange row count.
- `columnCount` (number) required — Positive subrange column count.

**Schema returns:**

- `range` (Range) — Bounded relative subrange.

#### `range.getRow`

Select one zero-based row relative to the current range.

**Schema parameters:**

- `row` (number) required — Zero-based row offset within the current range.

**Schema returns:**

- `range` (Range) — One-row relative range spanning the current columns.

#### `range.merge`

Merge the target range as one region or as separate row-wise regions when across=true.

**Schema parameters:**

- `across` (boolean) — Merge each target row independently when true.

**Schema returns:**

- `range` (Range) — The same range after merge creation.

#### `range.offset`

Return an equally sized range shifted by zero-based row and column offsets, rejecting worksheet overflow.

**Schema parameters:**

- `rowOffset` (number) required — Signed row offset.
- `columnOffset` (number) required — Signed column offset.

**Schema returns:**

- `range` (Range) — Equally sized shifted range within XLSX bounds.

#### `range.resize`

Return a range at the same upper-left cell with explicit positive row and column counts.

**Schema parameters:**

- `rowCount` (number) required — Positive output row count.
- `columnCount` (number) required — Positive output column count.

**Schema returns:**

- `range` (Range) — Resized range with the same upper-left cell.

#### `range.setNumberFormat`

Assign one number format or an evenly tiling matrix of Excel-invariant number-format codes.

**Schema parameters:**

- `format` (string|string[][]) required — Excel-invariant number-format code or an evenly tiling format matrix.

**Schema returns:**

- `range` (Range) — The same range after number-format assignment.

#### `range.unmerge`

Remove merged regions intersecting the target range.

**Schema returns:**

- `range` (Range) — The same range after intersecting merges are removed.

#### `range.write`

Write a mixed matrix or one explicit values/formulas/formulasR1C1 payload from the range anchor and return the actual written range.

**Schema parameters:**

- `value` (unknown[][]|unknown[]|object) required — Mixed values/formulas matrix, or exactly one of { values }, { formulas }, or { formulasR1C1 }.

**Schema returns:**

- `range` (Range) — Actual rectangular range written from the receiver's upper-left cell.

#### `range.writeValues`

Write a one- or two-dimensional value matrix from the range anchor.

**Schema parameters:**

- `values` (unknown[][]|unknown[]) required — One row or a rectangular value matrix written from the range anchor.

**Schema returns:**

- `result` (undefined) — No return value; inspect the target range after writing.

#### `sheet.charts.add`

Create an inspectable worksheet chart from a range or config; setData(range) infers category series, scatter per-series numeric xValues/xFormula plus y values/formula, or one exact X/Y/positive-Size bubble series. series.fill sets an explicit #RRGGBB solid color, series.line sets bounded RGB color/dash/width (series.stroke is an alias), line/scatter markers set direct symbol/size/RGB fill/bounded outline semantics, lineOptions controls standard/stacked/percent-stacked grouping, smooth interpolation, and direct vary-colors behavior, dataLabels controls plot-level value/category/series-name visibility and bounded position, and xAxis/yAxis configure primary titles, formats, intervals, and linear value bounds. Bar and line series accept up to 16 bounded native linear, exponential, logarithmic, power, polynomial, or moving-average trendlines plus one bounded native errorBars projection with fixed/percentage/standard-deviation/standard-error/custom semantics, one-/two-sided values, optional XLSX formula caches, cap policy, and bounded RGB line. Imported trendline count and error-bar presence are fixed; unsupported labels/extensions/unknown children/complex lines remain source-owned. Marker-only scatter rejects series.line/stroke and writes an explicit native no-fill series outline; use marker.line for marker borders. Bubble charts use two numeric axes and reject ambiguous range shortcuts or nonpositive sizes.

**Schema parameters:**

- `chartType` (string) required — Canonical OfficeKit XLSX chart type: bar, line, pie, area, doughnut, scatter, or bubble. Other model names fail closed on export.
- `source` (Range|object) — Source range or explicit chart config.
- `title` (string) — Chart title.
- `titleTextStyle` (object) — Optional chart-title style with fontSize from 1 through 4000 points.
- `lineOptions` (object) — Line-chart-only { grouping?, smooth?, varyColors? }. grouping is standard, stacked, or percentStacked; omission authors the standard default. smooth preserves explicit false as native c:smooth val=0. varyColors=true authors direct c:varyColors val=1; false or omission removes that optional node.
- `dataLabels` (boolean|object) — Optional plot-level labels. A boolean controls showValue; an object accepts boolean showValue/showCategoryName, optional presence-aware showSeriesName, and position: bestFit, bottom, center, insideBase, insideEnd, left, outsideEnd, right, or top. Per-series/per-point labels, number formats, and label text styles remain outside this bounded profile.
- `categories` (string[]) — Explicit shared categories for category charts. Scatter and bubble require this to stay empty and use per-series numeric xValues.
- `series` (object[]) — Explicit series definitions with name, optional numeric values/formula, category-chart categoryFormula, scatter/bubble numeric xValues/xFormula, and bubble-only positive bubbleSizes/bubbleSizeFormula exactly aligned with X/Y point counts. Optional #RRGGBB solid fill and optional line { fill, style, width } are supported; line/scatter marker { symbol, size, fill, line } remains marker-only. Bar and line series accept up to 16 trendlines with type linear/exp/log/power/poly/movingAvg, optional name, type-specific order/period, half-category forward/backward forecasts, intercept, equation/R-squared flags, and bounded RGB line. They also accept one errorBars object with reference type standardError/percentage/standardDeviation/none or canonical direction x/y, type both/minus/plus, valueType fixedVal/percentage/stdDev/stdErr/cust, bounded value or exact-count non-negative custom side values/formulas/caches, cap policy, and bounded RGB line. Imported trendline count and error-bar presence are fixed; labels, extensions, unknown children, malformed caches, and complex/theme lines remain source-owned. When internal range formulas are present, inspect/render/OfficeKit export resolve live category or numeric X/Y/Size caches from those cells. line.fill and marker.fill are #RRGGBB; both line objects use style solid, dashed, dotted, dash-dot, or dash-dot-dot and width 0 through 1584 points. Marker-only scatter rejects the series-level line/stroke aliases and uses marker.line only for marker borders. bubble3D, negative bubbles, custom scale, and non-area sizing are source-bound/read-only. marker.symbol is none, dot, circle, square, diamond, triangle, x, star, plus, or dash; marker.size is an integer from 2 through 72. stroke { color, style, weight } is a series-line compatibility alias and must not conflict with line.
- `xAxis` (object) — Primary text category axis with title.text, textStyle.fontSize, numberFormatCode, and tickLabelInterval; scatter and bubble instead use a numeric value axis with min, max, and majorUnit. Pie and doughnut charts have no axes.
- `yAxis` (object) — Primary numeric value axis with title.text, tick-label textStyle.fontSize, numberFormatCode, min, max, and majorUnit; tickLabelInterval is accepted as a compatibility alias for majorUnit. Pie and doughnut charts have no axes.
- `position` (object) — Pixel chart frame.

**Schema returns:**

- `chart` (WorksheetChart) — Editable worksheet chart facade.

#### `sheet.dataTables.__getDefinitions`

Return defensive inspectable definitions for the worksheet's canonical What-If data tables, including result range, native anchor, inputs, orientation, and display formula.

**Schema returns:**

- `definitions` (object[]) — Fresh defensive copies with zero-based result-range bounds, formulaRef, anchor, normalized rowInput/columnInput, rowOriented, twoVariable, and displayFormula. Mutating the returned objects does not mutate the worksheet.

#### `sheet.dataTables.add`

Create a canonical native Excel What-If data table from a rectangular formula/input grid and one row input, one column input, or both. Excel or another compatible host calculates the result values; the JavaScript evaluator does not simulate TABLE.

**Schema parameters:**

- `range` (string|Range) required — A rectangular A1 range at least 2x2. Its top-left cell must contain the formula to evaluate; its first row and first column contain substitution values, and the remaining rectangle is the native result range.
- `rowInput` (string) — Optional same-sheet single-cell A1 input reference. With no columnInput this authors a row-oriented one-variable table.
- `columnInput` (string) — Optional same-sheet single-cell A1 input reference. With no rowInput this authors a column-oriented one-variable table; provide both inputs for a two-variable table.

**Schema returns:**

- `result` (undefined) — No return value. Source-free canonical tables are authored as native t=dataTable formulas. Imported topology, input bindings, and orientation remain source-bound and read-only; unsupported or overlapping graphs fail closed without fallback.

#### `sheet.images.add`

Create an inspectable worksheet image from a data URL, URI, or prompt with one-cell, two-cell, or absolute pixel geometry plus optional percentage crop, bounded grayscale/luminance/opacity effects, rotation, and horizontal/vertical flips.

**Schema parameters:**

- `dataUrl` (string) — Embedded image data URL.
- `uri` (string) — External image URI metadata.
- `prompt` (string) — Generation/source prompt metadata.
- `alt` (string) — Alternative text.
- `anchor` (object) — One-cell { from, extent }, two-cell { type:'twoCell', from, to, editAs? }, or page-relative { type:'absolute', position:{leftPx,topPx}, extent } geometry. Cell markers use 0-based row/col plus optional rowOffsetPx/colOffsetPx; editAs is twoCell, oneCell, or absolute.
- `crop` (object) — Optional { leftPercent, topPercent, rightPercent, bottomPercent } source rectangle. Each signed offset is -100 through 100; opposing sums must remain below 100. Positive values inset and negative values outset.
- `effects` (object) — Optional { grayscale, brightnessPercent, contrastPercent, opacityPercent } profile. Brightness/contrast are -100 through 100; opacity is 0 through 100. OfficeKit maps it to bounded DrawingML blip effects.
- `transform` (object) — Optional { rotationDegrees, flipHorizontal, flipVertical } picture transform. Rotation is -360 through 360 degrees at DrawingML 1/60000-degree precision; flip booleans preserve explicit false values.
- `fit` (string) — contain or cover intent.

**Schema returns:**

- `image` (WorksheetImage) — Editable worksheet image facade.

#### `sheet.pivotTables.add`

Create a native bounded XLSX PivotTable with derived cached output, cache records, exact axis-item filters, and absolute whole-day date conditions. Relative-clock and sub-day filters remain model-only. Recognized imports are hash-bound and read-only except the separately verified refresh-on-load hardening primitive.

**Schema parameters:**

- `name` (string) — Stable pivot name.
- `sourceRange` (string|Range) required — Source data range.
- `targetRange` (string|Range) required — Destination anchor/range.
- `rowFields` (string[]) — Ordered row field names. Native source-free OfficeKit authoring accepts 1 through 8 fields in a tabular, no-automatic-subtotal profile.
- `columnFields` (string[]) — Column field names. Native source-free OfficeKit authoring currently accepts zero or one.
- `valueFields` (object[]) — One through 32 value-field definitions. Each names a source field and sum/count/average/min/max aggregation; repeated source fields with different aggregations are allowed.
- `rowGrandTotals` (boolean) — Add a native grand-total column and derived cached values when a column field is present.
- `columnGrandTotals` (boolean) — Add a native grand-total row and derived cached values.
- `groupFields` (object[]) — Derived group fields with unique name/sourceField. Calendar/time groupBy values years/quarters/months/days/hours/minutes/seconds form OOXML base/par hierarchies and accept bounded groupInterval values; range uses numeric startNum/endNum/groupInterval buckets; discrete uses named groups of source items.
- `calculatedFields` (object[]) — Calculated value fields over grouped source-field sums with arithmetic, percent, concatenation, comparisons, string/boolean constants, 12 bounded numeric functions, AND/OR/NOT, lazy IF/IFERROR/IFNA, NA, ISERROR/ISNUMBER/ISTEXT, Excel Compatibility Version 2 surrogate-aware LEN/LEFT/RIGHT/MID, LOWER/UPPER/ASCII-space TRIM, and workbook-date-system-aware DATE/YEAR/MONTH/DAY/EDATE/EOMONTH/DAYS/WEEKDAY/TIME/HOUR/MINUTE/SECOND/NETWORKDAYS/WORKDAY/NETWORKDAYS.INTL/WORKDAY.INTL. Business-day functions accept standard or international weekend rules and one optional scalar holiday. Accepts [Field] or quoted field references; cell references, holiday arrays/ranges, calculated-field chaining, and non-whitelisted functions are rejected.
- `filters` (object|object[]) — Axis filters. Exact include/exclude lists of 1 through 1024 string, finite-number, boolean, or null items author standard hidden PivotField items. Absolute dateEqual/dateNotEqual/dateOlderThan/dateOlderThanOrEqual/dateNewerThan/dateNewerThanOrEqual/dateBetween/dateNotBetween filters with useWholeDay=true author schema-valid native x:filters and date-typed caches from ISO date/Date source cells. useWholeDay=false ISO date-time thresholds and relative UTC types yesterday/today/tomorrow, last/this/next week/month/quarter/year, and yearToDate remain model/preview-only; relative filters accept deterministic asOf and use Monday-start ISO weeks.
- `refreshPolicy` (object) — OOXML cache policy: refreshOnLoad, saveData, enableRefresh, invalid, missingItemsLimit, refreshedBy, and refreshedDateIso.

**Schema returns:**

- `pivot` (WorksheetPivotTable) — Native XLSX authoring is bounded to 1 through 8 tabular row fields without automatic subtotals, optional one column field, 1 through 32 sum/count/average/min/max value fields, and one exact-item or absolute whole-day date filter per axis field. Multiple values use the canonical x=-2 data-layout axis. Cached output is a derived projection; grouping, calculated fields, relative/sub-day/other conditional filters, compact/subtotal-bearing multi-row layouts, and other richer profiles remain model-only and fail closed on native export. Recognized imports expose semantics but keep config, source data, cached output, and topology read-only, except a separately proven cache-level refreshOnLoad true-to-false hardening capability.

#### `sheet.sparklineGroups.add`

Create standard Office 2010 line/column/stacked sparkline groups for inspect, SVG preview, and OfficeKit XLSX export. Source-free groups use reversible one-dimensional target/source mappings; recognized imported groups support fixed-topology semantic edits while unsupported native graphs remain source-bound.

**Schema parameters:**

- `type` (string) — line, column, or stacked.
- `targetRange` (string|Range) required — One-dimensional destination range. Each target cell receives one native sparkline.
- `sourceData` (string|Range) required — One-dimensional source for one target, or a reversible rectangle whose rows/columns map exactly to the target cells.
- `dateAxisRange` (string|Range) — Optional one-dimensional date axis with one entry per sparkline point.
- `seriesColor` (string|object) — RGB or native theme/indexed/automatic series color.
- `negativeColor` (string|object) — Optional negative-value color.
- `axisColor` (string|object) — Optional horizontal-axis color.
- `markersColor` (string|object) — Optional ordinary-marker color.
- `firstMarkerColor` (string|object) — Optional first-point marker color.
- `lastMarkerColor` (string|object) — Optional last-point marker color.
- `highMarkerColor` (string|object) — Optional high-point marker color.
- `lowMarkerColor` (string|object) — Optional low-point marker color.
- `lineWeight` (number) — Positive line weight in points; defaults to 1.
- `displayHidden` (boolean) — Whether hidden source cells contribute to the native sparkline.
- `displayEmptyCellsAs` (string|number) — span/connect, gap, zero, or compatible numeric value 1, 2, or 3.
- `markers` (object) — Optional show/high/low/first/last/negative marker booleans.
- `axis` (object) — Optional manualMin/manualMax, minMode/maxMode (individual/group/custom or 0/1/2), showAxis, and rightToLeft settings.

**Schema returns:**

- `sparkline` (SparklineGroup) — Editable standard Office 2010 x14 sparkline group for inspect/layout/SVG preview and OfficeKit XLSX I/O. Source-free groups use the documented reversible mapping; imported canonical groups are source-bound and permit property edits without topology changes. Unsupported native sparkline graphs remain opaque and unchanged.

#### `sheet.tables.add`

Create an ordinary worksheet table over an A1 range with headers, columns, totals metadata, style, and bounded filtering/sorting. QueryTable bindings cannot be authored; recognized imported bindings expose only table.setQueryRefreshPolicy for one-way automatic-refresh hardening, while all other QueryTable edits fail closed.

**Schema parameters:**

- `range` (string|Range) required — A1 range or range facade.
- `hasHeaders` (boolean) — Whether the first row contains headers.
- `name` (string) — Stable Excel table name.
- `style` (string) — Table style name.
- `columnNames` (string[]) — Compatibility projection of table-column names.
- `columnDefinitions` (object[]) — Rich columns with name, calculatedColumnFormula/array, and totalsRowFunction/label/formula/array metadata.
- `filters` (object[]) — Zero-based table-column exact-value/blank, grouped-date/calendar, one/two-criterion custom, dynamic type/threshold, top/bottom item/percent, standard icon-set, or stable cell-fill/font-color AutoFilters; color filters use kind='color', target='cell'|'font', and color without exposing dxfId.
- `sortState` (object) — Bounded row-oriented value/icon/color-sort state with reference, caseSensitive, optional sortMethod ('none'|'pinYin'|'stroke'), and ordered single-column conditions; value conditions may carry customList. Table AutoFilter sorts reject columnSort per SpreadsheetML.
- `showTotals` (boolean) — Expose the totals row required by totals metadata.

**Schema returns:**

- `table` (WorksheetTable) — Editable ordinary worksheet table facade. QueryTable bindings are import-only and read-only except that a recognized imported binding permits table.setQueryRefreshPolicy to harden automatic-refresh switches; all connection, command, field, sort, refresh-history, topology, and other QueryTable changes fail closed.

#### `SpreadsheetFile.exportCsv`

Export one worksheet or range as UTF-8 CSV, using calculated values unless formula output is explicitly requested.

**Schema parameters:**

- `workbook` (Workbook) required — Workbook facade to serialize.
- `sheetName` (string) — Worksheet name; defaults to the first sheet.
- `range` (string) — Optional A1 range.
- `formulas` (boolean) — Emit formulas instead of calculated values where present.
- `lineEnding` (string) — LF or CRLF output; defaults to CRLF.
- `includeBom` (boolean) — Prefix a UTF-8 BOM; defaults to false.
- `maxBytes` (number) — Maximum encoded output bytes; defaults to 10 MiB.
- `maxRows` (number) — Maximum exported rows; defaults to 100000.
- `maxColumns` (number) — Maximum exported columns; defaults to 16384.

**Schema returns:**

- `blob` (FileBlob) — UTF-8 CSV FileBlob.

#### `SpreadsheetFile.exportDelimited`

Serialize one workbook sheet/range as bounded CSV/TSV text with calculated-value defaults and RFC-style quoting.

**Schema parameters:**

- `workbook` (Workbook) required — Workbook facade to serialize.
- `delimiter` (string) — Single field delimiter; defaults to comma.
- `sheetName` (string) — Worksheet name; defaults to the first sheet.
- `range` (string) — Optional A1 range; defaults to the used range.
- `formulas` (boolean) — Emit formulas instead of calculated values where present; defaults to false.
- `lineEnding` (string) — LF or CRLF output; defaults to CRLF.
- `includeBom` (boolean) — Prefix a UTF-8 BOM; defaults to false.
- `maxBytes` (number) — Maximum encoded output bytes; defaults to 10 MiB.
- `maxRows` (number) — Maximum exported rows; defaults to 100000.
- `maxColumns` (number) — Maximum exported columns; defaults to 16384.

**Schema returns:**

- `blob` (FileBlob) — UTF-8 CSV/TSV FileBlob with row/column metadata.

#### `SpreadsheetFile.exportTsv`

Export one worksheet or range as UTF-8 tab-separated text with RFC-style quoting where needed.

**Schema parameters:**

- `workbook` (Workbook) required — Workbook facade to serialize.
- `sheetName` (string) — Worksheet name; defaults to the first sheet.
- `range` (string) — Optional A1 range.
- `formulas` (boolean) — Emit formulas instead of calculated values where present.
- `lineEnding` (string) — LF or CRLF output; defaults to CRLF.
- `includeBom` (boolean) — Prefix a UTF-8 BOM; defaults to false.
- `maxBytes` (number) — Maximum encoded output bytes; defaults to 10 MiB.
- `maxRows` (number) — Maximum exported rows; defaults to 100000.
- `maxColumns` (number) — Maximum exported columns; defaults to 16384.

**Schema returns:**

- `blob` (FileBlob) — UTF-8 TSV FileBlob.

#### `SpreadsheetFile.exportXlsx`

Serialize a Workbook facade through the single bundled OfficeKit codec.

**Schema parameters:**

- `workbook` (Workbook) required — Workbook facade to recalculate and serialize.
- `recalculate` (boolean) — Recalculate formulas before serialization; defaults to true.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — Native OOXML XLSX package bytes.

#### `SpreadsheetFile.importCsv`

Import UTF-8 CSV bytes into an editable Workbook through the bounded delimited parser.

**Schema parameters:**

- `input` (FileBlob|Uint8Array|string) required — UTF-8 CSV text or bytes.
- `sheetName` (string) — Imported worksheet name.
- `coerceTypes` (boolean) — Convert unquoted boolean/numeric-looking cells; defaults to false.
- `maxBytes` (number) — Maximum encoded input bytes; defaults to 10 MiB.
- `maxRows` (number) — Maximum parsed rows; defaults to 100000.
- `maxColumns` (number) — Maximum parsed columns per row; defaults to 16384.

**Schema returns:**

- `workbook` (Workbook) — Imported editable workbook facade.

#### `SpreadsheetFile.importDelimited`

Parse bounded RFC-style CSV/TSV bytes into an editable Workbook, including quoted delimiters, escaped quotes, and embedded newlines.

**Schema parameters:**

- `input` (FileBlob|Uint8Array|string) required — UTF-8 delimited text or bytes.
- `delimiter` (string) — Single field delimiter; defaults to comma.
- `sheetName` (string) — Imported worksheet name; defaults to Sheet1.
- `coerceTypes` (boolean) — Convert unquoted boolean/numeric-looking cells; defaults to false.
- `maxBytes` (number) — Maximum encoded input bytes; defaults to 10 MiB.
- `maxRows` (number) — Maximum parsed rows; defaults to 100000.
- `maxColumns` (number) — Maximum parsed columns per row; defaults to 16384.

**Schema returns:**

- `workbook` (Workbook) — Imported editable workbook facade.

#### `SpreadsheetFile.importTsv`

Import UTF-8 tab-separated bytes into an editable Workbook through the bounded delimited parser.

**Schema parameters:**

- `input` (FileBlob|Uint8Array|string) required — UTF-8 TSV text or bytes.
- `sheetName` (string) — Imported worksheet name.
- `coerceTypes` (boolean) — Convert unquoted boolean/numeric-looking cells; defaults to false.
- `maxBytes` (number) — Maximum encoded input bytes; defaults to 10 MiB.
- `maxRows` (number) — Maximum parsed rows; defaults to 100000.
- `maxColumns` (number) — Maximum parsed columns per row; defaults to 16384.

**Schema returns:**

- `workbook` (Workbook) — Imported editable workbook facade.

#### `SpreadsheetFile.importXlsx`

Load XLSX through the single bundled OfficeKit codec into an editable Workbook facade.

**Schema parameters:**

- `xlsx` (FileBlob|Uint8Array) required — XLSX package bytes.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `workbook` (Workbook) — Imported workbook facade with editable core cells, formulas, styles, ordinary tables, images, basic charts, validation, conditional formatting, threaded-comment roots/direct replies, canonical Office 2010 sparkline groups, and bounded dynamic-array metadata. A recognized imported QueryTable may only use table.setQueryRefreshPolicy to disable automatic refresh, a recognized connection may only use workbook.disableConnectionRefreshOnLoad to turn explicit refreshOnLoad=true off, and a recognized uniquely owned Pivot cache may only use pivot.disableRefreshOnLoad for the same one-way safety transition; imported dynamic-array topology, commands, fields, sorts, topology, nested reply graphs, mentions, other Pivot configuration/data/output, non-reversible sparkline graphs, and unsupported package graphs remain source-bound and read-only.

#### `SpreadsheetFile.inspectDelimited`

Inspect bounded CSV/TSV bytes as file/row records with dimensions, delimiter, quoting, and formula-like cell evidence.

**Schema parameters:**

- `input` (FileBlob|Uint8Array|string) required — UTF-8 CSV/TSV text or bytes.
- `delimiter` (string) — Single field delimiter; defaults to comma.
- `maxBytes` (number) — Maximum encoded input bytes.
- `maxRows` (number) — Maximum parsed rows.
- `maxColumns` (number) — Maximum parsed columns per row.
- `maxPreviewRows` (number) — Maximum row records in bounded output; defaults to 20.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `inspection` (object) — Delimited-file summary, bounded row records, and NDJSON evidence.

#### `SpreadsheetFile.inspectXlsx`

Inspect bounded XLSX parts, content types, the required workbook/root officeDocument relationship, and namespace-aware source XML r:id/r:embed/r:link references after raw-input, part-count, decompression, and optional compression-ratio budgets; verifyCrc32 additionally checks ZIP entry CRCs.

**Schema parameters:**

- `xlsx` (FileBlob|Uint8Array) required — XLSX package bytes.
- `includeText` (boolean) — Include bounded XML/JSON/relationship previews.
- `maxPreviewChars` (number) — Maximum preview characters per textual part.
- `maxInputBytes` (number) — Maximum compressed input bytes checked before JSZip parses the package.
- `maxParts` (number) — Maximum package part count.
- `maxPartBytes` (number) — Maximum uncompressed bytes per part.
- `maxTotalBytes` (number) — Maximum total uncompressed package bytes.
- `maxCompressionRatio` (number) — Optional maximum declared uncompressed/compressed ZIP-entry ratio; zero or omitted disables this extra check.
- `verifyCrc32` (boolean) — Verify every ZIP entry CRC32 before inspecting package structure; use for untrusted retained inputs.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `package` (object) — XLSX package result with ok, issues, parts, records, and bounded NDJSON.

#### `SpreadsheetFile.patchXlsx`

Apply path-validated XLSX part patches, build worksheet/table/drawing/image/chart/pivot source references, and atomically reject dangling content types or relationships.

**Schema parameters:**

- `xlsx` (FileBlob|Uint8Array) required — XLSX package bytes.
- `patches` (array|object) required — Safe part edits with text, xml, json, bytes, content, remove, or delete.
- `maxInputBytes` (number) — Maximum compressed input bytes checked before JSZip parses the package.
- `maxPatchBytes` (number) — Maximum bytes per replacement part.
- `maxParts` (number) — Maximum resulting package part count.
- `maxPartBytes` (number) — Maximum uncompressed bytes per source or resulting part.
- `maxTotalBytes` (number) — Maximum total uncompressed source or resulting package bytes.
- `maxCompressionRatio` (number) — Optional maximum declared uncompressed/compressed ZIP-entry ratio; zero or omitted disables this extra check.
- `syncContentTypes` (boolean) — Synchronize inferred or explicit content-type declarations; defaults to true.
- `syncRelationships` (boolean) — Remove relationships to deleted parts and apply relationship recipes; defaults to true.
- `syncSourceReferences` (boolean) — Apply opt-in standard sourceReference XML mutations for supported semantic recipes; defaults to true.
- `validateResult` (boolean) — Validate final content types and relationships atomically; defaults to true. Set false only for deliberate invalid-package fixtures.
- `recipe` (string|object) — Standard OOXML part recipe with optional source/id/target and sourceReference fields; XLSX supports worksheet/table lists, pivot cache/record bindings, typed pivotTable relationships, and explicit-anchor drawing/image/chart nodes.
- `sourceReference` (boolean|object) — Opt-in source XML mutation. Image/chart objects require explicit anchor geometry; pivotCacheDefinition requires a unique cacheId; pivotCacheRecords binds the cache root to its records relationship.
- `relationship` (object) — Per-patch source/id/type/target/targetMode relationship recipe; explicit ID collisions require replaceExisting:true. relationships accepts an array.

**Schema returns:**

- `blob` (FileBlob) — Patched XLSX FileBlob with part/relationship/content-type/source-reference update counts and validation metadata.

#### `table.setQueryRefreshPolicy`

On one recognized imported QueryTable, monotonically disable automatic refresh without changing its connection, command, fields, sort, refresh history, or topology.

**Examples:**

- table.setQueryRefreshPolicy({ disableRefresh: true, backgroundRefresh: false, refreshOnLoad: false })

**Schema parameters:**

- `policy` (object) required — One or more of exactly { disableRefresh: true, backgroundRefresh: false, firstBackgroundRefresh: false, refreshOnLoad: false }. Unknown keys, unsafe values, an empty object, QueryTable authoring, and non-source-bound mutations fail closed.

**Schema returns:**

- `queryTable` (object) — The same recognized imported QueryTable projection after a one-way automatic-refresh hardening request. The source connection, command/credential metadata, fields, deleted-field history, refresh-local sort, topology, and unknown XML remain immutable and are re-proved before export.

**Notes:**

- This is a source-bound safety operation, not QueryTable authoring or general editing. Each supplied field has exactly one permitted value: disableRefresh: true; backgroundRefresh, firstBackgroundRefresh, and refreshOnLoad: false. Export proves the original query part, immutable connection part, and normalized XML residual before reparsing the result. Commands, credentials, connection bindings, fields, deleted-field history, refresh-local sort state, unknown XML, and every other root attribute remain immutable; unsupported or altered source graphs fail closed.

#### `thread.addReply`

Append a direct reply to an Office threaded-comment root with independent author/person/date/done metadata. Nested or branched reply graphs and mentions fail closed.

**Schema parameters:**

- `text` (string) required — Direct reply text.
- `author` (string) — Reply author; defaults to comments.setSelf or the root author.
- `id` (string) — Optional brace-delimited native comment GUID; otherwise OfficeKit derives one deterministically.
- `personId` (string) — Optional brace-delimited native person GUID.
- `person` (object) — Optional displayName/userId/providerId identity record.
- `date` (string) — Optional ISO-8601 reply timestamp.
- `done` (boolean) — Optional native reply done state.

**Schema returns:**

- `thread` (CommentThread) — The same thread with one appended direct reply. Setting parentId to another reply, adding mentions, or creating a branched/nested graph makes canonical export fail closed.

#### `workbook.auditAccessibility`

Audit worksheet images and charts for explicit meaningful/decorative classification and non-visible xdr:cNvPr title/description coverage. Native reading order and broader worksheet semantics remain manual checks; the report never claims Excel Accessibility Checker, WCAG, or PDF conformance.

**Schema parameters:**

- `maxChars` (number) — Maximum bounded NDJSON size across machine issues and manual-review records.

**Schema returns:**

- `report` (object) — A host-neutral report with machineCheckPassed, conformanceClaimed: false, manualReviewRequired, stable sheet/object locators, drawing counts, machine issues for unclassified or textless meaningful images/charts, and separate native reading-order/worksheet-semantics checks. It never claims Excel Accessibility Checker, WCAG, or PDF conformance.

#### `workbook.comments.addThread`

Create one root Office threaded comment per thread with GUID/person metadata, date, and resolved state; attach bounded direct replies with thread.addReply().

**Schema parameters:**

- `target` (Range|object) required — Target single-cell range or cell descriptor.
- `text` (string) required — Initial comment text.
- `author` (string) — Root comment author; defaults to comments.setSelf identity.
- `id` (string) — Optional stable model thread ID.
- `comment` (object) — Optional native root metadata: brace-delimited GUID id/personId, person record, ISO date, and done state.
- `resolved` (boolean) — Initial thread resolution state.

**Schema returns:**

- `thread` (CommentThread) — Attached Office threaded-comment root. Direct replies added through addReply cross canonical OfficeKit export/import; nested/branched replies and mentions fail closed.

#### `workbook.connections`

Inspect bounded non-secret metadata for imported database connections. Connections are source-bound; the sole mutation is workbook.disableConnectionRefreshOnLoad(connectionId) for an explicit imported refreshOnLoad=true value.

**Schema returns:**

- `connections` (object[]) — Recognized imported connection roots exposed for inspection. Count, order, identity, provider strings, commands, credentials, source paths, children, extensions, and every field except explicit refreshOnLoad=true remain source-bound; use workbook.disableConnectionRefreshOnLoad(connectionId) for that sole one-way safety mutation.

#### `Workbook.create`

Create an empty workbook with an explicit date system and optional native SpreadsheetML theme colors.

**Schema parameters:**

- `dateSystem` (string) — Excel serial-date system: '1900' (default) or '1904'.
- `date1904` (boolean) — Boolean alias for dateSystem; true selects the 1904 system.
- `theme` (object) — Theme name and dk1/lt1/dk2/lt2, accent1-accent6, hlink, and folHlink colors written to xl/theme/theme1.xml.
- `calculation` (object) — Optional bounded workbook calcPr policy; omitted means no authored calculation-properties element.

**Schema returns:**

- `workbook` (Workbook) — Empty editable workbook facade with a normalized date system.

#### `workbook.definedNames.add`

Create a workbook or sheet-scoped defined name over an A1 range; exported as native workbook.xml definedName and usable in formulas such as SUM(RevenueData).

**Examples:**

- workbook.definedNames.add('RevenueData', 'Sheet1!G2:G4')
- sheet.getRange('E3').formulas = [['=SUM(RevenueData)']]

**Options:**

- name
- refersTo
- scope/sheetName
- comment
- hidden

**Schema parameters:**

- `name` (string) required — Defined name.
- `refersTo` (string) required — Sheet-qualified A1 reference.
- `scope` (string) — Optional worksheet scope.
- `comment` (string) — Optional description.
- `hidden` (boolean) — Optional native hidden flag; explicit false is preserved.

**Schema returns:**

- `definedName` (DefinedName) — Created or updated defined-name facade.

**Returns:**

DefinedName facade with id/name/refersTo/scope

#### `workbook.disableConnectionRefreshOnLoad`

On one recognized imported connection with explicit refreshOnLoad=true, set that sole root switch to false without changing its command, credentials, topology, or any other connection state.

**Examples:**

- workbook.disableConnectionRefreshOnLoad(7)
- workbook.disableConnectionRefreshOnLoad('connection/7')

**Options:**

- connectionId

**Schema parameters:**

- `connectionId` (number|string) required — Positive native connection ID or canonical connection/<id> locator. The recognized imported source must explicitly have refreshOnLoad=true.

**Schema returns:**

- `connection` (object) — The same source-bound connection projection with refreshOnLoad=false. This does not execute external data or modify commands, credentials, provider/path metadata, background/keepAlive/interval/saveData policy, topology, or unknown XML.

**Returns:**

The same imported connection projection with refreshOnLoad: false

**Notes:**

- This is a source-bound one-way safety operation, not a general connection editor. It accepts only an imported connection whose validated source explicitly has refreshOnLoad=true. Export proves the full source ConnectionsPart and target element hashes, removes only that attribute from the normalized residual proof, reparses the output, and fails closed if any other connection semantic or XML content changes. It does not run a refresh, edit commands, credentials, provider strings, paths, child graphs, extensions, keepAlive, background, interval, saveData, connection order, or connection identity.

#### `workbook.fontFamilies`

Return a fresh sorted, case-insensitively deduplicated list of workbook default and explicit cell font families.

**Schema returns:**

- `families` (string[]) — Font-family inventory; mutating the returned array does not mutate the workbook.

#### `workbook.formulaGraph`

Return a bounded dependency graph of formula nodes, edges, dependents, cycles, formula errors, and syntax-input/reference-budget refusals for workbook QA.

**Schema parameters:**

- `recalculate` (boolean) — Recalculate before reading the graph; defaults to true.
- `maxChars` (number) — Maximum bounded NDJSON graph-record size.

**Schema returns:**

- `graph` (object) — Bounded formula nodes, edges, cycles, errors, syntax-input/reference-budget refusals, and NDJSON.

#### `workbook.inspect`

Emit bounded NDJSON records for workbook, connections, sheets, worksheet protections, tables, formulas, matches, comments, validations, conditional formats, and drawings; narrow with search/target anchors and shape fields with include/exclude.

**Examples:**

- workbook.inspect({ kind: 'formula', target: 'Sheet1!E2', include: 'formula,value,precedents' })

**Options:**

- kind
- search/searchTerm
- target/targetId/id/anchor
- before/after/context
- include/fields
- exclude/omit
- maxChars

**Schema parameters:**

- `kind` (string) — Comma-separated record kinds such as connection, formula, table, style, computedStyle, chart, image.
- `target` (string) — Stable ID, anchor, or A1 cell/range to slice results around.
- `search` (string) — Case-insensitive text filter over inspect records.
- `include` (string) — Comma-separated top-level fields to keep.
- `exclude` (string) — Comma-separated top-level fields to omit.
- `maxChars` (number) — Maximum NDJSON output size before truncation notice.

**Schema returns:**

- `ndjson` (string) — Bounded newline-delimited JSON records.
- `truncated` (boolean) — True when maxChars truncated the output.

**Returns:**

{ ndjson, truncated } bounded NDJSON records

#### `workbook.layoutJson`

Return workbook/worksheet layout JSON with cell, table, chart, image, sparkline, rule bounding boxes, and target/search context slicing.

**Schema parameters:**

- `sheetName` (string) — Optional worksheet selector.
- `range` (string) — Optional A1 layout range.
- `target` (string) — Stable target ID/anchor.
- `search` (string) — Case-insensitive layout-record filter.
- `before` (number) — Context records before matches.
- `after` (number) — Context records after matches.

**Schema returns:**

- `layout` (object) — Workbook/worksheet layout tree with cells and drawing/rule bounds.

#### `workbook.recalculate`

Recalculate bounded workbook formulas and dynamic-array spills, with dependency edges, cycles, errors, and syntax-input/reference-budget refusals.

**Schema returns:**

- `graph` (object) — Updated bounded formula dependency graph including cycles, errors, and syntax-input/reference-budget refusals.

#### `workbook.render`

Return a lightweight SVG preview for a sheet/range or layout JSON when called with { format: 'layout' }.

**Schema parameters:**

- `sheetName` (string) — Worksheet name; defaults to the active worksheet.
- `range` (string) — A1 preview range.
- `format` (string) — svg by default or layout.
- `target` (string) — Stable layout target ID/anchor.
- `search` (string) — Case-insensitive layout filter.

**Schema returns:**

- `blob` (FileBlob) — Worksheet SVG preview or workbook layout JSON.

#### `workbook.resolve`

Resolve stable workbook, source-bound connection, worksheet, table, pivot, chart, image, sparkline, rule, comment, and defined-name IDs.

**Schema parameters:**

- `id` (string) required — Stable workbook, sheet, table, pivot, chart, image, sparkline, rule, comment, or defined-name ID.

**Schema returns:**

- `object` (object|undefined) — Resolved editable facade/record or undefined.

#### `workbook.setCalculation`

Set bounded workbook-level SpreadsheetML calculation mode, on-save/full-recalculation flags, iterative-calculation limits, and full-precision policy.

**Examples:**

- workbook.setCalculation({ mode: 'automatic', fullCalculationOnLoad: true, forceFullCalculation: true })
- workbook.setCalculation({ mode: 'manual', iteration: { enabled: true, maxIterations: 100, maxChange: 0.001 } })

**Options:**

- mode
- calculateOnSave
- fullCalculationOnLoad
- forceFullCalculation
- iteration
- fullPrecision

**Schema parameters:**

- `mode` (string) — automatic, automaticExceptTables, or manual.
- `calculateOnSave` (boolean) — Request calculation when a host application saves the workbook.
- `fullCalculationOnLoad` (boolean) — Request a full calculation when a host application opens the workbook.
- `forceFullCalculation` (boolean) — Force full rather than dependency-only recalculation.
- `iteration` (object) — Optional { enabled, maxIterations, maxChange } circular-calculation policy.
- `fullPrecision` (boolean) — Calculate using stored values rather than displayed precision when true.

**Schema returns:**

- `workbook` (Workbook) — The same workbook with a bounded native workbook.xml calcPr policy.

**Returns:**

Workbook facade with bounded native calcPr policy

#### `workbook.setDateSystem`

Select the Excel 1900 or 1904 serial-date system for formula calculation and native workbookPr export.

**Schema parameters:**

- `dateSystem` (string|boolean) required — '1900' or false for the 1900 system; '1904' or true for the 1904 system.

**Schema returns:**

- `workbook` (Workbook) — The same workbook after changing its formula and OOXML date-system context.

#### `workbook.sharedArrayFormulas`

Import and export bounded shared, legacy-array, and source-free XLDAPR dynamic-array formula metadata. Imported dynamic-array anchors remain source-bound and read-only; malformed or topology-changing edits fail closed.

**Schema parameters:**

- `xlsx` (FileBlob|Uint8Array) — XLSX bytes containing shared, legacy-array, or XLDAPR dynamic-array formula records.
- `formula` (string) — Shared, legacy-array, or bounded source-free dynamic-array formula expression. Imported dynamic-array expressions are read-only.
- `ref` (string) — Shared group, legacy-array range, or dynamic spill range. Source-free dynamic ranges use the canonical XLDAPR profile; imported ranges remain source-bound.

**Schema returns:**

- `metadata` (object) — Shared/legacy metadata is bounded and editable; one canonical source-free XLDAPR dynamic-array profile is authored through OfficeKit. Imported dynamic-array metadata remains source-bound, and malformed, detached, or topology-changing edits fail closed.

#### `workbook.spillReferences`

Use a direct or defined-name A1# reference to consume only an anchor's current, unblocked dynamic spill matrix. Supported range consumers and a direct re-spill read the verified matrix; scalar/general-vector coercion returns #VALUE!, non-spilling anchors return #REF!, and graph/trace record one spillReference edge to the anchor.

**Examples:**

- =SUM(A1#)
- =MATCH(12,'Source Data'!A1#,0)
- =FILTER(A1#,A1#>10)
- =CurrentSpill

**Schema parameters:**

- `formula` (string) required — Formula containing a direct or defined-name A1# dynamic spill reference.
- `anchor` (string) — Optional explanatory A1 anchor such as Source Data!A1; the actual formula must carry #.

**Schema returns:**

- `value` (unknown|unknown[][]|#REF!|#VALUE!|#SPILL!) — Current model-calculation spill value. Only the documented range consumers and direct re-spill profile accept A1#; scalar/general-vector coercion is #VALUE!, a non-spilling anchor is #REF!, and imported dynamic-array package topology remains source-bound.

**Notes:**

- A1# is a model-calculation range reference, not source-free XLSX dynamic-array topology authoring. The evaluator recalculates the formula anchor before reading it, verifies a current rectangular spill of at most 10,000 cells, charges each read against the 20,000-cell formula total, and preserves one spillReference dependency edge to the anchor. A blocked/error anchor propagates its current error; an ordinary scalar/non-spilling anchor is #REF!.

#### `workbook.structuredReferences`

Evaluate Excel table references including sections, column ranges/unions, space intersections, escaped special-character headers, unqualified calculated-column references, and @/#This Row context while expanding exact table-cell precedents.

**Examples:**

- =SUM(TableName[Column])
- =SUM(TableName[[#Data],[First]:[Last]])
- =SUM(TableName[[First]:[Second]] TableName[[Second]:[Third]])
- =[Revenue]-[Cost]
- =TasksTable[@Revenue]
- =SUM(TasksTable[[#This Row],[Revenue]:[Cost]])
- =TasksTable['#Items]
- =TasksTable[Bracket'[Value']]

**Schema parameters:**

- `formula` (string) required — Formula containing an Excel table structured reference.
- `table` (string) — Worksheet table name; omitted only for a calculated-column reference inside that table.
- `selector` (string) required — Column, escaped special-character header, section, current-row, range, union, or space-intersection selector.

**Schema returns:**

- `value` (unknown) — Calculated scalar/array value with stable table-cell precedents.

**Notes:**

- Supports #Headers/#Data/#All/#Totals/#This Row and @, unqualified current-row references inside tables, contiguous column ranges, comma-separated column unions, space intersections over common cells, and apostrophe escaping for [, ], #, ', and @ in column headers. Disjoint intersections return #NULL!; current-row references outside the referenced table return #VALUE!.

#### `workbook.trace`

Return a formula precedent tree and bounded NDJSON trace for a target cell, with circular references and syntax-input/reference-budget refusals flagged.

**Schema parameters:**

- `reference` (string|Range) required — Target A1 reference, optionally sheet-qualified, or range facade.
- `maxDepth` (number) — Maximum precedent recursion depth; defaults to 8.
- `maxChars` (number) — Maximum bounded NDJSON trace size.

**Schema returns:**

- `trace` (object) — Precedent tree plus bounded flat NDJSON trace; oversized syntax or sources are reported rather than walked.

#### `workbook.verify`

Return bounded QA issues for source-bound connections, sheets, formulas (including syntax-input and reference-budget refusals), tables, charts, and comments.

**Schema parameters:**

- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — Workbook formula/structure/drawing/rule QA result, including syntax-input and reference-budget refusals.

#### `workbook.windows`

Access the ordered workbook-window collection; window 0 is the primary view used by legacy worksheet-selection APIs.

**Schema returns:**

- `windows` (WorkbookWindowCollection) — Ordered windows. Index 0 is the primary window; additional windows retain independent active and selected worksheet state.

#### `workbook.windows.add`

Append an additional workbook window with its own active worksheet and selected tab group.

**Schema parameters:**

- `activeWorksheet` (string|number|Worksheet) — Visible worksheet name, zero-based worksheet index, or worksheet object. Defaults to the primary window's active worksheet.
- `selectedWorksheets` (Array<string|number|Worksheet>) — Optional non-empty unique visible selection for the new window.

**Schema returns:**

- `window` (WorkbookWindow) — Appended workbook window. Source-free XLSX export authors a matching workbookView and one sheetView per worksheet.

#### `workbook.worksheets.add`

Append an editable visible, hidden, or very-hidden worksheet with a stable name and ID.

**Schema parameters:**

- `name` (string) — Unique worksheet name; defaults to SheetN.
- `visibility` (string) — visible (default), hidden, or veryHidden.

**Schema returns:**

- `worksheet` (Worksheet) — Appended editable worksheet with bounded native visibility.

#### `workbook.worksheets.getSelectedWorksheets`

Return the visible worksheet-tab group selected in the primary workbook window, in workbook order.

**Schema returns:**

- `worksheets` (Worksheet[]) — Selected visible worksheet tabs in workbook order, always including the active worksheet.

#### `workbook.worksheets.setActiveWorksheet`

Select the visible worksheet opened by default and used by workbook operations that omit an explicit sheet.

**Schema parameters:**

- `worksheet` (string|number|Worksheet) required — Visible worksheet name, zero-based collection index, or worksheet object from this workbook.

**Schema returns:**

- `worksheet` (Worksheet) — Selected visible worksheet. XLSX export writes its zero-based position to workbookView activeTab and collapses the primary tab selection to that worksheet.

#### `workbook.worksheets.setSelectedWorksheets`

Select one or more visible worksheet tabs in the primary workbook window while retaining exactly one active worksheet.

**Schema parameters:**

- `worksheets` (Array<string|number|Worksheet>) required — Non-empty unique list of visible worksheet names, zero-based indexes, or worksheet objects. If the current active worksheet is omitted, the first requested worksheet becomes active.

**Schema returns:**

- `worksheets` (Worksheet[]) — Selected worksheet tabs in workbook order; native XLSX export writes sheetView tabSelected for workbookViewId 0.

#### `workbook.xlsxFormulaSyntax`

Write formulas with the names and spill syntax shown in Excel, such as STDEV.S(A1:A10), FILTER(A1:A10,A1:A10>0), and SUM(E1#). OfficeKit Codec maps modeled future functions plus A1# to their required _xlfn/_xlws/ANCHORARRAY XLSX storage forms, returns public formulas without those package prefixes, and preserves an unchanged imported cell formula's original storage spelling.

**Examples:**

- =STDEV.S(A1:A10)
- =FILTER(A1:A10,A1:A10>0)
- =SUM(E1#)

**Schema parameters:**

- `formula` (string) required — Excel-visible formula such as =STDEV.S(A1:A10), =FILTER(A1:A10,A1:A10>0), or =SUM(E1#).

**Schema returns:**

- `formula` (string) — The public formula string. OfficeKit Codec maps the bounded modeled future-function and spill syntax to XLSX package spelling at export, removes that package spelling at import, and retains an unchanged imported cell formula's original storage form.

**Notes:**

- Author model formulas with Excel-visible names and A1# spill references. OfficeKit Codec owns the bounded XLSX package spelling (_xlfn, _xlws, and ANCHORARRAY), strips it on import, and preserves an unchanged imported cell formula's original package spelling.

#### `workbookWindow.getActiveWorksheet`

Return the visible active worksheet for one workbook window.

**Schema returns:**

- `worksheet` (Worksheet) — Visible active worksheet for this window.

#### `workbookWindow.getSelectedWorksheets`

Return one window's visible selected worksheet tabs in workbook order.

**Schema returns:**

- `worksheets` (Worksheet[]) — Visible selected worksheet tabs for this window in workbook order, always including its active worksheet.

#### `workbookWindow.setActiveWorksheet`

Set one window's active worksheet and collapse that window's selected tab group to it.

**Schema parameters:**

- `worksheet` (string|number|Worksheet) required — Visible worksheet resolved within the owning workbook.

**Schema returns:**

- `worksheet` (Worksheet) — Selected worksheet; the window's selected group is collapsed to this worksheet.

#### `workbookWindow.setSelectedWorksheets`

Set one window's non-empty visible selected tab group, which must include its active worksheet.

**Schema parameters:**

- `worksheets` (Array<string|number|Worksheet>) required — Non-empty unique visible selection. If the current active worksheet is omitted, the first requested worksheet becomes active.

**Schema returns:**

- `worksheets` (Worksheet[]) — Selected worksheet tabs for this window in workbook order.

#### `worksheet.freezePanes.freezeColumns`

Freeze a leading column count in the worksheet view while preserving any frozen rows.

**Schema parameters:**

- `columnCount` (number) required — Integer number of leading columns to freeze; zero clears only the column freeze.

**Schema returns:**

- `freezePanes` (object) — Worksheet frozen-pane facade with rows, columns, topLeftCell, activePane, and frozen state.

#### `worksheet.freezePanes.freezeRows`

Freeze a leading row count in the worksheet view while preserving any frozen columns.

**Schema parameters:**

- `rowCount` (number) required — Integer number of leading rows to freeze; zero clears only the row freeze.

**Schema returns:**

- `freezePanes` (object) — Worksheet frozen-pane facade with rows, columns, topLeftCell, activePane, and frozen state.

#### `worksheet.freezePanes.unfreeze`

Remove all frozen worksheet panes and restore a single scrollable view.

**Schema returns:**

- `freezePanes` (object) — Worksheet frozen-pane facade reset to zero frozen rows and columns.

#### `worksheet.getRange`

Select an A1 range for values, formulas, formatting, merge, fill, and copy operations.

**Schema parameters:**

- `address` (string) required — A1 cell or range address such as A1:D10.

**Schema returns:**

- `range` (Range) — Editable range facade for values, formulas, formatting, and rules.

#### `worksheet.getUsedRange`

Return the worksheet used rectangle, optionally excluding formatting-only cells with valuesOnly=true.

**Schema parameters:**

- `valuesOnly` (boolean) — When true, exclude cells represented only by formatting or other non-value state.

**Schema returns:**

- `range` (Range) — Used worksheet rectangle, or A1 for an empty worksheet.

#### `worksheet.mergeCells`

Merge an A1 range as one region or merge each row separately with across=true, retaining only upper-left content.

**Schema parameters:**

- `address` (string|Range) required — A1 range to merge.
- `across` (boolean) — Merge each row as a separate region instead of one rectangular region.

**Schema returns:**

- `worksheet` (Worksheet) — The same worksheet with native merged-range state.

#### `worksheet.protection`

Author, inspect, edit, or remove one passwordless worksheet editing restriction with an intuitive allowed-operation list. Cell locked/hidden styles become effective only while protection is active. This is not encryption or access control; password/hash variants remain source-owned and fail closed on replacement.

**Schema parameters:**

- `enabled` (boolean) — Protection is active when present. Assign null, false, or { enabled: false } to remove a recognized passwordless restriction.
- `allow` (string[]) — Allowed operations: selectLockedCells, selectUnlockedCells, formatCells, formatColumns, formatRows, insertColumns, insertRows, insertHyperlinks, deleteColumns, deleteRows, sort, autoFilter, pivotTables, editObjects, or editScenarios. Omission allows selection of locked and unlocked cells only.

**Schema returns:**

- `protection` (object|undefined) — Passwordless worksheet editing restriction. OfficeKit contains SpreadsheetML's inverted lock flags and source binding; password/hash/extension profiles are preserved opaquely and semantic replacement fails closed. This is not encryption, authentication, or access control.

#### `worksheet.sortState`

Get or set bounded worksheet-level row/column sorting; columnSort=true uses unique single-row conditions across the sort range.

**Schema parameters:**

- `reference` (string) required — Whole worksheet range whose rows or columns are sorted.
- `caseSensitive` (boolean) — Whether text comparisons are case-sensitive.
- `sortMethod` ('none'|'pinYin'|'stroke') — Optional locale-specific SpreadsheetML method; omission remains distinct from explicit 'none'.
- `columnSort` (boolean) — Optional presence-aware direction. true sorts columns left-to-right; false explicitly selects ordinary row sorting.
- `conditions` (object[]) required — Ordered unique single rows when columnSort=true, otherwise unique single columns; value conditions may add customList and icon/color selectors reuse the table-sort shape.

**Schema returns:**

- `sortState` (object) — Bounded worksheet-level sort state. QueryTable refresh sorts may be inspected after import but remain immutable; the only QueryTable edit is root automatic-refresh hardening through table.setQueryRefreshPolicy.

#### `worksheet.unmergeCells`

Remove every merged region intersecting an A1 range without discarding the retained upper-left content.

**Schema parameters:**

- `address` (string|Range) required — A1 range whose intersecting merged regions should be removed.

**Schema returns:**

- `worksheet` (Worksheet) — The same worksheet after intersecting merges are removed.

#### `worksheet.visibility`

Read or assign native worksheet visibility as visible, hidden, or veryHidden; at least one sheet must remain visible.

**Schema parameters:**

- `visibility` (string) required — visible, hidden, or veryHidden.

**Schema returns:**

- `visibility` (string) — Normalized worksheet visibility; workbook verification/export rejects an all-hidden workbook.

#### `worksheetChart.accessibilityCapability`

Report sourceBound/editable/addable preflight for a worksheet chart graphic-frame xdr:cNvPr title/description/decorative leaf independently of ChartSpace editability.

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, editable, addable } preflight for the chart frame xdr:cNvPr leaf; ChartSpace has an independent capability.

#### `worksheetChart.setAccessibilityMetadata`

Transactionally add, change, or clear a worksheet chart's non-visible title/description/decorative metadata without changing its visible chart title. Ambiguous imported extension graphs fail closed.

**Schema parameters:**

- `update` (object) required — Partial { title?, description?, decorative? }; null clears a field, text is 1-1,024 XML-safe characters, decorative is boolean, and decorative true excludes text.

**Schema returns:**

- `chart` (WorksheetChart) — The same chart after one transactional metadata update. Source-free objects are editable; imported objects require accessibilityCapability.editable and export re-proves the residual graph.

#### `worksheetImage.accessibilityCapability`

Report sourceBound/editable/addable preflight for worksheet picture xdr:cNvPr title/description/decorative metadata.

**Schema returns:**

- `capability` (object) — Fresh { sourceBound, editable, addable } preflight for picture xdr:cNvPr metadata.

#### `worksheetImage.setAccessibilityMetadata`

Transactionally add, change, or clear worksheet picture title/description/decorative metadata. image.alt is the same description state and is never inferred from the object or file name.

**Schema parameters:**

- `update` (object) required — Partial { title?, description?, decorative? }; null clears a field, text is 1-1,024 XML-safe characters, decorative is boolean, and decorative true excludes text.

**Schema returns:**

- `image` (WorksheetImage) — The same image after one transactional metadata update. The legacy alt property is the description alias; imported ambiguous xdr:cNvPr graphs fail closed without disabling unrelated picture edits.

