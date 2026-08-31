# API catalog

Generated from `PUBLIC_HELP_CATALOG` in `src/help/index.mjs`.

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
| `officekit ppj build` | cli | Compile an authored .ppj to editable PPTX or lower a source-bound PPJ diff into a capability-proven local Edit Plan. A third-party no-op returns the exact source bytes; unsupported mutations fail closed. |
| `officekit ppj check` | cli | Validate PPJ schema, stable references, local assets, bounded component expansion, source hashes, and nativeRef capabilities before compilation. --fix is limited to deterministic formatting repairs. |
| `officekit ppj import` | cli | Project a PPTX into one strict JSON .ppj program. OfficeKit-authored PPTX files recover their embedded program when its map still matches; third-party files produce typed elements plus source-bound opaque nativeRef records without putting unknown OOXML into the program. |
| `officekit ppj inspect` | cli | Search or inspect a .ppj program by stable page and element IDs without evaluating code or changing the file. |
| `officekit ppj render` | cli | Compile and render selected PPJ pages for visual review without treating a successful render as design approval. |
| `officekit ppj resume` | cli | Materialize the latest valid immutable PPJ revision and all bound local resources from a durable OfficeKit task into a new editable workspace without restoring a JavaScript heap or modifying the task store. |
| `officekit ppj review` | cli | Review PPJ structure, layout, source fidelity, communication intent, motion, delivery evidence, and rendered pages. It reports visual capability honestly and never invents fact verification. |

### presentation details

#### `officekit ppj build`

Compile an authored .ppj to editable PPTX or lower a source-bound PPJ diff into a capability-proven local Edit Plan. A third-party no-op returns the exact source bytes; unsupported mutations fail closed.

**Adoption tier:** `golden`

**Use when:**

- The agent needs to create, inspect, validate, compile, render, or review the durable Presentation program.
- The operation must remain deterministic and resumable without a JavaScript heap.

**Avoid when:**

- Do not use raw OOXML, XPath, relationship IDs, or the internal Presentation object model as a substitute.
- Do not claim visual, playback, or factual review beyond the evidence returned by the command.

**Requires:**

- UTF-8 strict JSON using office-kit/ppj/v1
- local content-addressed assets and exact source hashes when referenced

**Review:**

- Run officekit ppj check before delivery.
- Render and review the pages affected by the current change; re-import the built PPTX when source fidelity matters.

**Recipes:**

- skills/presentations/skills/presentations/SKILL.md#routes

**Example paths:**

- skills/presentations/skills/presentations/references/ppj.md

**Examples:**

- officekit ppj build deck.ppj -o deck.pptx --json

**Schema parameters:**

- `input` (path) required — Checked .ppj program.
- `output` (path) required — New PPTX output; source inputs are never overwritten.
- `task` (string) — Optional task id used to persist the build revision.
- `json` (boolean) — Emit a structured build receipt.

**Schema returns:**

- `receipt` (object) — Output SHA-256, compile mode, mutation footprint, source-preservation evidence, and warnings.

#### `officekit ppj check`

Validate PPJ schema, stable references, local assets, bounded component expansion, source hashes, and nativeRef capabilities before compilation. --fix is limited to deterministic formatting repairs.

**Adoption tier:** `golden`

**Use when:**

- The agent needs to create, inspect, validate, compile, render, or review the durable Presentation program.
- The operation must remain deterministic and resumable without a JavaScript heap.

**Avoid when:**

- Do not use raw OOXML, XPath, relationship IDs, or the internal Presentation object model as a substitute.
- Do not claim visual, playback, or factual review beyond the evidence returned by the command.

**Requires:**

- UTF-8 strict JSON using office-kit/ppj/v1
- local content-addressed assets and exact source hashes when referenced

**Review:**

- Run officekit ppj check before delivery.
- Render and review the pages affected by the current change; re-import the built PPTX when source fidelity matters.

**Recipes:**

- skills/presentations/skills/presentations/SKILL.md#routes

**Example paths:**

- skills/presentations/skills/presentations/references/ppj.md

**Examples:**

- officekit ppj check deck.ppj --json

**Schema parameters:**

- `input` (path) required — Existing .ppj program.
- `fix` (boolean) — Apply deterministic non-semantic normalization only.
- `task` (string) — Optional task id used to persist the checked revision.
- `json` (boolean) — Emit structured diagnostics.

**Schema returns:**

- `report` (object) — Validation status, program hash, expanded budgets, warnings, and precise errors.

#### `officekit ppj import`

Project a PPTX into one strict JSON .ppj program. OfficeKit-authored PPTX files recover their embedded program when its map still matches; third-party files produce typed elements plus source-bound opaque nativeRef records without putting unknown OOXML into the program.

**Adoption tier:** `golden`

**Use when:**

- The agent needs to create, inspect, validate, compile, render, or review the durable Presentation program.
- The operation must remain deterministic and resumable without a JavaScript heap.

**Avoid when:**

- Do not use raw OOXML, XPath, relationship IDs, or the internal Presentation object model as a substitute.
- Do not claim visual, playback, or factual review beyond the evidence returned by the command.

**Requires:**

- UTF-8 strict JSON using office-kit/ppj/v1
- local content-addressed assets and exact source hashes when referenced

**Review:**

- Run officekit ppj check before delivery.
- Render and review the pages affected by the current change; re-import the built PPTX when source fidelity matters.

**Recipes:**

- skills/presentations/skills/presentations/SKILL.md#routes

**Example paths:**

- skills/presentations/skills/presentations/references/ppj.md

**Examples:**

- officekit ppj import input.pptx -o deck.ppj --json

**Schema parameters:**

- `input` (path) required — Existing PPTX input; never overwritten.
- `output` (path) required — New .ppj path.
- `task` (string) — Optional OfficeKit task id used to persist a revision.
- `json` (boolean) — Emit a machine-readable receipt.

**Schema returns:**

- `receipt` (object) — Projection mode, stable program identity, hashes, warnings, and output path.

#### `officekit ppj inspect`

Search or inspect a .ppj program by stable page and element IDs without evaluating code or changing the file.

**Adoption tier:** `golden`

**Use when:**

- The agent needs to create, inspect, validate, compile, render, or review the durable Presentation program.
- The operation must remain deterministic and resumable without a JavaScript heap.

**Avoid when:**

- Do not use raw OOXML, XPath, relationship IDs, or the internal Presentation object model as a substitute.
- Do not claim visual, playback, or factual review beyond the evidence returned by the command.

**Requires:**

- UTF-8 strict JSON using office-kit/ppj/v1
- local content-addressed assets and exact source hashes when referenced

**Review:**

- Run officekit ppj check before delivery.
- Render and review the pages affected by the current change; re-import the built PPTX when source fidelity matters.

**Recipes:**

- skills/presentations/skills/presentations/SKILL.md#routes

**Example paths:**

- skills/presentations/skills/presentations/references/ppj.md

**Examples:**

- officekit ppj inspect deck.ppj --query revenue --json

**Schema parameters:**

- `input` (path) required — Existing .ppj program.
- `query` (string) — Bounded fuzzy text or ID query.
- `page` (string) — Optional stable page ID.
- `json` (boolean) — Emit structured results.

**Schema returns:**

- `records` (object[]) — Stable-ID matches, element summaries, source capabilities, and diagnostics.

#### `officekit ppj render`

Compile and render selected PPJ pages for visual review without treating a successful render as design approval.

**Adoption tier:** `golden`

**Use when:**

- The agent needs to create, inspect, validate, compile, render, or review the durable Presentation program.
- The operation must remain deterministic and resumable without a JavaScript heap.

**Avoid when:**

- Do not use raw OOXML, XPath, relationship IDs, or the internal Presentation object model as a substitute.
- Do not claim visual, playback, or factual review beyond the evidence returned by the command.

**Requires:**

- UTF-8 strict JSON using office-kit/ppj/v1
- local content-addressed assets and exact source hashes when referenced

**Review:**

- Run officekit ppj check before delivery.
- Render and review the pages affected by the current change; re-import the built PPTX when source fidelity matters.

**Recipes:**

- skills/presentations/skills/presentations/SKILL.md#routes

**Example paths:**

- skills/presentations/skills/presentations/references/ppj.md

**Examples:**

- officekit ppj render deck.ppj -o previews/ --pages 1-4 --json

**Schema parameters:**

- `input` (path) required — Existing .ppj program.
- `output` (path) required — Preview directory.
- `pages` (string) — Optional bounded page selector.
- `json` (boolean) — Emit render evidence.

**Schema returns:**

- `receipt` (object) — Rendered page paths, renderer identity, hashes, and unavailable boundaries.

#### `officekit ppj resume`

Materialize the latest valid immutable PPJ revision and all bound local resources from a durable OfficeKit task into a new editable workspace without restoring a JavaScript heap or modifying the task store.

**Adoption tier:** `golden`

**Use when:**

- The agent needs to create, inspect, validate, compile, render, or review the durable Presentation program.
- The operation must remain deterministic and resumable without a JavaScript heap.

**Avoid when:**

- Do not use raw OOXML, XPath, relationship IDs, or the internal Presentation object model as a substitute.
- Do not claim visual, playback, or factual review beyond the evidence returned by the command.

**Requires:**

- UTF-8 strict JSON using office-kit/ppj/v1
- local content-addressed assets and exact source hashes when referenced

**Review:**

- Run officekit ppj check before delivery.
- Render and review the pages affected by the current change; re-import the built PPTX when source fidelity matters.

**Recipes:**

- skills/presentations/skills/presentations/SKILL.md#routes

**Example paths:**

- skills/presentations/skills/presentations/references/ppj.md

**Examples:**

- officekit ppj resume t_0123456789ab -o resumed/deck.ppj --json

**Schema parameters:**

- `task` (string) required — Existing OfficeKit task id containing a valid PPJ revision.
- `output` (path) required — New editable .ppj path outside the immutable task store.
- `json` (boolean) — Emit the program, source, candidate, and review descriptor.

**Schema returns:**

- `receipt` (object) — Materialized PPJ path, program/source hashes, revision status, copied resources, candidate hash, and review status.

#### `officekit ppj review`

Review PPJ structure, layout, source fidelity, communication intent, motion, delivery evidence, and rendered pages. It reports visual capability honestly and never invents fact verification.

**Adoption tier:** `golden`

**Use when:**

- The agent needs to create, inspect, validate, compile, render, or review the durable Presentation program.
- The operation must remain deterministic and resumable without a JavaScript heap.

**Avoid when:**

- Do not use raw OOXML, XPath, relationship IDs, or the internal Presentation object model as a substitute.
- Do not claim visual, playback, or factual review beyond the evidence returned by the command.

**Requires:**

- UTF-8 strict JSON using office-kit/ppj/v1
- local content-addressed assets and exact source hashes when referenced

**Review:**

- Run officekit ppj check before delivery.
- Render and review the pages affected by the current change; re-import the built PPTX when source fidelity matters.

**Recipes:**

- skills/presentations/skills/presentations/SKILL.md#routes

**Example paths:**

- skills/presentations/skills/presentations/references/ppj.md

**Examples:**

- officekit ppj review deck.ppj --task defense-deck --json

**Schema parameters:**

- `input` (path) required — Existing .ppj program.
- `task` (string) — Optional task id used to bind review evidence to a revision.
- `json` (boolean) — Emit the complete review report.

**Schema returns:**

- `report` (object) — Blocking issues, warnings, review scope, playback/visual evidence, and program/source hashes.

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

