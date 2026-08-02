# Read and review an existing PDF

Read-only review never needs model re-export.

## Route by evidence need

- `pdfinfo`: page count, page size, encryption, metadata summary.
- `pdftoppm`: final visual truth for every page.
- pdfplumber: text, words with geometry, tables, lines, and rectangles.
- pypdf: metadata, fields, annotations, outlines, attachments, encryption, and object-level quick checks.
- MuPDF.js through `scripts/mupdf.mjs inspect`: default native page/object/text/image/link/annotation/widget evidence.
- PyMuPDF: optional specialist inspection when its separate workflow is selected.
- qpdf through `scripts/qpdf_provider.py inspect`: bounded xref/object-stream
  structure, warnings/recovery evidence, encryption/linearization facts, and
  signature-policy indicators bound to the source SHA-256.
- pyHanko: signature and trust validation.
- veraPDF: PDF/A or PDF/UA machine-verifiable rules.

## Extraction

```bash
python3 scripts/pdfplumber_extract.py input.pdf \
  --output tmp/pdfs/extraction.json \
  --max-pages 200
```

Extraction is not layout fidelity. Compare extracted text/table candidates against rendered pages, especially multi-column layouts, rotated text, merged cells, OCR layers, and scanned pages.

For a bounded ruled-table task, use the typed `table` primitive instead of
turning every extracted word into a cell:

```bash
python3 scripts/pdf_provider.py check --provider pdfplumber --require
python3 scripts/pdf_provider.py plan --task extract --provider pdfplumber \
  --strategy read-only --input inputs/source.pdf --require-provider
python3 scripts/pdfplumber_extract.py table inputs/source.pdf \
  outputs/regional-revenue.json outputs/regional-revenue.csv \
  --table-name "Regional Revenue"
```

The primitive requires exactly one `pdfplumber` table per page, preserves
merged title geometry (`colspan`), emits one confidence value per cell, and
keeps explicit `*`/`Note:` lines outside the ruled table in a separate
`footnotes` array. Footnotes are never mixed into `cells` or the CSV. It is a
bounded geometry primitive, not a general table-understanding or reading-order
algorithm; inspect the result against Poppler renders and refuse ambiguous
topology rather than guessing.

Its JSON contains `table`, `cells`, `pages`, `footnotes`, source/provider
evidence, and read-only strategy. The CSV contains the stable header
`page,text,bbox,rowspan,colspan,confidence`. Record the canonical audit with
`operation.type: "extract-table"`, mirror the JSON file in `output`, and bind
both files under named `outputs.json`/`outputs.csv` evidence. Validate all bytes:

```bash
python3 scripts/pdf_audit.py validate outputs/audit.json \
  --source inputs/source.pdf \
  --artifact-json outputs/regional-revenue.json \
  --artifact-csv outputs/regional-revenue.csv \
  --require-operation extract-table
```

For this table workflow, `validation.renderReview` must record the Poppler
renderer/version, the three rendered pages, and the overlay decision. Either
set `bboxOverlayReviewed: true` at the render-review level, or set
`tableBboxOverlayReviewed: true` and `result: "passed"` on every page entry;
the latter is the preferred auditable form. Keep the rendered PNGs and
overlays under `tmp/` unless the task explicitly declares them as deliverables.
The installed package under `node_modules/` and the mounted `.agents/` Skill
tree are read-only inputs: never patch them to work around an extraction
failure. Stop and report the provider error instead of changing the runtime.

Narrative columns, rotated labels, low-evidence OCR cells, and table-like text
outside the ruled geometry stay out of `cells`; report them as warnings or
separate evidence instead of fabricating coordinates.

For qpdf structural evidence:

```bash
PYTHON_BIN="${OFFICE_KIT_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/qpdf_provider.py probe
"$PYTHON_BIN" scripts/qpdf_provider.py inspect input.pdf \
  > tmp/pdfs/qpdf-inspect.json
```

qpdf warning exit status is evidence, not automatic permission to rewrite. Use
the returned source hash and the separate [repair/linearize workflow](repair_linearize.md)
only after reviewing encryption and signature constraints.

`PdfFile.importPdf(...)` uses MuPDF.js by default for agent-facing extraction/QA of an arbitrary PDF; the result remains a reconstructed view. Never export that model as an edit to the original file. Use `PdfFile.editPdf(...)` or `scripts/mupdf.mjs edit` on the original bytes for supported mutations. PDF.js is an optional explicitly injected independent parser.

## Attachment quarantine

Never write an embedded filename directly to disk. A FileSpec may contain `../`, absolute paths, platform separators, control characters, reserved device names, or duplicate names. Inventory and extract through the shipped read-only primitive:

```bash
PYTHON_BIN="${OFFICE_KIT_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pypdf_edit.py inspect input.pdf \
  --output tmp/pdfs/pypdf-inspect.json
"$PYTHON_BIN" scripts/pdf_provider.py check --provider pypdf --require
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task extract-attachments --provider pypdf --strategy read-only \
  --input input.pdf --require-provider
"$PYTHON_BIN" scripts/pypdf_edit.py extract-attachments input.pdf outputs/quarantine \
  --manifest outputs/attachments.json \
  --max-attachments 1000 \
  --max-total-bytes 1073741824 \
  --max-attachment-bytes 536870912
```

The manifest records provider/version, immutable source hash, scope (`document` or `page`), page/annotation identity, display name, internal key, MIME and its evidence source, decoded byte size, SHA-256, sanitized saved name/path, and transaction validation. Duplicate or colliding names receive deterministic suffixes and remain separate. A malformed FileSpec, unreadable stream, exceeded budget, pre-existing destination, hash mismatch, or source change fails closed and removes partial quarantine output. The primitive never opens, executes, imports, or recursively extracts an attachment.

Create the canonical operation audit with `savePolicy.strategy: "read-only"`, `operation.type: "extract-attachments"`, and `output` bound to the exact `attachments.json` bytes. Validate it with `pdf_audit.py validate --source input.pdf --artifact outputs/attachments.json --require-operation extract-attachments`.

## Review output

Report confirmed facts separately from inferences:

- file structure and metadata;
- visible page content and layout;
- extracted text/table candidates;
- forms, annotations, attachments, and signatures;
- accessibility/conformance evidence;
- warnings, unsupported structures, and missing providers.
