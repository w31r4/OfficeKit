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

## Verified ruled cross-page table profile

`pdfplumber_extract.py` remains a candidate-evidence route. It must not be
presented as a proven table when a page has multiple columns, merged cells, or
rotated text.

For one narrow, read-only profile, use the published workflow instead:

```bash
PYTHON_BIN="${OFFICE_KIT_PDF_PROVIDER_PYTHON:-python3}"
OFFICE_KIT_PDF_PROVIDER_PYTHON="$PYTHON_BIN" \
node examples/officekit-ruled-cross-page-table-workflow.mjs input.pdf \
  --table-title "Regional Revenue" \
  --expected-columns 4 \
  --header-rows 2 \
  --min-pages 3 \
  --footnote-prefix "*" \
  --json outputs/regional-revenue.json \
  --csv outputs/regional-revenue.csv \
  --audit outputs/audit.json \
  --render-dir tmp/pdfs/regional-revenue-review
```

It accepts only an explicitly titled, consecutive-page table with complete
ruled-grid coverage, fixed column boundaries, the same merged header geometry
on every segment, rectangular non-empty data rows, and an optional explicit
adjacent footnote prefix. It emits typed cells with `page`, `bbox`, `rowspan`,
`colspan`, and confidence evidence; CSV is flattened from the verified JSON.
It also renders every selected source page with Poppler and writes a table-bbox
overlay plus a byte-bound read-only audit.

There is no fallback from this profile to heuristic extraction. If the proof
fails, publish neither JSON nor CSV; use generic candidate evidence and request
human review instead of guessing cells or silently mixing nearby narrative.

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
