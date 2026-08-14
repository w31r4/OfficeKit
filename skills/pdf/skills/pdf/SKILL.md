---
name: "pdf"
description: "Create, inspect, edit, redact, sign, render, and verify PDF files through explicit, fail-closed provider routes. Use for new tagged documents, imported PDFs, forms, annotations, OCR, conformance, and visual QA."
---

# PDF Skill
Use `../office-kit/references/workspace.md` for `inputRoot`, `taskRoot`, `outputRoot`, `evidenceRoot`, absolute PDF paths, SHA-256, locators, and visual-review status.
For multi-step PDF work use `officekit repl` and `../office-kit/references/repl.md`; import through `ctx.import`, keep summaries in `ctx.state`, publish with `ctx.publish`, register QA with `ctx.recordEvidence`, and inspect before retrying `maybeApplied` mutations. Provider setup, OCR, signing, repair, and conformance remain explicit commands.
After the final PDF is reopened, follow `../office-kit/references/review.md`.
The optional AnyDoc text reading view (`contentView: "anydoc"`) is bounded, not OCR, signature/redaction evidence, PDF/UA validation, or visual QA.
Use the public `office-kit` package for OfficeKit PDF work. Never import or use `@oai/artifact-tool`: it is a different host runtime, not an OfficeKit alias or fallback, and its output must never be attributed to OfficeKit.
This Skill gives an agent bounded, auditable PDF primitives. PDF is independent
from the OfficeKit DOCX/XLSX/PPTX codec: do not add a PDF protobuf/WASM codec or
reconstruct an imported PDF through `PdfArtifact` or PDF.js and call that a
fidelity-preserving edit.

`PdfArtifact` is for new semantic/tagged documents. `PdfFile` plus required,
runtime-lazy MuPDF.js is the default for arbitrary PDF inspect, render, and bounded direct-original edits. Every specialist route is explicit. A failed
provider is an error, never a silent fallback.

Use `createPdfjsParser()` only as an optional read-only PDF.js adapter for page geometry, positioned text, heuristic tables, and bounded image evidence; it is never an imported-PDF edit representation.

## Route every job first

Use this sequence for every request:

1. Preserve original bytes and inspect with MuPDF.js.
2. Resolve the declared intent with `PdfProviders.resolve`.
3. If the result is `installable`, obtain explicit project-policy authority and
   call `ensure`; otherwise use only the returned `ready` provider.
4. Probe the selected provider, perform one explicit save strategy, then
   inspect, audit, render, and review the output.

```js
import { PdfFile } from "office-kit";
import { PdfProviders } from "office-kit/pdf/providers";

const inspection = await PdfFile.inspectPdf("input.pdf");
let resolution = await PdfProviders.resolve({
  task: "edit-content",
  inspection,
  savePolicy: "rewrite",
  mutationAuthorized: true,
});
if (resolution.status === "installable") {
  resolution = await PdfProviders.ensure({ resolution });
}
if (resolution.status !== "ready") throw new Error(resolution.reason.message);
```

The public resolver returns only `ready`, `installable`, or `blocked`, with the
selected provider/pack, platform, pinned version/artifact information, download
and unpack estimates, licence acknowledgement, runtime paths, prerequisites, and operation boundary. It never selects an alternate provider.
`PdfProviders.ensure({ resolution })` never acquires P12/private keys,
HSM/remote-signing credentials, TSA/LTV access, or other secrets.
`PdfProviders.probe({ provider })` checks exactly that provider without a
download or fallback.

Run the default native CLI before an imported-PDF operation:

```bash
officekit run scripts/mupdf.mjs probe
officekit run scripts/mupdf.mjs inspect input.pdf
officekit run scripts/mupdf.mjs render input.pdf tmp/pdfs/page-1.png --page 1 --dpi 144
```

The CLI budgets work, refuses overwrite/symlink aliases, writes atomically, and
never falls back. Source-bound outline edits are in [edit existing](tasks/edit_existing.md).

## Capability-pack policy

The only required npm runtime is `mupdf@1.28.0`; it initializes only on the
first MuPDF-backed PDF operation. Root import and provider resolution do not
start WASM, download a pack, or modify the filesystem.

External capability packs live in the project-private
`.office-kit/providers/` cache. The conventional policy is
`.office-kit/pdf-providers.json`; a missing file means
`installPolicy: "disabled"`. An agent may install only after the user/project
sets `managed`, whitelists every provider and pack, accepts required licences,
allows the requested OCR languages, and supplies finite byte budgets:

```json
{
  "installPolicy": "managed",
  "allowedProviders": ["qpdf"],
  "allowedPacks": ["qpdf"],
  "acceptedLicenses": [],
  "allowedOcrLanguages": ["eng", "chi_sim"],
  "maxDownloadBytes": 250000000,
  "maxUnpackedBytes": 750000000
}
```

Only hash-pinned, versioned project release assets may be installed; an
enterprise mirror must serve identical bytes. The installer uses a lock,
temporary download, exact size/hash checks, safe archive extraction, atomic
publication, and a receipt. It rejects `latest`, package-manager/global-pip
installs, lifecycle hooks, path traversal, links, and undeclared URLs. The
managed targets are `darwin-arm64`, `linux-x64`, and `win32-x64`.

**Current release-catalog state:** qpdf `12.3.2-oat.2`, `python-foundation`/`python-specialists` `3.13.14-oat.2`, veraPDF/JRE `1.30.2-oat.2`, OCR core `17.8.1-oat.3`, and `eng`/`chi_sim` `4.1.0-oat.3` have published, attested `darwin-arm64`, `linux-x64`, and `win32-x64` assets; public live acceptance verifies those all-platform closures. Poppler QA `24.08.0-oat.2` is also published and attested for all three managed platforms; macOS/Linux use source-built relocatable closures with pinned Poppler data and declared native roots, while Windows retains its reviewed upstream closure.
qpdf covers repair, linearization, inspection, and the bounded AES-256 delivery-copy route. The foundation is isolated CPython with ReportLab,
pdfplumber, pypdf, and Pillow. Specialists contain PyMuPDF, pikepdf, pyHanko,
and certificate validation; they depend on qpdf and require the catalogued
AGPL-or-commercial acknowledgement. The veraPDF pack carries its managed JRE;
OCR installs qpdf, its isolated OCRmyPDF/Tesseract 5/Ghostscript/`pdftotext` core, and only explicit language packs. Poppler QA `24.08.0-oat.2` is published and attested for `darwin-arm64`, `linux-x64`, and `win32-x64`; a policy-authorized route on any catalogued platform resolves as `installable` and uses its pinned `pdfinfo`, `pdftoppm`, and `pdftotext` bytes. The dedicated live release check uses `resolve` → `ensure` → `probe`, native text extraction, and raster output against those exact bytes. Use a selected `system-only` policy with an explicitly configured local runtime when one is already managed by the deployment. Once a task chooses managed or system-only, do not switch routes.

See [provider setup and probes](tasks/provider_setup.md) for the full policy,
system-runtime, probe, and failure contract. The [provider matrix](references/PROVIDER_MATRIX.md)
states human-readable capability boundaries; the public catalog is the only
source of pack versions, hashes, sizes, and installation facts.
## Choose the narrowest provider

| Need | Explicit route | Detailed task |
| --- | --- | --- |
| New tagged/semantic PDF | `PdfArtifact` | [create](tasks/create.md) |
| New layout-oriented PDF | ReportLab | [create](tasks/create.md) |
| Imported-PDF inspect/render/bounded edit | MuPDF.js / `scripts/mupdf.mjs` | [read](tasks/read_review.md), [edit](tasks/edit_existing.md) |
| Text/table geometry evidence | pdfplumber | [read](tasks/read_review.md) |
| Attachments, complex forms, verified static-form delivery, merge/reorder/stamp | pypdf | [forms](tasks/forms_annotations.md), [transform](tasks/transform.md) |
| Strict scrub, residue/OCR redaction, advanced bounded edit | PyMuPDF | [redact](tasks/redact.md) |
| Repair or linearize | `scripts/qpdf_provider.py` | [repair](tasks/repair_linearize.md) |
| Create an AES-256 encrypted delivery copy | `scripts/qpdf_provider.py` | [encryption](tasks/encryption.md) |
| Active/auxiliary structure cleanup | `scripts/pikepdf_provider.py` | [structure cleanup](tasks/structure_clean.md) |
| Searchable-layer OCR | `scripts/ocrmypdf_provider.py` | [OCR](tasks/ocr.md) |
| Local PKCS#12 sign, validation, or controlled P=2 form finalisation | `scripts/pyhanko_sign_provider.py`, `scripts/pyhanko_provider.py`, `scripts/pyhanko_certified_form_fill.py` | [sign](tasks/sign_verify.md) |
| PDF/A or PDF/UA machine checks | `scripts/verapdf_provider.py` | [accessibility](tasks/accessibility.md) |
| Independent native visual QA | Poppler | [render review](tasks/render_review.md) |

## Imported-PDF invariants

Keep the source immutable, bind its SHA-256, choose `read-only`, `rewrite`,
`incremental`, or `sanitize` before mutation, and publish a distinct output.
Inspect signatures, ByteRange, DocMDP/FieldMDP, encryption, forms, annotations,
attachments, metadata, active content, page boxes, and page count first.
`incremental` preserves old bytes; it is not signature authorization.

MuPDF.js supports only bounded direct-original operations. Its inspect output
keeps raw `mediaBox`/`cropBox` as unrotated PDF-space facts and emits a rotated
effective `mupdf-page-space` bbox for 0/90/180/270-degree placement. Use its
returned `sourceSha256`, `mupdf-link`/annotation/form locators,
snapshots/capabilities, and `appearanceBbox`; re-inspect after every output.

- `add_text_annotation`, `add_free_text_annotation`, `add_text_markup` (Highlight/Underline/StrikeOut/Squiggly), compatibility `add_text_highlight`, `add_link`, `delete_annotation`,
  `update_annotation`, `delete_link`, `update_link`, and `update_form_field`
  are source-bound operations; placement uses the inspected page snapshot.
  `add_text_annotation` takes a visible pin; `add_free_text_annotation` takes a
  fixed-Helvetica visible text box and refuses an appearance that omits text; both require rewrite. Its recognized profile permits content/reviewer updates with a complete snapshot and renewed fit proof. Text markup requires a unique native text selection and rewrite.
- On untagged input, `delete_page`, `duplicate_page`, and changed `rearrange_pages` require exact source SHA-256 and page snapshots, run as the only operation in a full rewrite, then require re-inspect and mapped Poppler render/pixel identity.
- `delete_embedded_file` binds one inspect-returned canonical catalog NameTree locator and complete snapshot; rewrite removes that entry only and never claims sanitize or physical payload erasure.
- `set_metadata` binds the exact source plus the complete `mupdfDocumentMetadata` snapshot and updates standard Document Info fields. For bounded `field-safe-v1` XMP, inspection reports separate `xmpMutableFields` and `xmpBlockedFields`; the same transaction updates only requested proven fields and leaves all other packet bytes unchanged. A unique `x-default` remains editable in a multilingual title/description, common scalar description attributes are editable, and an unrelated multi-author or irregular field blocks only that field. Missing or blocked requested fields, CDATA/DTD/invalid entities, malformed XML, stale/no-op evidence, and unsupported stream graphs fail closed without switching providers.
- `set_page_crop` is raw unrotated CropBox visibility only, not redaction.
  `rotate_page` sets an absolute right-angle `/Rotate`; neither enables content
  reflow. Delete/redaction operations cannot be incremental. Mixed-document OCR needing automatic rotate or deskew is audit-only `failed_closed` with `savePolicy.strategy: "none"`; see [OCR](tasks/ocr.md).
- General Word-style reflow, arbitrary text replacement, Dynamic XFA, complex
  JavaScript, 3D, and RichMedia are not made safe by these primitives. Current
  typed routes cannot prove opaque closure or runtime behavior, so inventory
  such objects and fail closed with audit-only `save: none` unless a dedicated
  provider and oracle prove unchanged bytes plus runtime semantics. Never route
  around this boundary through pypdf, ReportLab, PDF.js, content-stream edits,
  or another provider; do not publish a cover-only edit as preserved behavior.

For the complete operation schemas and edge cases, read [edit existing](tasks/edit_existing.md) and [forms and annotations](tasks/forms_annotations.md), not this overview.

## Specialist safety boundaries

- qpdf receives a source SHA-256 before repair or linearize. Its separate
  `encrypt` primitive creates one AES-256 copy from an unencrypted source using
  caller-owned restricted password files and private argument files; it does
  not open/decrypt/re-encrypt existing encrypted PDFs or edit permissions. Both
  routes are structural/full-rewrite operations, not redaction or sanitize;
  use [repair](tasks/repair_linearize.md) or [encryption](tasks/encryption.md).
- pikepdf offers only `active-content` and `active-and-auxiliary` profiles. It
  is not redaction, metadata cleanup, or XFA cleanup; use
  [structure cleanup](tasks/structure_clean.md).
- PyMuPDF `redact_ocr_text` is sanitize-only: require exact
  `expected_rotation` (0/90/180/270), a named language, match count, bounded
  raster work, residue evidence, rewrite, and invalidation acknowledgement.
  Coordinates remain unrotated PyMuPDF page space. A complete imported PDF
  OCR workflow is not a sanitizer; see [OCR](tasks/ocr.md).
- pyHanko local PKCS#12 signing uses a passphrase on stdin only. Its P=2 route
  finalises one verified empty decimal field under an exact FieldMDP lock set;
  `pyhanko_provider.py` validates under an explicit trust root. Timestamp,
  LTV, PKCS#11/HSM, remote signing, and network evidence are external workflows.
- veraPDF's `verapdf_provider.py` is a machine-rule gate, not repair or a substitute for
  human review of PDF/UA. See [accessibility](tasks/accessibility.md).

When a Python specialist is selected, one configured virtual environment
executable remains provider identity so its `pyvenv.cfg` is preserved. Do not
retry via a different Python interpreter after a failed probe.

## Delivery gate

1. Record selected provider, version, policy/receipt (if any), source hash, save strategy, and no-fallback evidence in the canonical audit envelope. A no-mutation refusal uses `pdf_audit.py failed-closed`; a verified DocMDP P=1 refusal additionally supplies its pyHanko report with `--signature-verification --require-docmdp-no-changes --trust-root`. Do not hand-write either audit.
   A refused PAdES/TSA/LTV upgrade must pass `--capabilities-json` with explicit boolean `false` values for `timestampAuthoritySupported`, `ltvEmbeddingSupported`, and `padesProfileConformanceClaimed`; prose or missing fields are not evidence.
2. Reopen and verify the intended change; use `scripts/pdf_audit.py validate`.
3. Run the requested specialist evidence. Sanitization must pass its residue
   and single-revision gates; signatures and PDF/UA retain their separate
   validation/human boundaries.
4. Render every final page with MuPDF.js or independently with Poppler and
   inspect clipping, overlaps, glyphs, images, fields, annotations, signatures,
   redactions, and page geometry when visual input is available. Without it, retain
   renders, run structural/page-geometry checks, and mark visually material results
   for human review. Request AnyDoc only to close a text/table coverage gap; it is not OCR.

The project and its required MuPDF.js dependency are GNU AGPL-3.0-or-later.
Managed and system providers retain their own licences; the resolver exposes
the acknowledgement required before any installation or operation.
