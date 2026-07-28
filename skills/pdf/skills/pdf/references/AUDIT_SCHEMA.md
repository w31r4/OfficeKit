# PDF operation audit schema

Every imported-PDF mutation and security-sensitive read-only extraction emits one canonical `office-kit.pdf-audit.v1` JSON record. Do not invent aliases such as `outputs.pdf`, `actual_provider`, `provider_version`, `save_strategy`, or `silent_fallback`; downstream Agent and evaluator code reads the stable camelCase fields below.

Required success shape:

```json
{
  "schema": "office-kit.pdf-audit.v1",
  "status": "succeeded",
  "source": { "path": "/absolute/input.pdf", "bytes": 123, "sha256": "..." },
  "output": { "path": "/absolute/output.pdf", "bytes": 123, "sha256": "..." },
  "provider": {
    "actual": "pymupdf",
    "version": "1.27.2.3",
    "licenseChoice": "agpl",
    "silentFallback": false
  },
  "savePolicy": { "strategy": "sanitize" },
  "preflight": { "probeCompleted": true, "planCompleted": true },
  "operation": { "type": "replace_text" },
  "validation": {}
}
```

`source.sha256` and `output.sha256` identify the exact delivered bytes. `provider.actual`, `provider.version`, `provider.silentFallback`, `savePolicy.strategy`, and `operation.type` are never inferred from prose. Provider-specific evidence, fit checks, signature policy, residue scans, Poppler results, warnings, and task-specific assertions belong in additional fields without renaming the canonical envelope.

When one delivery executes multiple typed steps, keep the umbrella
`operation.type` and add `operation.orderedOperations`: an array of the
provider's executed typed records in order. Preserve page, match, coordinate,
or cleanup facts from the provider report. This is execution evidence, not a
human-written summary; a generic `redact-and-sanitize` label alone is not an
auditable substitute for `redact_ocr_text`, `redact_text`, and `scrub` records.

For a multi-source operation, keep `source` as the exact operation manifest and add an `inputs` array containing the same file-evidence shape for every imported PDF. The validator requires a matching repeated `--input` path for every record and recomputes each byte count and SHA-256; ordering is not significant, but missing, duplicate, or extra records fail.

```json
{
  "source": { "path": "/absolute/merge-stamp.json", "bytes": 456, "sha256": "..." },
  "inputs": [
    { "path": "/absolute/cover.pdf", "bytes": 123, "sha256": "..." },
    { "path": "/absolute/report.pdf", "bytes": 234, "sha256": "..." }
  ]
}
```

For attachment quarantine, use `savePolicy.strategy: "read-only"`, `operation.type: "extract-attachments"`, and bind `output` to the delivered `attachments.json` manifest. The quarantine file hashes and contained paths remain task-specific validation evidence inside that manifest and the audit `validation` object.

For bounded qpdf encryption, use `savePolicy.strategy: "rewrite"` and
`operation.type: "qpdf-encrypt-aes-256"`. Record the catalog credential
declaration, qpdf version, AES-256/key-bit evidence, signature decision,
authorized `checkAfter`, topology comparison, and password-safe render review.
Never record either password, a password hash/fingerprint, a secret-file path,
private argument-file contents, or a command line containing a secret.

For `failed_closed`, set `output` to `null`, include a non-empty `reason`, keep the source/provider/save-policy/operation evidence, record `preflight.probeCompleted` and `preflight.planCompleted` truthfully (either may be `false` when that gate caused the refusal), and do not leave a partial modified PDF at the requested output path. A `succeeded` record requires both preflight fields to be `true`.

For an independently auditable no-mutation refusal, keep those facts typed rather
than relying on `reason` or warning prose: use
`operation.mutationAttempted: false`,
`validation.sourceIdentity.sourcePreserved: true`, and
`validation.artifactChecks.modifiedPdfPresent: false` plus
`partialArtifactPresent: false`. These fields bind the refusal to the source
and the actual output directory without implying that an unsupported operation
was performed.

For that common safe-refusal path, do not hand-write a near miss. After the
read-only provider/source inspection, use the shipped generator. It rejects a
non-empty delivery directory, writes only `audit.json` atomically, binds the
source hash, and emits the complete typed no-mutation/no-artifact envelope:

```bash
python3 scripts/pdf_audit.py failed-closed outputs/audit.json \
  --source inputs/source.pdf \
  --provider mupdf-js --provider-version 1.28.0 \
  --operation add-footer-work-order \
  --reason "professional print preflight is unavailable" \
  --probe-completed --plan-completed --source-inspected
```

Keep inspection logs under `tmp/`, not `outputs/`. The generator validates its
own envelope before publication; run `validate` as an independent final check
when the workflow requires it.

For a fully verified DocMDP P=1 certification refusal, the generic envelope is
not enough: bind the actual `office-kit.pyhanko-verify.v1` report instead of
copying its conclusions by hand. The failure-closed generator requires the
report, an exact explicit trust-root file, `pyhanko`, `read-only`, and completed
probe/plan/source inspection; it accepts only one intact, trusted, full-file
P=1 certification with all requested integrity, trust, DocMDP, and bottom-line
gates passing. It then writes `signaturePolicy` and
`validation.signatureVerification` with the exact root hash and no-mutation
decision. Use `--require-docmdp-no-changes --trust-root ...` again with
`validate` to recheck that evidence. P=2/P=3 changes are not blanket refusals:
they require a separately bounded field/policy workflow.

Validate before delivery:

```bash
python3 scripts/pdf_audit.py validate outputs/audit.json \
  --source inputs/source.pdf \
  --artifact outputs/modified.pdf \
  --require-operation replace_text
```

For a merge, use the manifest as `--source` and repeat `--input` for every source PDF.

The validator recomputes source and artifact byte counts and SHA-256 hashes. It does not replace semantic, signature, residue, conformance, or render verification; those results remain required entries under `validation` when applicable. The machine-readable envelope is in [`pdf-audit-v1.schema.json`](pdf-audit-v1.schema.json).
