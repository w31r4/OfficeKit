# Sign and verify PDFs

Use pyHanko for PDF signatures. The shipped `scripts/pyhanko_sign_provider.py`
adapter inventories signature fields and adds exactly one local-PKCS#12
approval or certification signature as a bounded incremental revision. The
separate read-only `scripts/pyhanko_provider.py` validates exact bytes and emits
typed integrity, trust, revision, and DocMDP evidence for an Agent. PyMuPDF,
MuPDF.js, pypdf, and qpdf are not signature-trust authorities.

## Resolve and probe the signing runtime

The adapter requires the pyHanko core library, not the separately packaged
`pyhanko` command. First resolve the exact `sign` task through the public
capability API. A signing runtime is installable only through an authorized,
hash-pinned managed pack; otherwise select an already-provisioned
`system-only` Python runtime in [provider setup](provider_setup.md). Do not
repair a missing runtime with `pip`, `uv`, a package manager, or a global
installation command.

```js
import { PdfFile } from "office-kit";
import { PdfProviders } from "office-kit/pdf/providers";

const inspection = await PdfFile.inspectPdf("input.pdf");

let resolution = await PdfProviders.resolve({
  task: "sign",
  provider: "pyhanko",
  inspection,
  savePolicy: "incremental",
  mutationAuthorized: true,
  credentials: ["local-pkcs12"],
  policyPath: ".office-kit/pdf-providers.json",
});
if (resolution.status === "installable") {
  resolution = await PdfProviders.ensure({ resolution, policyPath: ".office-kit/pdf-providers.json" });
}
if (resolution.status !== "ready") throw new Error(resolution.reason.message);
await PdfProviders.probe({ provider: "pyhanko", task: "sign", policyPath: ".office-kit/pdf-providers.json" });
```

After the selected route is ready, run the task-specific probes through the
same configured Python executable:

```bash
PYTHON_BIN="${OFFICE_KIT_PDF_PROVIDER_PYTHON:?select a ready pyHanko runtime first}"
"$PYTHON_BIN" scripts/pyhanko_sign_provider.py probe
"$PYTHON_BIN" scripts/pyhanko_provider.py probe
"$PYTHON_BIN" scripts/pdf_provider.py check --provider pyhanko --require
```

Neither adapter uses a system trust store, fetches certificates, CRLs, or OCSP
responses, invokes a CLI, or routes to another provider. The signer supports
local PKCS#12 credentials only; TSA, LTV/DSS, PKCS#11, remote signing, and
complete PAdES conformance remain external.

## Inspect and sign one exact source

Hash the source and credential immediately before use. Inspect current fields,
signature count, certification state, revision count, and the selected page's
unrotated CropBox before choosing a field mode:

```bash
PYTHON_BIN="${OFFICE_KIT_PDF_PROVIDER_PYTHON:-python3}"
SOURCE_SHA256="$(shasum -a 256 input.pdf | awk '{print $1}')"
CREDENTIAL_SHA256="$(shasum -a 256 /secure/signer.p12 | awk '{print $1}')"

"$PYTHON_BIN" scripts/pyhanko_sign_provider.py inspect input.pdf \
  --expected-sha256 "$SOURCE_SHA256" --page-index 0 --trusted-input \
  > tmp/pdfs/signature-inventory.json

# Invisible approval signature. A terminal gets a hidden prompt.
"$PYTHON_BIN" scripts/pyhanko_sign_provider.py sign \
  input.pdf tmp/pdfs/signed.pdf \
  --expected-sha256 "$SOURCE_SHA256" --trusted-input \
  --credential /secure/signer.p12 \
  --credential-sha256 "$CREDENTIAL_SHA256" --passphrase-stdin \
  --field-name Approval --field-mode create-invisible \
  --signature-kind approval --subfilter pades \
  --expected-signature-count 0 \
  > tmp/pdfs/signing-report.json
# Automation pipes stdin directly from its secret manager without staging the
# value in argv, env, or a file.
```

Use `--no-passphrase` only for a deliberately unencrypted PKCS#12. Secrets are
never accepted on argv or through an environment option and are omitted from
the versioned report. The provider rejects symlink credentials, stale hashes,
encryption, source overwrite, output collisions, oversized inputs/outputs,
unsupported runtime versions, and missing trust/isolation declarations.

Field modes are explicit:

- `existing` fills exactly one named empty signature field;
- `create-invisible` creates an invisible field on page 0;
- `create-visible` also requires `--page-index` and an integer
  `--box x1,y1,x2,y2` wholly inside an unrotated inspected CropBox.

A certification signature must be first and requires
`--docmdp-permission no-changes|fill-forms|annotate`. Finalize content and
Poppler visual QA before applying restrictive certification. A later approval
signature requires both the exact `--expected-signature-count` and
`--allow-existing-signatures`. This acknowledges a new revision; it never means
an earlier signer approved it.

The output is a distinct transaction: it preserves the complete source byte
prefix, appends one revision, adds one signature, passes internal integrity and
DocMDP validation, and is promoted without replacement. Re-run qpdf structure
inspection, explicit-root validation, and Poppler rendering before delivery.

## Validate one exact source

Record a fresh SHA-256, then choose one of two trust policies:

```bash
SOURCE_SHA256="$(shasum -a 256 input.pdf | awk '{print $1}')"
PYTHON_BIN="${OFFICE_KIT_PDF_PROVIDER_PYTHON:-python3}"

# Integrity/difference evidence only. Trust is deliberately reported false.
"$PYTHON_BIN" scripts/pyhanko_provider.py verify input.pdf \
  --expected-sha256 "$SOURCE_SHA256" \
  --trust-policy cryptographic-only \
  --require-signature \
  --require-all-integrity-valid \
  > tmp/pdfs/signature-cryptographic.json

# Delivery gate against caller-supplied trust roots.
"$PYTHON_BIN" scripts/pyhanko_provider.py verify input.pdf \
  --expected-sha256 "$SOURCE_SHA256" \
  --trust-policy explicit-roots \
  --trust-root /trusted/root-ca.pem \
  --other-cert /trusted/intermediate.pem \
  --revocation-policy hard-fail \
  --require-signature \
  --require-all-integrity-valid \
  --require-all-trusted \
  --require-docmdp-compliant \
  --require-all-bottom-line \
  > tmp/pdfs/signature-validation.json
```

`--moment` accepts an ISO 8601 timestamp with an explicit UTC offset. The
revocation policies are `none`, `soft-fail`, `hard-fail`, and `require`; choose
one deliberately. Network fetching is always disabled, so strict revocation
policies can succeed only when pyHanko already has adequate embedded/local
evidence. Do not weaken the policy merely to turn a report green.

The versioned `office-kit.pyhanko-verify.v1` report keeps these
facts separate for every signature:

- signed revision and ByteRange coverage;
- byte integrity and cryptographic validity;
- certificate path trust under the exact supplied roots, moment, and
  revocation policy;
- signer certificate identity, digest, and signature mechanism;
- timestamp evidence;
- difference-analysis modification level and changed form fields;
- DocMDP/FieldMDP constraints and DocMDP compliance;
- seed-value status and a policy-specific bottom line.

The adapter validates a private source snapshot under hard input, signature,
certificate, subprocess-time, stdout, and stderr budgets, then proves the
source and trust inputs did not change. An unsigned file can be inventoried, but
`--require-signature` fails it. A stale source hash, implicit/system trust,
encrypted input, unsupported pyHanko version, incomplete signature validation,
cryptographic/DocMDP failure, or unmet required gate fails closed with
structured JSON and no fallback. The configurable byte/time limits can be
lowered for a task but never raised above the adapter's hard maxima.

The report does **not** claim complete PAdES profile conformance. `trusted` means
only that pyHanko accepted the certificate path under the recorded validation
policy. An intact older ByteRange proves the signed revision is unchanged; it
does not mean the signer approved arbitrary later revisions. Review
`coverage`, `modificationLevel`, `docMDPCompliant`, timestamps, and every
signature in revision order.

## Refuse a forbidden P=1 certification change

For a requested content change on a fully verified DocMDP `P=1` certification,
do not append a revision, paint an overlay, or hand-write an audit. First save
the exact explicit-root verification report outside the delivery directory,
then let the canonical audit primitive bind its proven signature facts:

```bash
SOURCE_SHA256="$(shasum -a 256 inputs/source.pdf | awk '{print $1}')"
PYTHON_BIN="${OFFICE_KIT_PDF_PROVIDER_PYTHON:?select a ready pyHanko runtime first}"

"$PYTHON_BIN" scripts/pyhanko_provider.py verify inputs/source.pdf \
  --expected-sha256 "$SOURCE_SHA256" \
  --trust-policy explicit-roots \
  --trust-root inputs/credentials/test-root.pem \
  --require-signature \
  --require-all-integrity-valid \
  --require-all-trusted \
  --require-docmdp-compliant \
  --require-all-bottom-line \
  > tmp/pdfs/signature-validation.json

"$PYTHON_BIN" scripts/pdf_audit.py failed-closed outputs/audit.json \
  --source inputs/source.pdf \
  --provider pyhanko --provider-version 0.35.2 \
  --operation requested-content-change-under-docmdp-p1 \
  --reason "DocMDP P=1 permits no post-certification content changes." \
  --strategy read-only --probe-completed --plan-completed --source-inspected \
  --signature-verification tmp/pdfs/signature-validation.json \
  --require-docmdp-no-changes \
  --trust-root inputs/credentials/test-root.pem

"$PYTHON_BIN" scripts/pdf_audit.py validate outputs/audit.json \
  --source inputs/source.pdf \
  --require-operation requested-content-change-under-docmdp-p1 \
  --require-docmdp-no-changes \
  --trust-root inputs/credentials/test-root.pem
```

This route accepts only one intact, trusted, full-file P=1 certification under
the selected root with all signature policy gates passing. It publishes only
`audit.json`, including the source-bound trust/DocMDP evidence and explicit
no-mutation decision. A P=2 or P=3 form change needs its own field and policy
analysis; do not route it through this blanket refusal.

## Finalise one allowed DocMDP P=2 field

Use this route only when all preconditions are already known and the requested
result is a one-time static delivery value: one certification signature named
by the caller, `DocMDP P=2`/`fill-forms`, exactly one `FieldMDP Include` locked
field with an expected value, and one flat empty visible `/Tx` target. It turns
that target into a visible read-only decimal. It does not claim that an earlier
signer approves arbitrary later revisions, and it is not a general signed-form
editor.

First inspect the original and resolve the exact mutating capability. The trust
root is declared as caller-supplied evidence; it is never fetched or installed:

```js
import { PdfFile } from "office-kit";
import { PdfProviders } from "office-kit/pdf/providers";

const inspection = await PdfFile.inspectPdf("inputs/source.pdf");
let resolution = await PdfProviders.resolve({
  task: "fill-certified-form",
  provider: "pyhanko",
  inspection,
  savePolicy: "incremental",
  mutationAuthorized: true,
  credentials: ["caller-supplied-trust-root"],
  policyPath: ".office-kit/pdf-providers.json",
});
if (resolution.status === "installable") {
  resolution = await PdfProviders.ensure({ resolution, policyPath: ".office-kit/pdf-providers.json" });
}
if (resolution.status !== "ready") throw new Error(resolution.reason.message);
await PdfProviders.probe({ provider: "pyhanko", task: "fill-certified-form", policyPath: ".office-kit/pdf-providers.json" });
```

Then bind every precondition to exact source bytes. The output path must not
exist. The script creates a private source/root snapshot, checks the baseline
under that root, writes exactly one incremental revision without unrelated
metadata drift, verifies the output under the same root, and atomically
publishes its typed audit report. `--caller-isolated` means the caller has
already isolated this untrusted input; otherwise use `--trusted-input`.

```bash
PYTHON_BIN="${OFFICE_KIT_PDF_PROVIDER_PYTHON:?select a ready pyHanko runtime first}"
SOURCE_SHA256="$(node -e 'const fs=require("fs"),c=require("crypto"); process.stdout.write(c.createHash("sha256").update(fs.readFileSync(process.argv[1])).digest("hex"))' inputs/source.pdf)"

"$PYTHON_BIN" scripts/pyhanko_certified_form_fill.py probe
"$PYTHON_BIN" scripts/pyhanko_certified_form_fill.py fill \
  inputs/source.pdf outputs/approved-amount.pdf \
  --expected-source-sha256 "$SOURCE_SHA256" \
  --trust-root inputs/credentials/test-root.pem \
  --field ApprovedAmount --value 12500.00 \
  --expected-signature-field Certification \
  --expected-locked-field LockedAmount --expected-locked-value LOCKED-9000 \
  --caller-isolated \
  > outputs/audit.json

OUTPUT_SHA256="$(node -e 'const fs=require("fs"),c=require("crypto"); process.stdout.write(c.createHash("sha256").update(fs.readFileSync(process.argv[1])).digest("hex"))' outputs/approved-amount.pdf)"
"$PYTHON_BIN" scripts/pyhanko_provider.py verify outputs/approved-amount.pdf \
  --expected-sha256 "$OUTPUT_SHA256" \
  --trust-policy explicit-roots --trust-root inputs/credentials/test-root.pem \
  --require-signature --require-all-integrity-valid --require-all-trusted \
  --require-docmdp-compliant --require-all-bottom-line \
  > tmp/pdfs/approved-amount-signature-validation.json
```

The final report requires `coverage: entire-revision`,
`modificationLevel: form-filling`, `changedFormFields: ["ApprovedAmount"]`,
and `docMDPCompliant: true`; it also records the exact source prefix, root,
field before/after state, revision count, and no-replace transaction. Reinspect
with a separate form reader and render every page with Poppler before delivery.
Fail closed if any field is hierarchical/shared, any source revision already
follows certification, the lock set/value differs, the target is nonempty or
read-only, a signature is missing/untrusted, any non-target field changes, or
the operation would need reflow, a new signature, timestamp, LTV/DSS, or
interactive form preservation.

## Capabilities outside the shipped signer

Use an explicit external pyHanko workflow for PKCS#11/HSM credentials, remote
signing services, timestamp authorities, revocation-material embedding, LTV/DSS
updates, or a claimed PAdES profile. Keep private keys, tokens, PINs, and
passphrases outside scripts, logs, shell history, reports, and repository files.
Review pyHanko's official
[signing](https://docs.pyhanko.eu/en/latest/cli-guide/signing.html) and
[validation](https://docs.pyhanko.eu/en/latest/lib-guide/validation.html)
documentation.

After any later incremental form, annotation, DSS, or timestamp update, run the
typed validator again against the new exact bytes. A material content rewrite
normally belongs in a new version that is signed again.
