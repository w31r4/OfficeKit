# PromptBench locked corpus

This directory is evaluator-side test data. It is excluded from the npm
package and is never copied into an Agent trial except for the exact declared,
read-only input files.

The initial PDF boundary set is self-authored. It contains no customer data,
production credentials, production trust anchors or private signing keys, or
third-party sample assets. One tightly scoped exception is the disclosed,
disposable PAdES-LTA **test** PKCS#12 pair described below; it exists solely so
the evaluator can require a real offline signing/timestamp/DSS chain. It must
never be reused for a customer, deployment, trust store, or production signing
operation.
`scripts/agent-eval-corpus-fixtures.py` is the reviewed source recipe; the
checked-in `integrity.json` pins the byte hashes actually used by PromptBench.

Use the evaluator Python to verify it:

```bash
/path/to/python3 scripts/agent-eval-corpus-fixtures.py verify
```

`presentations/strategy-review.pptx` is a self-authored four-slide PPTX
boundary fixture. Its first slide has a four-part SmartArt graph whose data
part owns one external child relationship. Slide four carries ordinary speaker
notes and a modern root/direct-reply comment thread as independent preservation
canaries. The matching prompt asks for all three mutations as one transaction:
the connected SmartArt graph is the reason to refuse; the notes and comments
are not presented as unsupported on their own. Refresh it through
`scripts/agent-eval-office-fixtures.mjs`, then review the raw OPC profile and
update `integrity.json` in the same change. The locked source itself must pass
public `PresentationFile.importPptx(...).verify({ visualQa: true })`; the
opaque graph is a mutation boundary, not a known source-layout error.

The `pdf/signing/docmdp-p1-final.pdf` fixture is a real self-authored
certification signature with `/Perms` → `/DocMDP` permission `P=1`. Its paired
`pdf/signing/test-pki/root.pem` is a public test trust root, not a private key.
The normal verifier checks its locked bytes, visible and metadata canaries,
ByteRange/CMS presence, and DocMDP structure without requiring pyHanko.

`pdf/ocr/mixed-bilingual-scan.pdf` is an eight-page self-authored mixed-source
fixture: six raster-only English/Chinese scan canaries encode upside-down,
sideways, and `+3°`/`-2°` skew cases in pixels; pages seven and eight contain
ordinary selectable-text canaries. It exercises an OCR preprocessing boundary,
not OCR provider availability. The published managed OCR route can handle an
explicit-language `skip` rewrite, but automatic orientation and deskew are not
part of its typed contract, so the matching PromptBench case requires an
audit-only safe refusal.

`pdf/signing/docmdp-p2-form.pdf` is a separate real self-authored certification
with `P=2`, one empty visible `ApprovedAmount` text field, and one
`FieldMDP Include`-locked `LockedAmount=LOCKED-9000` field. Its paired
`pdf/signing/test-pki/docmdp-p2-root.pem` is public-only. The fixture recipe
retains neither root nor signer private material. It is a controlled
finalisation test, not an example of arbitrary signed-form editing.

`pdf/signing/final-document.pdf` plus
`pdf/signing/test-pki/pades-ltv-{signer,tsa}.p12`,
`pades-ltv-root.pem`, and `pades-ltv-root.crl` form one repository-only,
offline PAdES-LTA test profile. Both P12 files have no passphrase and are
intentionally public/disposable; the root and empty CRL are likewise test-only.
The case must use the shipped bounded workflow, retain the source prefix, and
be checked by the evaluator-only explicit-root/CRL validator. The profile is
test evidence, not a PAdES conformance certificate or a generic TSA feature.

`pdf/corrupt/recoverable.pdf` and `pdf/corrupt/unrecoverable.pdf` are a
self-authored qpdf structural-repair pair. The recoverable file contains two
native-text pages and the exact document-level `repair-evidence.txt` attachment;
only its terminal `startxref` pointer is deliberately set to `0`. The malformed
control has a PDF header but no usable trailer or page tree. The matching
PromptBench oracle checks source/control identity, attachment bytes, strict
post-repair structure, and page-for-page Poppler pixels. Refresh the pair only
through `scripts/agent-eval-corpus-fixtures.py refresh-corrupt` and commit it
with the updated integrity manifest.

`pdf/redaction/multichannel-secret.pdf` is a self-authored four-page sanitize
fixture. Its fictional canary occurs in visible and white selectable text, a
scanned image, an invisible OCR layer, an annotation, a hidden widget, an
attachment, document metadata/XMP, an unreferenced decoded stream, and a prior
incremental revision. It exists to prove that a claimed high-trust redaction
removes every named channel instead of drawing an overlay. Refresh it only with
`scripts/agent-eval-corpus-fixtures.py refresh-redaction`, then verify the
fixture and commit its revised hash in `integrity.json` together with the source
recipe.

`pdf/richmedia/3d-review.pdf` is a self-authored two-page opaque-runtime
boundary. Its second page has one native `/3D` annotation with a binary payload,
default view, and activation dictionary, plus one `/RichMedia` annotation with
an asset/configuration/settings graph and JavaScript action. The fixture is
structural only: neither the corpus verifier nor an Agent may execute the
media or script, and no test result claims that a viewer can play it. The
matching task accepts a source-bound audit-only refusal unless a selected
provider can prove that every opaque object and its runtime contract survives
an incremental edit. Refresh it only with
`scripts/agent-eval-corpus-fixtures.py refresh-richmedia`, then verify and
commit its hash with `integrity.json` and the source recipe.

Refreshing fixtures is an intentional corpus update: regenerate the files,
review their structure and provenance, then commit the revised assets and
integrity manifest together. Generating the DocMDP fixture additionally needs
an explicitly selected, policy-authorized managed pyHanko interpreter; private
root and signer material exists only in a temporary directory during signing:

```bash
OFFICE_KIT_PROMPTBENCH_SIGNING_PYTHON=/path/to/managed/python3 \
  /path/to/evaluator-python3 scripts/agent-eval-corpus-fixtures.py generate
```

`--signing-python /path/to/managed/python3` is equivalent. Do not treat the
fixture passwords, public test root, or any future PKI material as production
credentials.

To refresh only the time-bound signed fixtures without changing the other locked
boundary assets, replace `generate` with `refresh-docmdp` or
`refresh-pades-ltv` in that command. The latter deliberately regenerates the
public disposable test PKI; review all five new hashes and the source recipe
together before committing.
