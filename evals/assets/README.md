# PromptBench locked corpus

This directory is evaluator-side test data. It is excluded from the npm
package and is never copied into an Agent trial except for the exact declared,
read-only input files.

The initial PDF boundary set is self-authored. It contains no customer data,
production credentials, production trust anchors or private signing keys, or
third-party sample assets.
`scripts/agent-eval-corpus-fixtures.py` is the reviewed source recipe; the
checked-in `integrity.json` pins the byte hashes actually used by PromptBench.

Use the evaluator Python to verify it:

```bash
/path/to/python3 scripts/agent-eval-corpus-fixtures.py verify
```

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

`pdf/corrupt/recoverable.pdf` and `pdf/corrupt/unrecoverable.pdf` are a
self-authored qpdf structural-repair pair. The recoverable file contains two
native-text pages and the exact document-level `repair-evidence.txt` attachment;
only its terminal `startxref` pointer is deliberately set to `0`. The malformed
control has a PDF header but no usable trailer or page tree. The matching
PromptBench oracle checks source/control identity, attachment bytes, strict
post-repair structure, and page-for-page Poppler pixels. Refresh the pair only
through `scripts/agent-eval-corpus-fixtures.py refresh-corrupt` and commit it
with the updated integrity manifest.

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
boundary assets, replace `generate` with `refresh-docmdp` in that command.
