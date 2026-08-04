# PromptBench locked corpus

This directory is evaluator-side test data. It is excluded from the npm
package and is never copied into an Agent trial except for the exact declared,
read-only input files.

The initial PDF boundary set, DOCX modern-comment and complex-table boundary
fixtures, and XLSX threaded-reply boundary fixture are self-authored. They contain no customer data, production credentials,
production trust anchors or private signing keys, or third-party sample
assets. `spreadsheets/reviewed-budget-nested.xlsx` is generated from the
OfficeKit threaded-comment fixture and then adds one explicit reply-of-reply;
its locked bytes are the input for the safe-refusal PromptBench case.
`documents/modern-comment-replies.docx` is generated from the OfficeKit modern
comment fixture and then adds one explicit reply-of-reply across the
`commentsExtended`, `commentsIds`, and `people` parts; its locked bytes are the
input for the Documents safe-refusal PromptBench case.
`documents/clinical-form.docx` is generated from a two-table OfficeKit document
and then adds a custom styled second table with vertical merges, a nested table,
one tracked cell, and one SDT; its locked bytes are the input for the complex
table topology safe-refusal case.
`scripts/agent-eval-corpus-fixtures.py` is the reviewed source recipe; the
checked-in `integrity.json` pins the byte hashes actually used by PromptBench.
The XLSX boundary source recipe is
`scripts/agent-eval-office-fixtures.mjs` (`generateXlsxNestedReplyBoundary`).
The DOCX boundary source recipe is the same module's
`generateDocxModernCommentReplyBoundary`.
The complex-table source recipe is
`generateDocxComplexTableTopologyBoundary` in the same module.
The PPTX SmartArt boundary source recipe is
`generatePptxSmartArtNotesCommentsBoundary` in the same module; its four-part
SmartArt graph, page-four NotesSlide, and modern root/direct-reply comment are
all self-authored and locked for a no-partial-output refusal case.
The branded-template source recipe is `scripts/agent-eval-branded-template.mjs`;
it generates the eight-slide `presentations/quarterly-board-template.pptx` and
the fixed-size `presentations/replacement-product.png`. The template is
self-authored and includes only synthetic text/data; its Master/Layout/theme,
notes/comments/transition, SmartArt, custom show, and embedded XLSX OLE parts
are preservation canaries for the bounded six-edit PromptBench transaction.
The XLSX operating-plan inputs under `spreadsheets/operating-plan/` are also
self-authored and locked: `actuals.csv` contains 24 months of synthetic actuals
and `assumptions.json` contains the three scenario profiles plus machine-check
requirements. They are input data only; the workflow must derive every output
metric through formulas and retain the two source hashes in its audit.

Use the evaluator Python to verify it:

```bash
/path/to/python3 scripts/agent-eval-corpus-fixtures.py verify
```

The `pdf/signing/docmdp-p1-final.pdf` fixture is a real self-authored
certification signature with `/Perms` → `/DocMDP` permission `P=1`. Its paired
`pdf/signing/test-pki/root.pem` is a public test trust root, not a private key.
The normal verifier checks its locked bytes, visible and metadata canaries,
ByteRange/CMS presence, and DocMDP structure without requiring pyHanko.

`pdf/signing/docmdp-p2-form.pdf` is a separate real self-authored certification
with `P=2`, one empty visible `ApprovedAmount` text field, and one
`FieldMDP Include`-locked `LockedAmount=LOCKED-9000` field. Its paired
`pdf/signing/test-pki/docmdp-p2-root.pem` is public-only. The fixture recipe
retains neither root nor signer private material. It is a controlled
finalisation test, not an example of arbitrary signed-form editing.

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
