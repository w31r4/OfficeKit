# Agent black-box evaluations

`office-kit-agent-promptbench-v1` measures whether an Agent can turn a realistic Office/PDF request into a verified artifact through the published Skills and npm package. It evaluates the complete workflow, not isolated API calls:

```text
prompt + immutable inputs
  -> selected native Skill
  -> packed office-kit candidate
  -> explicit provider and save policy
  -> exact outputs + audit
  -> semantic, visual, security, and trace grading
```

The suite is evaluator-side repository infrastructure. `evals/`, the runner, locked assets, trial traces, and grading oracles are deliberately excluded from the npm consumer package.

## Current suite

- 41 cases: 21 PDF, 7 Documents, 7 Spreadsheets, and 6 Presentations. PDF is 51.2% of the suite and is required to remain an absolute majority, so a meaningful Office vertical slice never creates unrelated PDF filler prompts.
- 41 `ready` cases: 21 PDF cases, including the three-page cross-page Regional Revenue table extraction, a real signed-input PAdES/TSA/LTV capability-boundary refusal, and a self-authored mixed-scan OCR rotate/deskew capability-boundary refusal, seven bounded XLSX workflows (threaded-comment direct reply, nested-reply safe refusal, formula-assumption update, auditable operating plan, source-bound connection refresh-on-open, source-bound Pivot refresh-on-open, and source-bound opaque-enterprise local edit), seven DOCX workflows (classic-comment, mixed classic/modern board-review surgical edit, modern-comment nested-reply safe refusal, complex-table topology safe refusal, and source-bound header/footer text plus section page numbering), and six PPTX workflows (title plus fixed-topology rich-speaker-notes run edit, source-bound slide-name edit, source-bound complete section-boundary edit, closed-leaf slide clone, SmartArt/notes/comment-reply safe refusal, and a self-authored branded-template preservation transaction).
- Twelve locked, self-authored PDF corpus inputs are now ready. Seven are auditable safe-refusal boundaries: AES-256 owner/user permission split, native annotation reply chain, untagged complex report, Dynamic XFA packet, print-production risk structure, a RichMedia/3D opaque runtime boundary, and a real DocMDP P=1 certification signature with a public test root. The eighth is a real DocMDP P=2/FieldMDP form whose one permitted controlled finalisation is independently graded. The ninth is a two-file qpdf repair pair: a damaged xref/EOF source whose pages and attachment survive recovery, plus an unrecoverable comparison that must be rejected. The tenth is a four-page multichannel redaction source with selectable/hidden text, raster/OCR layer, active content, form/annotation/attachment/XMP/decoded-stream residues, and an old incremental revision; its success grader requires full rewrite, OCR evidence, residue removal, and Poppler masks. The eleventh is the self-authored three-page Regional Revenue table with merged and repeated headers, coordinates, spans, rotated label, footnotes, and parenthesized negatives; its success grader independently checks pdfplumber topology, CSV/JSON parity, geometry, no narrative leakage, and Poppler overlay evidence. The twelfth is an eight-page mixed scan with six image-only rotated/skewed pages and two born-digital canaries; its refusal grader proves the source profile and rejects unsupported rotate/deskew preprocessing before any mutation. The PAdES/TSA/LTV case reuses the signed P=1 source and public root to prove the baseline before requiring an explicit fail-closed refusal when the shipped provider reports no timestamp or LTV embedding capability. Every asset has a committed SHA-256/byte manifest and parser/render oracle; generic “no output” does not pass either kind of case.
- The branded-template case is now pinned as a self-authored eight-slide asset pair (`quarterly-board-template.pptx` plus `replacement-product.png`). Its evaluator-side oracle covers the six requested edits, Master/Layout/theme and opaque-part preservation, notes/comments/transition/SmartArt/custom-show/OLE retention, second import, and eight-page render evidence. The completed matrix at `/tmp/officekit-promptbench-branded-template-matrix-1785833930` used package SHA-256 `c27177ef1a44442c45ec05aac01213b07620b40755f2b02b1c7013b818f16417`, input-contract SHA-256 `9721cc094d5e8ed46c51037c67acf1782c679a846455075922f949aa234dfb6e`, and oracle SHA-256 `32495cf4a890cfcf5a1d0c970de0a0889c61ea5d664bf9945843f1eea07912f7`. Candidate passed `3/3` at `100/100`; the copied reference Skill passed `1/3` at `100/100` and failed two trials at `90/100` on explicit audit/trace compatibility fields. All six trials completed without timeout, and no third-party template bytes are included.
- Every family has both a success and a fail-closed case. Some advanced PDF cases accept either verified success or an explicit safe refusal.
- The default policy uses three trials per subject. Trial count is recorded per case rather than silently inferred by the Agent.
- Hosted CI for `da34c4f0` passed the Windows global-CLI lane. The Linux node lane was canceled by Actions in both attempts during the heavyweight default-template-library stage (no assertion failure was reported; [run 30899918301](https://github.com/w31r4/OfficeKit/actions/runs/30899918301)); the isolated local candidate tree completed the full `npm test` and release gates, and the cancellation remains an infrastructure blocker for this commit.

The committed JSONL contains evaluator-only expected outcomes, sources, and grade specifications. The Agent receives only `PROMPT.md`, declared inputs, the selected Skill, the installed candidate tarball, and exact deliverable paths. When a candidate/reference or repeat matrix uses one explicit `--run-root`, the runner creates one schema-versioned, read-only, hash-validated input snapshot for that case and copies those exact bytes into every trial. `run.json` records the fixture root, input-contract hash, and source hashes; a missing, changed, or mismatched cache fails preparation rather than silently regenerating a different input.

Every `run` has a bounded Codex execution deadline of 20 minutes by default. Use `--timeout-ms` or `OFFICE_KIT_AGENT_EVAL_TIMEOUT_MS` for a deliberate override. A timeout terminates the spawned Agent process tree, writes `timedOut: true`, the selected deadline, and a stable status `124` to `evaluator/exit.json`, then continues through the ordinary scorer. It is an incomplete hard-gate failure, never a safe refusal or a passing trial; this keeps a hung model/tool invocation from blocking an entire repeat matrix.

Use `matrix` for an auditable candidate/reference repeat run instead of manually
coordinating trial directories:

```sh
npm run eval:agents -- matrix pdf-damaged-xref-recovery \
  --subjects candidate,reference --trials 3 \
  --run-root /tmp/officekit-promptbench-qpdf-matrix \
  --timeout-ms 600000
```

The matrix uses one run root and therefore one immutable input fixture snapshot.
Before the first Agent starts, every trial must bind the same package tarball
SHA-256, input-contract/input hashes, and hidden-oracle fingerprint; a mismatch
fails closed instead of mixing incomparable records. All requested trials run
after a timeout or ordinary failure, and the root `matrix.json` records each
trial's subject, exit status, timeout flag, score, and report path. The command
exits non-zero unless every trial exits successfully and its case-specific
grader passes. `matrix` does not make asset-required cases ready and does not
replace the independent corpus/PKI acceptance required for those cases.

## Commands

```sh
npm run eval:agents -- validate
npm run eval:agents -- list --family pdf --status ready
npm run eval:agents -- show pdf-bounded-contract-id-replace
npm run eval:agents -- prepare pdf-bounded-contract-id-replace --subject candidate --trial 1
npm run eval:agents -- prepare pdf-encrypted-owner-policy-boundary --subject candidate --trial 1
npm run eval:agents -- prepare pdf-bounded-contract-id-replace --subject reference --trial 1
npm run eval:agents -- prepare pdf-source-bound-text-highlight --subject candidate --trial 1
npm run eval:agents -- run pdf-source-bound-text-highlight --subject candidate --trial 1
npm run eval:agents -- run pdf-overflow-replace-refusal --subject candidate --trial 1
npm run eval:agents -- matrix pdf-source-bound-text-highlight --subjects candidate,reference --trials 3 --run-root /tmp/officekit-promptbench-highlight-matrix
npm run eval:agents -- score pdf-overflow-replace-refusal --trial-root /absolute/trial/path
npm run eval:agents -- prepare xlsx-threaded-reply-resolve --subject candidate --trial 1
npm run eval:agents -- run xlsx-threaded-reply-resolve --subject candidate --trial 1
npm run eval:agents -- prepare xlsx-growth-assumption-update --subject candidate --trial 1
npm run eval:agents -- run xlsx-growth-assumption-update --subject candidate --trial 1
npm run eval:agents -- prepare xlsx-connection-refresh-on-open --subject candidate --trial 1
npm run eval:agents -- run xlsx-connection-refresh-on-open --subject candidate --trial 1
npm run eval:agents -- prepare xlsx-pivot-refresh-on-open --subject candidate --trial 1
npm run eval:agents -- run xlsx-pivot-refresh-on-open --subject candidate --trial 1
npm run eval:agents -- prepare docx-classic-comment-text-edit --subject candidate --trial 1
npm run eval:agents -- run docx-classic-comment-text-edit --subject candidate --trial 1
npm run eval:agents -- prepare docx-surgical-board-review-edit --subject candidate --trial 1
npm run eval:agents -- run docx-surgical-board-review-edit --subject candidate --trial 1
npm run eval:agents -- prepare docx-header-text-edit --subject candidate --trial 1
npm run eval:agents -- run docx-header-text-edit --subject candidate --trial 1
npm run eval:agents -- prepare docx-footer-text-edit --subject candidate --trial 1
npm run eval:agents -- run docx-footer-text-edit --subject candidate --trial 1
npm run eval:agents -- prepare docx-section-page-numbering-edit --subject candidate --trial 1
npm run eval:agents -- run docx-section-page-numbering-edit --subject candidate --trial 1
npm run eval:agents -- prepare pptx-title-and-notes-edit --subject candidate --trial 1
npm run eval:agents -- run pptx-title-and-notes-edit --subject candidate --trial 1
npm run eval:agents -- prepare pptx-source-bound-slide-name-edit --subject candidate --trial 1
npm run eval:agents -- run pptx-source-bound-slide-name-edit --subject candidate --trial 1
npm run eval:agents -- prepare pptx-source-bound-section-boundary-edit --subject candidate --trial 1
npm run eval:agents -- run pptx-source-bound-section-boundary-edit --subject candidate --trial 1
npm run eval:agents -- prepare pptx-closed-leaf-slide-clone --subject candidate --trial 1
npm run eval:agents -- run pptx-closed-leaf-slide-clone --subject candidate --trial 1
```

Generated PDF fixtures and the locked-corpus verifier require Python with ReportLab, pypdf, and Pillow. The eleven artifact-producing ready PDF graders additionally require pdfplumber; their applicable visual oracles require `pdftoppm`. The PAdES/TSA/LTV boundary is intentionally audit-only and does not claim a signature artifact. Set `OFFICE_KIT_AGENT_EVAL_PYTHON` to that evaluator interpreter and, only when it is not on `PATH`, set `OFFICE_KIT_AGENT_EVAL_PDFTOPPM`. A prepared PDF prompt binds its actual provider interpreter from `OFFICE_KIT_AGENT_EVAL_PROVIDER_PYTHON`, then legacy `OFFICE_KIT_PDF_PROVIDER_PYTHON`, then the evaluator interpreter. This keeps a small parser/render evaluator separate from a policy-authorized managed specialist runtime such as pyHanko; `run.json` records both paths. In Codex, use the Python executable returned by `load_workspace_dependencies` for the evaluator and a managed provider path only when the case needs it.

## Locked boundary corpus

The eleven ready signature/boundary/repair/table fixtures live below `evals/assets/`, outside the npm
payload. `integrity.json` records every declared file's SHA-256 and byte count;
the runner rejects a missing manifest entry, a mismatched byte stream, or a
symbolic link before copying a file into the read-only trial workspace. The
reviewed self-authored recipe is
`scripts/agent-eval-corpus-fixtures.py`; it intentionally contains only test
data and a user password that opens the encryption fixture. The committed
corpus never carries the owner password or retained private key material.

The evaluator-side oracle parses the source fixture independently of the
candidate provider. It verifies the actual AES-256 permission dictionary and
attachment/form canaries; `/IRT` reply, Popup, Highlight, author/date and
review-state graph; missing PDF structure tree plus visual-complexity canaries;
XFA template/datasets, repeat-subform, FormCalc, JavaScript and
`NeedsRendering`; or DeviceN, Separation, overprint, transparency, OCG and
OutputIntent dictionaries. A response passes only if that input proof, a
failed-closed audit bound to the source hash, an explicit inspection trace,
no-fallback/no-mutation evidence, and the case diagnostic all agree. At clean
commit `42cd55bdf78696e1b56302981a49ce7dbcc9325f`, all three AES candidate
trials passed at `100/100` against tarball SHA-256
`3adba43ab0bb20d736dbd03fcd069eab754ec6f3f3b15e3a0e7003fc51acf35c` and
candidate Skill SHA-256
`1baeae66ef4ba5723b395f2318537aa757f7f5eccd7d93840494b9c50932e0a7`.
The same tarball with the copied reference Skill (SHA-256
`0a09e468825a8be83345fd6c34e848c9c383bea66fc67e09dc36ecb5dfb2f0b1`) had
one `100/100` pass and two hard-gate failures. Those two reference trials left
the source intact and produced only `audit.json`, but used noncanonical audit
aliases such as `actual_provider` and `save_policy`; the evaluator correctly
did not infer required provider/save-policy/no-artifact proof from those
aliases. The same package and Skill hashes were then used for three-repeat
matrices of the other four fixtures: annotation reply-chain refusal was
candidate `3/3` and reference `3/3`; untagged complex PDF/UA refusal was
candidate `3/3` and reference `3/3`; Dynamic XFA refusal was candidate `3/3`
and reference `1/3`; print-production refusal was candidate `2/3` and
reference `1/3`. Every zero in the latter two reference matrices, and the one
candidate print zero, still preserved the source and emitted only an audit, but
failed hard gates because the Agent wrote a noncanonical or incomplete typed
no-artifact/no-mutation envelope. This is recorded as an audit-contract
reliability gap, not relabeled as a semantic PDF success. `pdf_audit.py
failed-closed` now owns the canonical audit-only refusal generation.
The additional AES-256 owner-policy repeat at
`/tmp/officekit-promptbench-aes-owner-matrix-v5` used candidate package
SHA-256 `133828263cb1f4541fa03825809eacf1055a37ff61cc9e4622c4a8bcd367447c`,
input-contract SHA-256 `272b25a9b97fef7cb392bb0a73f7baa88dde3d6f20db55fb4b462bdde213c21e`,
source SHA-256 `e6272671fc5f6a753d268373c19e5b87f7c31dc1120dd7ff66063d7963c04f4d`,
and oracle SHA-256 `8de2ba30fefbbd8baf04d3b2339e3734b59da8fa5a9c1ab64c066deda29ea751`.
The candidate passed `3/3` at `100/100`; the copied reference passed `1/3`
(`4/6` total trials passed, no timeout). The two reference failures preserved
the source and emitted only `audit.json`, but used explicit structured variants
of the refusal envelope that were not yet accepted by that matrix's grader.
The current grader now recognizes the tested snake/camel provider-version,
operation, source-preservation, save-policy, and no-artifact aliases without
accepting omitted or prose-only evidence; the regression fixtures remain a
compatibility guard. This result is an audit-contract repeat, not a claim that
the reference Skill or arbitrary encrypted-PDF permission edits succeeded.
At clean
commit `51d3275fa2e94aabe2c381a2a16dbc4d942f9054`, a new package SHA-256
`812067781a90d0efc8058fc841a1369eb6b1497e52f71a1fde6b9316f7730e47`
passed a fresh print-production matrix at candidate `3/3` and reference `3/3`,
all `100/100` with every hard gate. The candidate Skill SHA-256 was
`bce7065aba7e4836f171b9103c49857afef84cf3222ea0518ae91a857780c7fd`; the
reference Skill SHA-256 remained
`0a09e468825a8be83345fd6c34e848c9c383bea66fc67e09dc36ecb5dfb2f0b1`.
Every one of those six traces invoked `pdf_audit.py failed-closed`; the output
had only `audit.json`, `output: null`, an explicit provider/version, typed
source preservation, typed no-artifact checks, and no mutation command. This
is a new package identity, not a retroactive rewrite of earlier evidence. The
new DocMDP case uses the self-authored signed PDF SHA-256
`6ad55dd93543921c3b13d96f9cffed7a000ddea3b7da54643dae915034d19060` and
public test-root SHA-256
`ab15a064bf134b4c8409a08669b9308c5c9ba25d7d66dae74e99f30ccb7c606b`. Its
parser oracle proves the source title canary, `/Perms` → `/DocMDP`, a full-file
ByteRange, CMS contents, one DocMDP transform, and permission `P=1`; the
prepared prompt supplies the root and an explicit managed pyHanko runtime for
trust/difference validation before the requested title change is refused. Its
DocMDP hard gate requires the explicit-root pyHanko verification command and
the matching audit-bound signature, integrity, trust, DocMDP-compliance, and
bottom-line evidence; a probe or structural parse alone does not pass. The
certificate lasts ten years so this fixed test asset does not need a routine
annual rollover. The
fresh six-record matrix was run at base commit `885d5cdc` from one explicitly
recorded dirty worktree: its status SHA-256 was
`60ad721d87a3c1d0f8766a812201ddacaf7b60a946eff44e7882b30d6099272f` and
tracked-diff SHA-256 was
`cfef635e5b7581d7573eddaddc84dc6351e2bdb215f8f609a563700f2c3e8281`.
All three candidate and all three copied-reference trials used the same packed
candidate tarball SHA-256
`b29c915944dbac4d1641fd9a1e53757cc83da2ff2cd75be33960e4c503faf1fc`, the
same oracle fingerprint
`5e28dad9e2553d2fddfaa50a4e9a95177f9fcd0d696efe9c46f3e60bc0423e25`, and
the source/root hashes above. The candidate Skill SHA-256 was
`429bf9aec9b2ebc7a1c75b1674807252074a3e9e8c1b61a5276be622a6006a0c`; the
reference Skill SHA-256 was
`0a09e468825a8be83345fd6c34e848c9c383bea66fc67e09dc36ecb5dfb2f0b1`.
Every trial passed `100/100` with every hard gate. Both subjects ran
explicit-root pyHanko validation and the typed
`pdf_audit.py failed-closed --require-docmdp-no-changes` route, which binds the
real verification report, selected public root, full-file coverage, P=1 policy,
and no-mutation decision into the sole `audit.json` output. This is evidence
that both workflow texts can use the published safe-refusal primitive. P=2,
P=3, FieldMDP, and arbitrary signed PDFs remain outside that refusal primitive
except for the separate, explicitly constrained P=2 finalisation described
below.

The current repeat was rerun from `3f5c1e86de284d7d898e46e33e544326e579b032`
with the published `python-specialists` `3.13.14-oat.2` managed runtime and
the same real signed source. The matrix at
`/tmp/officekit-promptbench-docmdp-p1-matrix-v3` passed candidate `3/3` and
copied-reference `3/3` at `100/100` (six completed trials, no timeout). Every
record bound package SHA-256
`133828263cb1f4541fa03825809eacf1055a37ff61cc9e4622c4a8bcd367447c`, input
specification SHA-256
`79a91a0854f6d19041611fd06c50ce18aabc269d2f1894419006189070569ef4`, source
SHA-256 `6ad55dd93543921c3b13d96f9cffed7a000ddea3b7da54643dae915034d19060`,
public root SHA-256
`ab15a064bf134b4c8409a08669b9308c5c9ba25d7d66dae74e99f30ccb7c606b`, and
oracle fingerprint
`411607d09ad126ac7fe1abdec67c34ec8671710017cba5a242370d518e579813`.
The candidate/reference Skill tree hashes were
`970c046507c32b8568f8826d51ca5952c3e77007e7300eaaf6d9bc2301c9335d` and
`0a09e468825a8be83345fd6c34e848c9c383bea66fc67e09dc36ecb5dfb2f0b1`.
Tracked repository content was unchanged (`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`); the protected local authorization document remained untracked and was
not part of any trial input or package. All six traces ran explicit-root
pyHanko verification and `pdf_audit.py failed-closed --require-docmdp-no-changes`,
binding the trusted signature, full-file coverage, P=1 policy, and no-change
decision into the sole `audit.json` output.

The Dynamic XFA boundary was then rerun after making the refusal envelope
explicit in the case prompt. The v6 matrix at
`/tmp/officekit-promptbench-dynamic-xfa-matrix-v6` passed candidate `3/3` and
copied-reference `3/3` at `100/100` (six completed trials, no timeout). Every
record bound package SHA-256
`133828263cb1f4541fa03825809eacf1055a37ff61cc9e4622c4a8bcd367447c`, input
specification SHA-256
`4315c7c006fe378b8b715cba23aaecfdb171abf47f7eed644554edf38b081c5f`, source
SHA-256 `727ebf02feb699c4d9e657fa65808aade639dda8d31fa6ad6bb06b3cfb2e44f9`,
and oracle SHA-256
`d0a949c4e29a82ba4b187b5491bceee063eaefcd3409cbde1773e77faf48bb80`.
Both subjects emitted only `audit.json`, preserved the source, reported an
actual provider/version with `silentFallback:false`, recorded the requested
mutation as unexecuted, and proved no modified or partial artifact. The
grader accepts equivalent structured aliases such as `operations[]`,
`savePolicy.mode`, and nested `validation.noArtifact/sourceUnchanged`; it
still rejects prose-only claims. This is a repeatable Dynamic XFA safe-refusal
contract, not XFA reflow, FormCalc/JavaScript execution, flattening, or
AcroForm conversion.

At clean commit `f09a68f2596f483af311e69f944ca82166ef4cbe`, the fresh
`pdf-pades-ltv-signature` matrix passed candidate `3/3` and copied-reference
`3/3` at `100/100`. All six trials used the same candidate package SHA-256
`80817445041369bbab1f3b52f1d1500ed0c0ce199108b9ccc665a2b7347c3207`, input
specification SHA-256
`5e6df51746f41d047cb05b7effeedccee77fe123946bf12026fba89c6002216f`, source
SHA-256 `6ad55dd93543921c3b13d96f9cffed7a000ddea3b7da54643dae915034d19060`,
public test-root SHA-256
`ab15a064bf134b4c8409a08669b9308c5c9ba25d7d66dae74e99f30ccb7c606b`, and
oracle SHA-256
`82ce4d72b9fca9d3285c227d25c1b7a29eef71964d9df9d67700e7cd933b090b`. The
matrix run root was
`/tmp/officekit-promptbench-pades-f09a68f2/pdf-pades-ltv-signature`. Every
trial performed a safe refusal: `output` was `null`, `mutationAttempted` and
`performed` were false, the source hash was preserved, and the only delivered
artifact was the capability-bound audit. The candidate and reference traces
both now bind structured false values for timestamp-authority support, LTV
evidence embedding, and complete PAdES profile conformance; this is a
repeatable refusal-contract result, not a claim that TSA/LTV/PAdES conformance
or signature upgrading is implemented.

`pdf-docmdp-allowed-field-fill` uses a different self-authored source and
public root: P=2 allows form filling, while a FieldMDP Include transform locks
only `LockedAmount=LOCKED-9000`. The only success contract is to convert the
empty visible `ApprovedAmount` `/Tx` widget into the canonical decimal
`12500.00` with the published `pyhanko_certified_form_fill.py` primitive. The
primitive rejects any post-certification baseline revision, non-flat form,
wrong source/root hash, non-P=2 policy, wrong FieldMDP lock, locked target,
noncanonical amount, output collision, and symlink path before publication. It
uses one incremental revision, makes the target static/read-only with a normal
appearance, suppresses unrelated metadata updates, revalidates through the
explicit root, then publishes without replacement. The hidden evaluator does
not trust that report: it reopens both bytes with pypdf, checks the original
signature contents and ByteRange, DocMDP/FieldMDP transforms, catalog
references, exact field scope, strict source prefix, revision count, and a
Poppler target-widget mask. Repository smoke has exercised that fixture through
the managed pyHanko runtime and the independent evaluator. One isolated Agent
candidate trial against the dirty implementation candidate also passed all 37
applicable checks at `100/100`; it remains neither a clean-commit result nor a
candidate/reference three-repeat claim.

A subsequent clean-commit matrix provides that missing repetition. At
`5599aa6094db8940fecb2197d0f66f3b54e3c5fc`, all three candidate trials and all
three reference-Skill trials passed every hard gate and scored `100/100`. Every
recorded worktree was clean (empty-status SHA-256
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`), used
the same candidate tarball SHA-256
`7e31440be693705e40cf9652f067b2d4921c104049ecd30706edc857656d2217`, and
used oracle fingerprint
`2bd647ff60210eeb2e4932424bd4723402934dd6403995e5781cc374a77a370e`. The
candidate Skill tree hash was
`0a5c820ef1d8014966e0e398616c8282020e3c27727b146e4a7fb34d0acff848`; the
reference Skill tree hash was
`0a09e468825a8be83345fd6c34e848c9c383bea66fc67e09dc36ecb5dfb2f0b1`. The
reference copy recorded no package-name patches. This is a constrained P=2
workflow compatibility result under the declared source/root inputs, managed
provider, and hidden structural/visual oracle—not a general PDF-fidelity or
reference-superiority claim.

P2 typed-operation trace evidence is deliberately narrower than a shell
interpreter: it accepts either a direct packaged Script path or a variable
bound to that path or its scripts directory in the same completed command
record. The accepted paths are the installed `.agents` PDF Skill or the
installed `office-kit` package; an unbound variable, a cross-command binding,
or an arbitrary same-named workspace script does not count. The independent
post-fill verifier uses the same rule and accepts only its escaped shell line
continuations—not an ordinary following line. The official OfficeKit
Poppler-runtime variable counts
as a render command. This preserves shell-safe invocation without treating a
trace as authorization evidence.

The evaluator normalizes the published audit envelope and structured equivalent
forms: an explicit
no-mutation result can be `executed: false`, `performed: false`/`"none"`,
`mutationAttempted: false`, or a typed `refused_not_attempted` operation. It
also recognizes the corresponding explicit no-artifact fields. It never trusts
that audit alone: it separately inventories the trial output directory and
requires it to contain only `audit.json`.

There are four ready PPTX graders. The source-bound slide-name edit uses the canonical `launch-review.pptx` two-slide fixture; the rich-notes edit uses a separate `rich-notes-review.pptx` two-slide fixture; the section-boundary edit uses a four-slide complete native-section partition fixture; and the clone uses `release-review.pptx` with a deliberately small notes/comments/chart leaf profile plus one relationship-free custom-show run action.

The ready XLSX, DOCX, and PPTX fixtures are generated through the public OfficeKit facade and need no Python provider. Their independent native visual graders require `soffice`, `pdfinfo`, and `pdftoppm`; a missing executable reports `grader-unavailable`, never a product pass or failure. The rich title/notes case requires the named title and one exact ordinary notes run edit, fixed paragraph/run counts and order, a preserved character bullet plus auto-numbered sibling paragraph, fixed slide topology/direct background/supporting text, exactly the target slide XML and notes XML as changed package parts, a byte-stable appendix, byte-bound audit provenance, second import, and LibreOffice/Poppler page evidence. The source-bound slide-name case changes only the first SlidePart's `p:cSld/@name`; it independently requires fixed topology, all non-target package parts byte-identical, source-part/name/attribute audit binding, successful second import, and pixel-identical native renders of both pages. It deliberately permits Open XML SDK canonicalization inside the changed target SlidePart rather than claiming a lexical one-attribute XML delta.

`--subject candidate` installs `skills/<family>/skills/<skill>`. `--subject reference` copies the matching Skill from the pinned `reference/office-artifact-tool/skills` submodule into that trial and changes only its `office-artifact-tool` package-name occurrences inside the isolated copy. The reference-source commit and complete Skill-tree hashes are locked by `skills/reference-sync.json` and checked before preparation. Both subjects install the same freshly packed `office-kit` tarball, so the comparison changes the Skill instructions rather than the product candidate.

The ready `pdf-damaged-xref-recovery` case uses two self-authored locked inputs. The recoverable PDF has a complete two-page object graph and an attachment but a broken `startxref` pointer and missing `%%EOF`; qpdf must repair it with warnings, keep the attachment, and preserve both pages under Poppler comparison. The second input has no trailer, page tree, or EOF and must be rejected. The hidden oracle renders the damaged source and promoted output through Poppler, parses only the repaired output strictly with pypdf, verifies both source hashes in the audit, and rejects any ReportLab/PdfWriter/page-merge raster reconstruction. The Agent-facing contract requires the packaged qpdf `checkAfter` result to remain under `validation.qpdfCheckAfter` with `status: clean` and `exitCode: 0`; renamed near-miss fields do not satisfy the audit gate. A qpdf-unavailable environment may produce an explicit audit-only safe refusal because the case outcome is `success-or-safe-refusal`; an unverified approximation never passes.

The first stable contract trial at clean `5ad33e50a3524703ddf84d51571890a2708e3b40` passed the candidate subject `1/1` at `100/100`. It bound candidate tarball `0d599353ec5f87655fdd9e28e052bc47d4dbb83919f33c975e76e25321ac4a62`, input-contract `f988f5024e690600d599be67452c6dde35c35a97e0da95049849d23c3f998ec6`, recoverable input `83bf5cb65b6a6919583d76777aea7ee73828cdd52ee3c37a883aab0e55be3c03`, unrecoverable input `279d99a2f2d0ecf23e518a2fbfa062ec646257e13a57d0cffa9bb4447ba296ac`, and oracle `b337215690faa37c1fdaab59c737491d8e9c524d629bd9219d2e5c12a1cf1fb2`. The audit carried the canonical `validation.qpdfCheckAfter` plus fresh inspection, both source hashes, repair warnings, attachment identity, and byte-identical Poppler pages.

The complete candidate/reference repeat matrix at clean `381f21d02a987f1b071e5283d2d6f441c20e2bd0` closed the remaining repetition gap: all six trials passed `100/100` (`candidate 3/3`, `reference 3/3`) with the same package, input-contract, two source, and oracle fingerprints above. The matrix used one preflighted immutable run root and started no Agent until every trial identity matched; this is a compatibility result for the bounded qpdf repair workflow, not a general PDF repair or arbitrary recovery claim.

A current-main repeat at `6f470aede64f48c67445e54b767599ba5cd64bfc` reran the same six trials from `/tmp/officekit-promptbench-damaged-xref-matrix-20260804` with package SHA-256 `848e749368cc9dc2816b1043d0ca94f9eaa032361ab674f4fc0ae931b199bb5d`, input-contract SHA-256 `f988f5024e690600d599be67452c6dde35c35a97e0da95049849d23c3f998ec6`, recoverable/unrecoverable inputs `83bf5cb65b6a6919583d76777aea7ee73828cdd52ee3c37a883aab0e55be3c03` and `279d99a2f2d0ecf23e518a2fbfa062ec646257e13a57d0cffa9bb4447ba296ac`, and oracle `e6e11570c6f9a90a1568fbbcfbf8810cf79c564408dd89d004cdc72bbf86fe93`. With the bundled workspace Python evaluator, candidate passed `3/3` at `100/100`; the copied reference Skill passed `1/3` at `100/100`, with two hard-gate failures at `80/100` and `70/100` from missing source-provenance/both-input binding and provider/no-fallback trace fields. All six Agents completed without timeout and all six business hard-gate envelopes passed. This current repeat exposes a reference Skill compatibility regression; it does not weaken the shipped candidate result or broaden qpdf repair claims.

The fresh current-main attachment-quarantine repeat at
`/tmp/officekit-promptbench-attachment-quarantine-matrix-20260805-v2` completed
all six trials without timeout. It bound candidate package SHA-256
`1763a98609bcc6355ca50e4e7ec294ebfb96f324d8f716bc93847ff34c07d40a`,
input-contract SHA-256 `1c0d5682acb4b729815844863db22cb253d5674d951e812241d763f1a890de79`,
source `7e02892e252150d257ce41660f0c303a647792a57ba2c0ba22fe5988bc79feb4`,
and oracle `05e50ad50b9aeb7d3b0a5c75e4a2b2151d540da6445f274b637ec2354b905ddb`.
Candidate and copied-reference Skills each passed `3/3` at `100/100`. The six
payloads were extracted into traversal-safe quarantine with duplicate names,
exact hashes, immutable source, no open/execute, final-page render review,
and direct audit evidence for source preservation, all hashes, contained
paths, duplicate separation, and `attachmentsOpenedOrExecuted: false`. This
closes the bounded attachment-quarantine repeat contract; it does not broaden
the primitive into arbitrary attachment execution or mutation.

### PDF multichannel redaction matrix

The final `pdf-redact-multichannel-secret` matrix was run from candidate commit
`090194bfe381b1df2ac106b6031e03fe8b929170` after the evaluator's mask,
structured-operation, and quoted-script-path fixes. The managed private
capability cache supplied PyMuPDF `1.27.2.3`, Tesseract 5, and Poppler; no
global Python runtime was used. Candidate and copied-reference Skills each
passed `3/3` trials at `100/100` (six completed trials, no timeout or refusal).
The matrix run root was
`/tmp/officekit-promptbench-redact-managed-090194bf-rerun7`, with candidate
package SHA-256
`80817445041369bbab1f3b52f1d1500ed0c0ce199108b9ccc665a2b7347c3207`, input
specification SHA-256
`43267335d6780c1ca104dfc1729201feac8292d7cb1c00e4fd73dae3f1038eb1`, source
SHA-256 `f162221bea256ec3c96bcf79a9ab57e447c4f292e833b2ccb3e0b3b5546c4fdc`,
oracle SHA-256
`46f282037e99b042c9b434030a3ed3b9f0a21efd53dfabccbcd63675fd104075`,
candidate Skill tree SHA-256
`1a6fd0363b50921a974ef989f9fcd554c1fad326d4f1e4f3284fdcfe0c322787`, and
reference Skill tree SHA-256
`0a09e468825a8be83345fd6c34e848c9c383bea66fc67e09dc36ecb5dfb2f0b1`.
The independent oracle confirmed four pages, removal of every canary and
residual channel (selectable/hidden text, raster/OCR image and index, active
content, form, annotation, attachment, XMP, decoded streams, and old
incremental revision), a full rewrite with one revision and no original-byte
prefix, source immutability, and the declared Poppler masks. This is a
bounded self-authored multichannel fixture and typed-operation compatibility
result; it does not generalize to arbitrary hostile PDFs or fuzzy OCR intent.
Earlier diagnostic runs are not evidence: they were discarded after the
evaluator corrected its visual masks and trace normalization.

### PDF Regional Revenue table matrix

The bounded `pdf-cross-page-table-extraction` candidate/reference repeat matrix
now passes candidate `3/3` and copied-reference `3/3` at `100/100` from one
immutable run root. All six trials used the same typed workflow:
provider check, explicit read-only table plan, the shipped
`pdfplumber_extract.py table` primitive, Poppler page rendering/overlays, and
canonical audit validation. The matrix identity is candidate package
`d32bf6883e04eaccde71655571ebf0f387bde4d67f7d6967fd1ae8339481ee6d`, input
contract `3001e80795108176f00100e0eec1ca4c340da568a869d6d09cf6b70d44fb7d66`,
source `f06f948511f60df356ddfc7bcb82e87428a358bbe0c523e070c6acf38c458d74`,
and oracle `793d00a276acc36c936fd7bb670580d3c0659617d17ac824cfa9ea2fbaf6d778`.
Every trial produced 78 geometry-bound cells, three merged title spans,
two separate footnotes, matching JSON/CSV hashes, and passed the per-page
bbox overlay, no-narrative-leakage, source-provenance, and no-fallback gates.
This is compatibility evidence for the bounded typed table workflow, not a
general PDF table-understanding claim.

### PDF RichMedia/3D opaque-preservation refusal matrix

The `pdf-richmedia-opaque-preservation` matrix now closes this high-risk
boundary as an explicit safe refusal. The two-page self-authored source has a
normal cover plus a second page containing a 3D annotation, default view/model,
RichMedia resources, activation data, and a JavaScript canary. Current typed
routes can inventory that opaque closure but cannot prove viewer runtime
semantics after even a cover-only mutation. The prompt therefore requires an
inspection/probe followed by an audit-only `failed_closed` result: no
`reviewed.pdf`, `savePolicy.strategy: "none"`, an unchanged source hash, an
unexecuted operation, and no provider fallback. The contract explicitly rules
out pypdf, ReportLab, PDF.js, direct content-stream edits, and other bypasses.

The v3 candidate/reference repeat matrix at
`/tmp/officekit-promptbench-richmedia-matrix-v3` passed all six trials at
`100/100` (candidate `3/3`, copied reference `3/3`, no timeout). Every trial
emitted only `outputs/audit.json` with normalized `provider`, `operation`,
`savePolicy`, source-provenance, and no-fallback evidence. The trial package,
input contract, source, and oracle identities are respectively
`2835ba6168e46541130c0e5c50b06acfcdd368c0306f4828dc93cbd0bfe74422`,
`151454a783b6cae4f9c376334abaa0c2496a22704643010f96e030a3d8804432`,
`d09e3997f0124717a9cd569899bf6ed123de7ee8a36a9a85bc149b37fa84579d`, and
`00599100e638506a15c76b6270239e886dfd10704a6d1ab6e76a627f7f099f54`.
This evidence is a bounded refusal contract; it does not claim 3D/RichMedia
editing, runtime preservation, or general opaque-part authoring.

### PDF mixed-scan OCR refusal matrix

The `pdf-mixed-scan-ocr-boundary` matrix used one immutable eight-page
mixed-profile scan (`image-only` pages with 180-degree and 90-degree rotations,
plus two born-digital canaries) and the same packed candidate for every trial.
The required operation asks for searchable OCR while correcting upside-down
pages and 2--4 degree skew. Because the published OCRmyPDF adapter exposes no
typed rotate or deskew primitive, the only accepted result is a source-bound
`failed_closed` audit with no PDF output. The final candidate/reference matrix
passed all six trials at `100/100` (`candidate 3/3`, copied `reference 3/3`).
Every trial recorded the same source SHA-256
`bd9cc618bc2a83e95c207a37a8f10f177905b3852549950f5c483649128c1096`, provider
`OCRmyPDF 17.8.1`, missing rotate/deskew capabilities, and no mutation or
fallback. The normalized refusal envelope binds `status: "failed_closed"`,
`savePolicy.strategy: "none"`, top-level `provider.silentFallback: false`,
an unexecuted requested operation, source preservation, and explicit audit-only
output-directory checks. Matrix identity is candidate package
`133828263cb1f4541fa03825809eacf1055a37ff61cc9e4622c4a8bcd367447c`, input
contract `66b94b40c729b4767cf51ba7c0363cdae07c6cc343e2d3f1479d81dfe5595646`,
source `bd9cc618bc2a83e95c207a37a8f10f177905b3852549950f5c483649128c1096`, and
oracle `473de083a41eefea5370c4dc2bbb04500695ebab63dafa9e303e73be96e286e1`.
This repeat matrix proves the shipped fail-closed OCR boundary for the real
scan fixture and reference-Skill contract compatibility; it does not claim
general rotate/deskew or OCR quality.
The contract and repeat evidence were finalized in commit `97026755`.

### PDF AcroForm visible-preserved matrix

The final `pdf-acroform-visible-preserved` candidate repeat used one generated,
immutable one-page AcroForm with text, radio, checkbox, blank TIN, and blank
signature canaries. The Agent filled only `full_name`, `company_type`,
`address`, and `effective_date`; it kept the form interactive and used
incremental save. Each trial reopened the result, checked widget appearance
states and field values, rendered the final page through independent Poppler
`pdftoppm`, and ran `pdf_audit.py validate` after the typed fill in the same
trace. Candidate passed `3/3` at `100/100` with no timeout, source mutation,
ad-hoc writer, or undeclared output.

The matrix identity is candidate package
`70edfe048ceb9b7e95a5ec817a23999f174285938cb76021bf08634b186fbdb9`, input
contract `0eacb72e03ebbde263b26ebfed25052e68142f0d856f620c8b3b325f0fbd2e7c`,
source `0ebb7bca3cd52a185ec6b68fe5b0acf52de86718af6d02322714d827caf6424e`,
and oracle `229c62a094f720a1d24c18371fb6bafffc0d1cafa1e0cf43965a985d362fde16`.
The candidate Skill tree is
`ee7c34e93718838f61dad23da8d5d289a687d156514f59478ccf5891d0d2923e`; all
three trials were prepared from commit `af02fb31133ee7a363b82ffac476ea2dbbcd7b80`
with the same package and input identities.

An earlier candidate/reference diagnostic run exposed two evaluator-facing
reliability issues: quoted installed Script paths were not recognized by the
trace parser, and the reference Skill selected rewrite or omitted the required
incremental/audit envelope. The parser now accepts the normal quoted paths and
has a regression test; the reference result remains a compatibility gap rather
than a waived failure. This is a bounded editable-form transaction, not a
general AcroForm hierarchy editor or signature-preserving arbitrary mutation.

## Isolation and provenance

Each preparation creates a fresh trial tree outside the repository by default:

```text
<run-root>/<case>/<subject>-trial-<n>/
  workspace/
    .agents/skills/<skill>/
    inputs/
    outputs/
    PROMPT.md
    node_modules/office-kit/
  evaluator/
    package/<candidate.tgz>
    run.json
    trace.jsonl
    final.txt
    report.json
```

`run.json` records the git commit, whether the source worktree was dirty, hashes of its status and tracked diff, the actual packed tarball hash, Skill-tree hash, input fingerprints, reference-only package-name patches, and an oracle fingerprint. That fingerprint covers the case outcome/grade specification plus the checked-in runner and applicable PDF or Office grader source bytes, so a later evaluator implementation change invalidates an older prepared trial rather than silently regrading it under a new oracle. The tarball SHA-256 is the authoritative identity of the candidate bytes, including when a local WIP is intentionally evaluated; a recorded HEAD is never presented as if it alone identified a dirty trial. The record also fingerprints `PROMPT.md`, the workspace package/lock files, `.agents`, and the complete installed `node_modules` tree, including file modes and symlink targets; changing the Skill or executed dependency graph therefore fails a hard gate. Skill copying filters ignored Python bytecode, the installed tree is read-only, and Codex runs with Python bytecode generation disabled, so local cache noise cannot enter the immutable Skill subject. File inputs are made read-only and SHA-256 checked after execution. Directory inputs such as test PKI are recursively made read-only and hashed as a deterministic tree, so added, removed, renamed, mode-changed, or content-changed files fail the source-immutability gate.

The runner starts Codex with an ephemeral, ignored-config `workspace-write` sandbox. This is useful process isolation, but it is not an oracle-confidentiality boundary on every host. A production benchmark must mount only the trial workspace into a no-network container or VM and keep the evaluator directory and repository unavailable to the Agent.

## Scoring boundary

Current generic hard gates verify:

- oracle version, completed execution, trace, and final response;
- immutable prompt, Skill, installed dependency graph, and file/directory inputs;
- exact required artifacts, non-empty bytes, basic PDF/OOXML magic, and valid JSON deliverables;
- no modified artifact for fail-closed outcomes;
- a valid `failed_closed` audit when one is emitted;
- correct branching for `success-or-safe-refusal` cases.
- regular-file outputs only and no undeclared success deliverables.

Generic gates alone do **not** constitute a passing task score. Reports for cases without a case grader continue to say `partial-generic-only` or `generic-refusal-gates` and keep their semantic, visual, security, and trace evidence in `pending`.

Eleven ready PDF cases have complete case-specific grading:

| Case | Machine | Visual | Security | Trace |
| --- | --- | --- | --- | --- |
| `pdf-bounded-contract-id-replace` | Independent page/text/font/box assertions | Poppler renders all five pages; non-target pages must be pixel-identical and the page-3 diff must stay inside the source text mask | Raw, extracted, decoded-stream, and metadata residue scan; one `startxref`/`%%EOF`; canonical audit hashes must match final bytes | PyMuPDF/version, explicit `sanitize`, no fallback, shipped probe before edit, typed mutation, and no low-level stream mutation |
| `pdf-source-bound-text-highlight` | pypdf confirms exactly one new native `/Highlight` on page 2, exact RGB/review metadata, source-text-stable geometry, and quadrilaterals bounded to the independent pdfplumber text box | Poppler/Pillow requires three nonblank same-size pages, pixel-identical non-target pages, and a page-2 diff contained in the source text mask | Exact input/output hashes, an unrotated source-page snapshot, no caller-supplied coordinates, decodable streams, and one rewrite revision | MuPDF/version, explicit `rewrite`, probe + inspect before typed `add_text_highlight`, output re-inspection/render, no fallback or direct native/object mutation |
| `pdf-cross-page-table-extraction` | The published `pdfplumber_extract.py table` primitive proves three pages × six rows, merged title topology, repeated headers, exact 78 geometry-bound cells, parenthesized negatives, page/coordinate/span/confidence fields, and JSON/CSV parity; explicit `*`/`Note:` footnotes stay in a separate JSON field rather than becoming cells | Poppler renders all three source pages; a table-bbox overlay audit passes on pages 1–3 and rejects blank/edge-clipped pages | Source and named JSON/CSV deliverable hashes are bound in the audit; read-only extraction cannot publish a rewritten PDF or fabricated cell | pdfplumber/version, explicit `read-only`, provider check + table plan before typed extraction, Poppler overlay and audit validation after extraction, no fallback or narrative leakage |
| `pdf-overflow-replace-refusal` | Independent ReportLab font metric plus pdfplumber 70pt-box proof and structured audit geometry | Not applicable | No partial artifact, no mutation claim, and source provenance | PyMuPDF capability evidence, no fallback, and no mutation command after failed preflight |
| `pdf-acroform-visible-preserved` | Independent pypdf field-tree, exact value, widget-topology, `/AP`, `/V`, `/AS`, `NeedAppearances`, and editability assertions | Poppler renders every page; only the four requested widgets may change; TIN, signature, the unselected radio, and a pre-checked checkbox remain pixel-stable | Exact original-byte prefix plus one appended revision, sensitive fields blank, all streams decodable, canonical audit hashes match | pypdf/version, explicit `incremental`, inspect/check/plan before typed fill, Poppler and audit validation after mutation, no ad-hoc writer |
| `pdf-attachment-quarantine-inventory` | Independent pypdf enumeration of document/page FileSpecs, exact raw identity/MIME/size/SHA fields, duplicate preservation, and extracted-byte equality | The immutable source remains the only PDF; no page-transform artifact is accepted | Traversal-safe regular files, exact payload set, source/manifest provenance, explicit no-open/no-execute evidence | pypdf/version, explicit `read-only`, inspect/check/plan before typed extraction, audit validation after extraction, no ad-hoc parser or payload execution |
| `pdf-active-content-public-sanitize` | pypdf proves the fixture contains root/additional JavaScript, Launch/SubmitForm actions, attachments, invisible text, a comment, populated form values, and personal metadata, then proves every channel is inert | Poppler renders every page and permits changes only inside the removed widget-value/comment masks | All canaries absent from object strings, streams, attachments, annotations/forms, text, and metadata; one revision; original prefix absent; canonical audit hashes match | Probe and route plan before typed scrub, then standalone inert residue scan, Poppler render, audit byte validation, no fallback or low-level mutation bypass |
| `pdf-redact-multichannel-secret` | Independent pypdf/PyMuPDF checks prove the four-page source's selectable/hidden text, raster/OCR layer, active content, form, annotation, attachment, XMP, decoded-stream, and old-revision canaries are all removed after the requested typed redaction | Poppler renders all four pages; permitted differences are confined to the declared text, raster/OCR, annotation, and form masks, with non-target pixels stable | Full rewrite, no original-byte prefix, one revision, no residue in raw/decoded streams, metadata/XMP, attachments, forms, annotations, hidden text, or OCR extraction; source SHA and audit hashes bind the transaction | Managed PyMuPDF/version and Tesseract evidence, explicit `sanitize`/full-rewrite policy, typed `redact_text` + `redact_ocr_text` + `scrub`, probe/plan before edit, post-edit audit/render, no fallback or low-level content-stream bypass |
| `pdf-greenfield-accessible-report` | Strict pypdf catalog and structure-tree traversal proves six pages, title/language, H1-H3, one logical Table spanning pages, Figure alt text, Link annotation/StructParent/OBJR, reading-order IDs, and running artifacts | Poppler/Pillow renders all six pages, rejects blank or edge-clipped pages, checks common geometry, and requires readable table text on both physical segments | Canonical source/output hashes, no PDF/UA overclaim, and separate modeled, veraPDF-machine, and human-review evidence | `artifact-tool`/version, explicit `rewrite`, no fallback, shipped typed report example, Poppler evidence, and no ad-hoc PDF writer |
| `pdf-merge-reorder-stamp-links` | Independent pypdf proves exact six-page source order, preserved boxes/rotation, 20% watermark placement, six outlines, six named destinations, and six resolved internal links | Poppler/Pillow maps every output page to its source; four non-target pages must be pixel-identical and both report pages must change | Manifest plus all source/output hashes, one revision, decodable streams, navigation resolution, and watermark absence from sources | pypdf/version, `rewrite`, check/plan before typed merge, typed Poppler comparison and multi-source audit after mutation, no ad-hoc writer |
| `pdf-docmdp-allowed-field-fill` | Independent pypdf proves the signed P=2 baseline, exact FieldMDP Include lock, one empty visible target, `12500.00` output value, static/read-only target, unchanged locked/non-target fields, original signature contents, and stable catalog references | Poppler/Pillow requires every page to remain nonblank and confines every changed pixel to the target widget; the shipped P2 smoke separately requires a native MuPDF render whose glyph pixels clear the widget borders | Original bytes must be a strict output prefix, exactly one revision is appended, the original signed range remains bound to the baseline, and the raw typed pyHanko audit binds source/output hashes, explicit root, integrity, trust, DocMDP, FieldMDP, changed field, and no-replace transaction | Published P2 probe and finalisation primitive before post-fill explicit-root verification and render; no fallback or ad-hoc form/object writer |

One clean candidate run of `pdf-source-bound-text-highlight` at commit `535d875504f84ccd469cb05922ce94528cfd14d8` passed `100/100`: all hard gates plus 6 machine, 3 visual, 3 security, and 8 trace checks passed. It bound package SHA-256 `c1569217b49725d0e31cc818fc5e2ee035eccc8690d7ee49c57a3e8dc74d33de`, copied-Skill SHA-256 `dc4a294e27b76540cf43e5a7ff1989ccbdcb822ea6e621d9ead413feef9fc20f`, input SHA-256 `3e31e282b7702940c572316958bb3d3263e54f62357605c84303b2fa1a7f31da`, and oracle SHA-256 `20286e3071ecdf9381db2c0543c96dd13c2082c43d3d815003e851d81455fc09`. This is one candidate run, not the default three-repeat matrix or a reference-Skill comparison.

A subsequent clean six-trial matrix at
`a1f966e646e06a1d9f8d8c89830dcaa6a54a09fc` makes the repeat result explicit:
the OfficeKit PDF Skill passed candidate `3/3` at `100/100`; the reference
Skill passed `2/3` at `100/100`. The remaining reference trial produced a
semantically and visually correct native highlight, but its audit omitted the
required `rewrite` save-policy and explicit no-fallback fields. Its raw score
was `90/100`, then the trace hard gate correctly forced its final score to
zero. Every record used the same clean worktree status hash
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`, source
SHA-256 `3e31e282b7702940c572316958bb3d3263e54f62357605c84303b2fa1a7f31da`,
candidate tarball SHA-256
`7e31440be693705e40cf9652f067b2d4921c104049ecd30706edc857656d2217`, and
oracle fingerprint
`0d21f3da37fb369bbafa641c143b48a9515d7bd7576247b1672e471fcf9beae6`. The
candidate/reference Skill tree hashes were respectively
`0a5c820ef1d8014966e0e398616c8282020e3c27727b146e4a7fb34d0acff848` and
`0a09e468825a8be83345fd6c34e848c9c383bea66fc67e09dc36ecb5dfb2f0b1`; the
reference copy recorded no package-name patches. This is a bounded workflow
and auditability comparison, not a general PDF-fidelity or Skill-superiority
claim.

`pdf-bounded-contract-id-replace` now also has a clean, six-trial matrix at
`f57db7ba61eed23987c85127a88f02f13ebe028e`: candidate and reference Skills
each passed `3/3` at `100/100`. The fixed five-page source SHA-256 was
`1b016ee5263d83e29c554e6af14808bd30537a5342de29b7391a90e4e7595995`; every
record bound the clean-worktree hash
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`, the
candidate tarball SHA-256
`7e31440be693705e40cf9652f067b2d4921c104049ecd30706edc857656d2217`, and
the oracle fingerprint
`9f7e1e88614e820ad548dfa93cabc80bcd2be4ff49385e620092a41af0c30eaa`.
Candidate/reference Skill tree hashes were respectively
`0a5c820ef1d8014966e0e398616c8282020e3c27727b146e4a7fb34d0acff848` and
`0a09e468825a8be83345fd6c34e848c9c383bea66fc67e09dc36ecb5dfb2f0b1`; the
reference copy again recorded no package-name patches. Each trial used an
explicit-policy, receipt-backed private `python-specialists` provider cache
after `resolve` → `ensure` → `probe`, reporting PyMuPDF `1.27.2.3`; no global
Python dependency was used. The oracle proved the one page-3 fixed-width
replacement, unchanged Helvetica 11-point geometry, pixel-identical pages
1–2 and 4–5, a target-mask-contained page-3 diff, no old-term residue, one
rewrite revision, and the shipped typed `replace_text` operation. This is a
bounded source-owned replacement contract, not a general PDF reflow claim.

The category weights are machine 45, visual 25, security 20, and trace 10. A category earns its weight only when every check in it passes; a not-applicable category is removed from the denominator. `rawScorePercent` preserves the evidence-weighted result; a failed safety hard gate forces `scorePercent` to zero. `taskPassed` is true only when generic hard gates and every applicable case category pass. Missing evaluator dependencies produce `grader-unavailable` with an infrastructure error, never a candidate failure or an inferred pass. A present but unreadable or unrenderable candidate PDF is instead a definitive graded failure.

The PDF oracle is evaluator-side and never copied into the Agent workspace. It uses pypdf/pdfplumber for semantic and structural evidence and Poppler/Pillow for visual evidence, independently of the PyMuPDF mutation provider. Audit claims and Codex command traces are graded separately from final bytes.

## Ready Office slices

`xlsx-threaded-reply-resolve` begins the same black-box discipline for Office files. The runner generates a canonical `Forecast!F19` workbook with a root threaded comment and one direct reply, then treats it as an immutable uploaded XLSX. The independent grader decodes the output package rather than trusting the model: it checks one threaded-comments part and one persons part, exact preservation of the original comment IDs/person IDs/dates/text/target/order, one new GUID-backed direct root-child reply, resolved state for all three comments, and absence of a legacy `comments*.xml` note downgrade. It also binds audit source/output hashes, explicit OfficeKit/rewrite/no-fallback evidence, public `SpreadsheetFile.importXlsx`/`exportXlsx` trace evidence, a second-import assertion, and a LibreOffice-to-PDF/Poppler nonblank render of every final page.

The bounded contract deliberately rejects reply-of-reply and branched graphs. The ready `xlsx-threaded-nested-reply-boundary` case uses a self-authored locked XLSX with root, direct-reply, and reply-of-reply identities. Its independent oracle requires all three threaded comments/person records and unresolved state to remain inspectable, permits only a source-bound `failed_closed` refusal with no modified workbook, rejects classic-note downgrade and scratch flattening, and grades the refusal trace separately from the generic no-artifact gate. An autonomous Agent result is recorded only after the runner executes the candidate/reference matrix, rather than inferred from the repository smoke.

The final immutable-trial matrix at `/tmp/officekit-promptbench-xlsx-nested-reply-29c` passed candidate/reference `3/3` at `100/100` (six completed trials, no timeout). All trials share package SHA-256 `0d599353ec5f87655fdd9e28e052bc47d4dbb83919f33c975e76e25321ac4a62`, input-contract SHA-256 `bd5821b95238a8f793756805c1996c9ef4d2e618a67b2b1bfc67bb8927d5d666`, fixture SHA-256 `1b16849b5cff680d1b7c7914c5f35e43df2ea76a8cb0ce27252fa2d49be228a4`, and oracle fingerprint `900b1a4b272f8b7b2b99c62153aee8e4975e3ba8df6c022201cd8bb726b93ac5`. The corrected grader accepts the reference Skill's `saveStrategy`/`save_strategy` and `fallback_policy`/`fallback` audit spellings through one canonical adapter; it does not weaken the no-artifact, source-immutable, no-export, or no-fallback gates.

One clean candidate trial at commit `d558a924ad63528f2b2dca5e1bbeb1fb0dc120a7` passed `100/100`: all 6 machine, 1 visual, 2 security, and 6 trace checks passed, including immutable prompt/Skill/package/input gates. It bound package SHA-256 `f6889efc2cff246b59bc557761e998299fdb30466d98111096d25d562d72f618`, copied-Skill SHA-256 `17a4b6a9f13e7921f953ac704ac353ce0d55267ac722a84615d1dfee2fe63d53`, input SHA-256 `f2b0839f53a6e875af27d65291352e191774207e0d55a613d08f016407884208`, and oracle SHA-256 `c5696086a6a840c0e448983da895381338c4243043b7214c92d8a71c004cc991`. This is one candidate trial, not the default three-repeat matrix and not a reference-Skill comparison.

`xlsx-growth-assumption-update` adds a second, non-comment Spreadsheet path: the generated two-sheet operating plan allows only `Forecast!B9` to move from 8% to 10%. The independent XLSX grader decodes both packages, requires the `B5:B7` formula chain and recalculated values `110`, `121`, and `133.1`, protects `Forecast!B10`, sheet names/order/identity, and the byte-stable `Approved Baseline` worksheet. It permits exactly the target worksheet part to change, binds a public OfficeKit rewrite audit to source/output bytes, requires a second import, and uses LibreOffice/Poppler to require changed Forecast pages plus a pixel-stable baseline page. The published workflow is an explicit bounded profile, not a claim that arbitrary imported workbooks can be safely reconstructed or globally recalculated.

One clean candidate trial at `3df67f5a083758eb1f6fc0c37cdc6c53f228e2eb` produced `100/100`: all 5 machine, 2 visual, 2 security, and 6 trace checks passed, including immutable prompt/Skill/package/input gates. It bound package SHA-256 `720400c754d9cdbfaa949120c7939467702202d2cf507f7bef91d5068c1a7503`, copied-Skill SHA-256 `7c59bf7f7085482414399c54147003f8eaa00c06654c7724cc23793b62c65edf`, input SHA-256 `0a40d6c2ba10a3f358508a55197e9474e9779db142b61a979f880591519c47b8`, and oracle SHA-256 `995a48cdfd62c098d123b82dde4e0636976af694575c86f72a77ecded3815d82`. It is one candidate trial, not the default three-repeat matrix or a reference-Skill comparison.

`xlsx-auditable-operating-plan` promotes a forward, self-authored fixture into the ready set. The immutable inputs are a 24-month `actuals.csv` (`b758991f9fcf6c029bd1a5c6a6af37f30e5aa6f2bf063bd878895eb270665a3e`) and a three-profile `assumptions.json` (`dcb46836d23936b3240a55e9337f49afdf68f12561460c732e1e0d4899932d58`). The published workflow creates five sheets—`Sources`, `Assumptions`, `Forecast`, `Dashboard`, and `Checks`—and requires formula-backed revenue, gross profit, EBITDA, and cash chains; Base/Upside/Downside validation; low-cash conditional formatting; two root threaded comments with person identities; line and pie charts; frozen headers; second import; model/native render evidence; and an audit binding both input hashes and the output hash. The independent grader requires at least 120 formulas with no formula errors, both validation and warning rules, exact chart coverage, passing check cells, all five frozen sheets, nonblank native pages, no fallback, rewrite policy, and the published typed workflow trace. This is a reproducible greenfield operating-plan contract, not a claim that arbitrary imported workbooks can be safely rebuilt.

`xlsx-connection-refresh-on-open` covers a deliberately narrower external-data safety boundary. Its generated source has one recognized `Fixture warehouse` connection with native ID `7`, explicit `refreshOnLoad=true`, a fixed command and opaque extension canary, and one associated QueryTable whose own `refreshOnLoad` is already false. The independent grader requires exactly the connection semantic to move to false: it compares the parsed connection projection, requires the complete normalized ConnectionsPart residual plus both QueryTable and Table parts to remain stable, and permits only `xl/connections.xml` to differ. Every other part or semantic drift is a hard failure. It binds source/output hashes to a no-fallback OfficeKit rewrite audit, requires a second import, and checks native output pixels. The trace accepts the published no-overwrite workflow or the complete public typed import/edit/export sequence; a scratch XML patch does not pass. This is not a general provider, credential, QueryTable, or host-refresh editor.

`xlsx-pivot-refresh-on-open` is the analogous but independent Pivot cache safety boundary. Its generated source contains one named `Revenue by region` native PivotTable on `Pivot Summary`, one uniquely owned cache definition with explicit `refreshOnLoad=true`, and byte-bound PivotTable/cache-record canaries. The only permitted semantic change is that cache-root request becoming false. The grader decodes the XLSX package rather than trusting the model: it requires one named Pivot root, one cache definition, one cache-record part, identical paths, byte-identical PivotTable/cache-record parts, an otherwise normalized-identical cache definition, and exactly that cache definition as the changed package part. It binds source/output hashes to a no-fallback OfficeKit rewrite audit, requires the capability to be withdrawn after a second import, and checks native output pixels. The trace accepts the published no-overwrite transaction or the complete public typed import/edit/export sequence; a scratch XML patch does not pass. This neither refreshes data nor claims to suppress manual, macro, external-data, or other host-triggered refreshes.

`xlsx-opaque-enterprise-local-edit` is the ready high-fidelity preservation slice. Its four-sheet source contains a typed connection/QueryTable, native Pivot/cache, sparkline, threaded root/direct-reply graph, array-formula plus XLDAPR metadata canary, and disconnected Power Query, slicer/cache, and combo bar+line ChartPart bytes. The public workflow changes only `Assumptions!B4` from `0.065` to `0.07` and the ordinary line chart `enterprise-line-chart` title; it does not claim to edit or refresh any advanced object. The independent raw-OPC grader requires the source/output inventories and every advanced canary hash to remain identical, accepts exactly `xl/worksheets/sheet1.xml` and `xl/drawings/charts/chart1.xml` as changed parts, reimports and verifies the two targets, and checks native Data/Summary render canaries by requiring a common stable suffix of at least two rendered pages. This avoids treating platform-specific chart continuation pagination as an untouched-sheet failure while still rejecting drift in the trailing canary pages. A direct XML patch, dynamic-array recalculation, combo/slicer/Power Query authoring, or any opaque-part drift fails closed.

`docx-classic-comment-text-edit` adds a deliberately narrower Documents slice. The runner generates a DOCX with one uniquely anchored classic Word comment, then the public workflow may change only that comment's text. The independent oracle decodes `word/comments.xml` and `word/document.xml`: it requires the same native comment ID, author, initials, creation time, target paragraph, `commentRangeStart`, `commentRangeEnd`, and `commentReference`, plus unchanged visible paragraph text and the requested comment text. It rejects a modern/reply graph (`commentsExtended.xml` or `people.xml`), a deleted/recreated comment, any topology change, a model-to-plain-text downgrade, missing byte-bound OfficeKit/rewrite audit evidence, or an unshipped/scratch workflow. LibreOffice-to-PDF plus Poppler requires the same nonblank page count and pixel-identical body render; structural XML is authoritative for comment balloons because headless LibreOffice does not make their visual treatment a stable oracle.

`docx-header-text-edit` adds the complementary page-furniture transaction. The generated DOCX has one uniquely used default HeaderPart with two ordinary text paragraphs, plus a PAGE footer and immutable body canaries. The Agent may change only `Northwind | Internal` to `Northwind | Reviewed` through the shipped no-replace workflow. The independent oracle reads raw OPC bytes rather than trusting a model: it requires exactly one `word/headerN.xml` part to differ, the entire package inventory and every other part—including the PAGE footer—to remain byte-identical, and the changed HeaderPart to become byte-identical again after normalizing precisely one source/replacement `w:t` payload. It separately proves companion/header/body text, source/output audit hashes, provider/version/no-fallback/rewrite evidence, target section/reference/part binding, second import, and a real LibreOffice/Poppler render with the requested header change visible. A scratch XML patch, a footer/PAGE replacement, a shared or rich page-furniture graph, or any second HeaderPart mutation fails closed rather than being accepted as a broad DOCX header editor.

`docx-footer-text-edit` is the symmetric, separately invoked FooterPart transaction. Its generated DOCX has one ordinary target and one ordinary companion in a uniquely used default FooterPart, a PAGE header canary, and immutable body canaries. The Agent may change only `Northwind | Internal` to `Northwind | Reviewed` through the shipped footer workflow. The independent raw-OPC oracle allows exactly one `word/footerN.xml` part and one normalized target `w:t` payload to differ; it proves the complete HeaderPart—including the PAGE field—body, package inventory, source/output audit binding, second import, and visible LibreOffice/Poppler change remain correct. A header mutation, scratch XML patch, shared/rich page furniture, or a second FooterPart mutation fails closed. The paired public entry points intentionally do not turn page furniture into a generic cross-kind editor.

`docx-section-page-numbering-edit` is a source-bound section-metadata transaction. Its generated DOCX has three native sections and a PAGE footer in each. The Agent may change only imported block 1 from `{ "start": 1, "format": "lowerRoman" }` to `{ "start": 1, "format": "decimal" }` through the shipped no-replace workflow. The independent oracle parses raw `word/document.xml` rather than trusting the model, permits only that package part to differ, masks only the target canonical `w:pgNumType`, and normalizes equivalent namespace-declaration placement and attribute order before binding every other XML element, attribute, text node, section, relation, and package part. It separately proves the upper-letter sibling, terminal section, three FooterParts, input/output hashes, source precondition, second import, provider/no-fallback/rewrite audit, first-page visual change, and stable later-page pixels. A guessed block, direct XML patch, field replacement, section reorder, footer edit, or any unrelated package drift fails; the case does not claim to add PAGE fields or implement general pagination.

`docx-surgical-board-review-edit` is the first composite source-bound DOCX edit in the ready set. Its immutable packet combines one classic comment with an Office modern root/direct/nested reply graph, a tracked insertion, a content control, TOC and REF fields, a footnote, a VML watermark, three mixed sections with first/even headers and PAGE footers, and a complex table with merge/SDT/nested-table canaries. The Agent may change only the recommendation run, the `Data migration`/`Amber` risk cell, and the uniquely anchored classic comment text. Because the codec deliberately treats mixed classic/modern comment families as source-bound, the shipped workflow first imports and verifies the typed paragraph/table targets, then uses the public `DocumentFile.patchDocx` package boundary for exactly `word/document.xml` and `word/comments.xml`; it never rebuilds the document or edits the modern graph. The independent oracle requires those exact two changed parts, byte-stable `commentsExtended.xml`/`commentsIds.xml`/`commentsExtensible.xml`/`people.xml`, unchanged footnotes/styles/headers/footers/fields/revisions/watermark, unchanged comment identity and anchors, source/output audit hashes, second import, and nonblank native pages. Ambiguous anchors, a missing modern identity part, table topology drift, a scratch XML script, or any silent fallback fails closed. This is a composite surgical transaction, not general Word reflow, mixed-comment authoring, or arbitrary package patching.

`docx-modern-comment-reply-boundary` now has a self-authored locked DOCX with one root, one direct reply, and one reply-of-reply spread across `comments.xml`, `commentsExtended.xml`, `commentsIds.xml`, and `people.xml`. The independent oracle requires every person/durable/parent identity, unresolved state, and the anchor canary to remain inspectable; because nested reply creation remains outside the bounded contract, the only accepted result is a source-bound `failed_closed` refusal with no modified DOCX. It rejects classic-comment downgrade, flattened graphs, scratch XML/export traces, and silent fallback. The final immutable matrix at `/tmp/officekit-promptbench-docx-modern-reply-30b` passed candidate/reference `3/3` at `100/100` (six completed trials, no timeout). Every trial bound package SHA-256 `0d599353ec5f87655fdd9e28e052bc47d4dbb83919f33c975e76e25321ac4a62`, input-contract SHA-256 `0eff80b550d73d897bf0c52079e99e75ea3a6c28bb5f1a9a9284c23ea9655f47`, fixture SHA-256 `30060ed6ba571dc81eb01394819b6107b5a050360349dd656630a3118ddab888`, and oracle fingerprint `ba56df205ff17b0675402dfdd3f4fc9c80403500eceb2c7006d8c885b3e31fce`.

`pptx-title-and-notes-edit` is the deliberately bounded rich-notes Presentations slice. The runner generates a two-slide deck and permits only one uniquely named title shape plus the second ordinary run in the first paragraph of one canonical source-bound rich NotesSlide on `Go-no-go decision`. The source has two paragraphs: the first owns two ordinary runs and a character bullet; the second owns one run and an `arabicPeriod` auto-number starting at 2. The public workflow rejects absent, irregular, ambiguous, or already-changed title/notes input; it preserves slide/order/title/notes identities, geometry, direct background, all non-target run text/styles, bullet/number semantics, and the unchanged appendix. It never accepts `notes.text` or `textFrame.setText()` as a substitute for fixed-topology editing.

The independent oracle inspects package bytes rather than trusting the candidate model: exactly `ppt/slides/slide1.xml` and `ppt/notesSlides/notesSlide1.xml` may differ; it checks the requested title/run text/style change, exactly one NotesSlide body shape, identical paragraph/run topology and siblings, a namespace-placement-normalized hash of all non-body NotesSlide XML, all package paths, a byte-stable appendix part, and stable native LibreOffice/Poppler appendix pixels. Namespace placement is normalized only because Open XML SDK may hoist an equivalent declaration while serializing; all non-body elements, attributes, and text remain bound. It also requires a byte-bound OfficeKit rewrite audit, source-bound paragraph/run coordinates, second import, no silent fallback, and a shipped rich-notes workflow or direct typed public round trip. Repository smoke and the candidate/reference matrix below cover this semantic version; the reference comparison remains explicitly partial. Fields, hyperlinks, picture bullets, list/body styles, arbitrary reflow, comments, and broader relationship edits remain outside this case and must not be inferred as supported.

A clean same-source matrix at `/tmp/officekit-promptbench-pptx-rich-notes-matrix-v2`, run from candidate commit `505b7b86505fbc927104d4234597adb8dd75c379`, proves the OfficeKit candidate path: all three candidate trials passed `100/100`. The three reference trials used the same input SHA-256 `34a8f37f44df08279f7f1c1e95aeb0dde54585459dc058c9e551296e8ee94de2`, input-contract SHA-256 `24c59b2bd843f3e960e22d226df3bb43238b1c2cb75e78f1a77aa1b28eb1067d`, candidate tarball SHA-256 `133828263cb1f4541fa03825809eacf1055a37ff61cc9e4622c4a8bcd367447c`, and oracle SHA-256 `c715cedcbcc272aff372fdead7aefb303ca0ffec252051b5970abb97c8d600f1`; candidate/reference Skill trees were `1b03b731f0e79e7d239c4c5d1b8092de204b03e458ec1842c611f6c0319856dc` and `fa9f6badd10fc4957b489ce10c9ccc2d4edc6357cf2396df38c447de6e7d0506`, with 34 isolated reference package-name patches. The copied reference Skill passed `1/3`; its two misses produced the requested visible artifact but used noncanonical green `#15803D` instead of the candidate profile's `#0F766E` and omitted one or more required no-fallback/rewrite/typed-roundtrip/second-import audit fields. Their raw scores were `45/100`, and the trace/machine hard gates correctly forced final scores to zero. No fixture, prompt, package, or oracle changed between subjects. This is a completed candidate operation with an explicit reference-Skill compatibility gap, not a general rich-notes or PowerPoint reflow claim.

`pptx-source-bound-slide-name-edit` exercises the complementary non-visual boundary. It gives the Agent the same immutable two-slide package but permits only the unique first-slide `p:cSld/@name` change from `Go-no-go decision` to `Go decision: controlled rollout`. The independent grader identifies the SlidePart by its original package path, requires the output to retain slide order and every visible title/notes/background/text canary, permits exactly that SlidePart to differ, and requires both native-rendered pages to be pixel-identical. It rejects an apparent title edit, a notes edit, a package topology/relationship change, a modified appendix, a missing source-bound audit record, a skipped second import, or a scratch/untyped workflow. Because the Open XML SDK can serialize the one changed XML part canonically, its promise is semantic `p:cSld/@name` correctness plus byte-identical non-target parts, not a byte-level claim about the changed XML's attribute ordering or namespace declarations.

`pptx-source-bound-section-boundary-edit` covers the complementary native PowerPoint navigation metadata transaction. The generated four-slide source has three canonical Office 2010 sections that partition the deck in order once: Context owns slides 1–2, Decision owns slide 3, and Appendix owns slide 4. The only permitted result moves slide 2 across that one boundary, retaining all three facade IDs, labels, native GUIDs, and ordinal positions. The independent raw-OPC oracle decodes `p14:sectionLst` from `ppt/presentation.xml` and maps its numeric slide IDs back through the presentation relationship graph; it does not trust the candidate model or audit. It requires the exact replacement partition, four unchanged SlideParts and native-rendered pages, identical package topology, and exactly one changed package part (`ppt/presentation.xml`). The source-bound audit must bind both byte hashes, the whole source/target partition, changed-section records, second import, non-section semantics, model render, and no-fallback rewrite policy. This is a complete-partition boundary edit, not a general section authoring, renaming, reordering, or slide-graph mutation API.

One clean candidate trial at `ffdaca79014afd82038dfcbf0002dcaacb51c54d` produced `100/100`: all hard gates plus 5 machine, 2 visual, 2 security, and 6 trace checks passed. It bound tarball SHA-256 `cfdda4953699f45443006005391dc72ee5fa29372e3f6dfbb2ac27ea98d6a063`, copied-Skill SHA-256 `5e2b9019cc499898b7b85f76577cf20547e4df9db6c33462b31ecd5f82c744fd`, immutable input SHA-256 `4ef9783e70e94a643e723b37f76515127aa1d56ed8bd2c3ec7f888f030842239`, and oracle SHA-256 `5318bad1c5a6028346c3515ee9bb5fb9700dabe109e61d2c482c49b2b71f644d`. The source worktree was clean. This is historical one-trial evidence only.

The strict same-source matrix now closes that gap. At clean `dcac620360d04944925ff20c29470055cd870a16`, one shared schema-v1 fixture snapshot bound every candidate and reference trial to input SHA-256 `00951b1248e3dfbc13cf034f4e384c603187e55c0e300d5b4951d7a73f340133` and input-contract SHA-256 `accf2718c1cbe94536fd0df9e352ed4e684f7c6c444f9ae6580b19fe95686005`. All six runs used candidate tarball SHA-256 `0d599353ec5f87655fdd9e28e052bc47d4dbb83919f33c975e76e25321ac4a62` and oracle SHA-256 `4501f78bc022ee05d42daf90296e27ac6898660c651b535bae6f0be441fa34ab`; every hard gate and score passed at `100/100`. The OfficeKit Presentations Skill passed candidate `3/3`; the copied reference Skill passed `3/3`. Their immutable installed-tree hashes were respectively `1b03b731f0e79e7d239c4c5d1b8092de204b03e458ec1842c611f6c0319856dc` and `fa9f6badd10fc4957b489ce10c9ccc2d4edc6357cf2396df38c447de6e7d0506`. The isolated reference copy recorded 34 necessary package-name text patches; no fixture, prompt, package, or oracle was patched between subjects. This is a compatibility result for the one bounded section-partition transaction, not a general PPTX authoring claim.

`pptx-closed-leaf-slide-clone` is the complementary bounded graph slice. Its generated source contains a named `Release decision` SlidePart with exactly one relationship-free literal-data ChartPart, one top-level OLE frame bound to one closed uniquely owned XLSX package plus preview ImagePart, one canonical NotesSlide leaf, one legacy SlideComments leaf, and one relationship-free run action targeting a native custom show whose members are the source and visible `Appendix canary`. The Agent must invoke the shipped duplicate workflow with `--allow-closed-leaves`; it may add only the adjacent clone SlidePart and relationship part, one distinct ChartPart, one distinct byte-identical XLSX EmbeddedPackagePart, one NotesSlide plus relationship part, and one SlideComments part. The oracle independently reads OPC relationships, content types, custom-show membership, run actions, OLE frame bindings, inbound package ownership, and raw part hashes. It requires chart and OLE package relationships to keep their slide-local IDs/types, requires distinct closed clone-local mutable parts with byte-identical payloads, and requires the OLE preview target to remain shared. It also requires the clone run to retain the exact custom-show native ID/return policy without a hyperlink or slide relationship, requires show membership to remain source plus Appendix without the clone, and checks byte-identical retained source parts, verbatim copied notes/comments XML, a clone NotesSlide back-reference to the clone, and shared NotesMaster and CommentAuthors catalog parts. Native LibreOffice/Poppler evidence must retain the source page, make the clone pixel-identical to it, and keep the appendix pixel-identical. Repository smoke executes the real public workflow and adversarially rejects ChartPart aliasing, OLE-package aliasing, and custom-show membership drift. The earlier autonomous `0e8824c` trial predates the OLE extension. The current-main candidate/reference repeat at `/tmp/officekit-promptbench-closed-leaf-matrix-fae5327-v1` passed all six trials at `100/100` (candidate `3/3`, reference `3/3`), with no timeout or matrix identity drift. It bound package `4220ca5f7801ef1a3a3a3ce05a463cd5ce2463e1de7a53d5578cecc41ebe542e`, input contract `17289d9f3877ffdf7302fc610c8dbf3190d58888774224dc8dd04607d9370c5d`, fixture `03a28ad9f8c5cbba5ca70d43b7f3bc5c79e2e7a0dc8c3565b25ce782d9246253`, and oracle `a7c86f6e426448cf5aa745ef3b81f00ee21211524365afdb5a82584d2330e174`; candidate/reference Skill hashes were `c4f1dbf3df7cd1fa819a06cca726aba81821d4232148bf31934343b8409b1122` and `fa9f6badd10fc4957b489ce10c9ccc2d4edc6357cf2396df38c447de6e7d0506`, with 34 reference package-name patches. The protected authorization document was outside the npm tarball and is recorded only as source-tree status. This proves this bounded closed-leaf clone transaction, not general relationship-graph authoring.

One clean autonomous candidate Agent trial at `0e8824cb3dac8332ff631b0cb75850ebe2a56f6b` scored `100/100`: all hard gates plus 6 machine, 2 visual, 2 security, and 7 trace checks passed. The Agent selected the published workflow, supplied the required closed-leaf opt-in, and did not write scratch OOXML. The run bound tarball SHA-256 `b9759b4ef9712958c335712166887a2dd96d5b4eaa6b1b5fba1127b10b1fcac6`, copied-Skill SHA-256 `971096a78fec3c01c178cae76d6a458712a1b40df344a724cadde3b3acc6a9ff`, prompt SHA-256 `2bd75b885f0d4fa0b36e3395b82bcde17fe559088c54f9b9c78a54d3bdc908ea`, input SHA-256 `fc3b4ce9bbfbf101cea046e30c9c8a7550e9165771874061536ed131f15ef63e`, and oracle SHA-256 `0e8a68981b16824d464e86bacf3744bab45593ef451812b458e36923e9f4225d`. The source worktree was clean. This is historical one-trial evidence only.

The strict same-source matrix now closes the extended OLE case. At clean `0027dfe5fefc803549ed7e1300eeda67fa6d9511`, all six runs shared input SHA-256 `cfb63e44ee7272bdbb061b23ff37e0970923a3c3fd8b059e01cad196ad30b218`, input-contract SHA-256 `17289d9f3877ffdf7302fc610c8dbf3190d58888774224dc8dd04607d9370c5d`, candidate tarball SHA-256 `0d599353ec5f87655fdd9e28e052bc47d4dbb83919f33c975e76e25321ac4a62`, and oracle SHA-256 `3ba5677ede743c145e3ce466debda7404b10539816a338bcfc18dfe00079ac64`. Candidate and copied reference Skills each passed `3/3` at `100/100`, with identical hard gates and category checks (8 machine, 2 visual, 2 security, 7 trace). Their installed-tree hashes were `1b03b731f0e79e7d239c4c5d1b8092de204b03e458ec1842c611f6c0319856dc` and `fa9f6badd10fc4957b489ce10c9ccc2d4edc6357cf2396df38c447de6e7d0506`; the isolated reference copy recorded 34 necessary package-name text patches. No fixture, prompt, package, or oracle was patched between subjects. This is a compatibility result for the bounded closed-leaf clone transaction, not a general relationship-graph authoring claim.

The ready `pptx-smartart-notes-comments-boundary` case locks
`evals/assets/presentations/strategy-review.pptx`, a self-authored four-slide
deck whose first slide owns a closed four-part SmartArt graph and whose fourth
slide owns a NotesSlide plus an Office 2021 modern comment root/direct-reply
graph. The independent raw OPC inspector proves the third SmartArt node,
all four diagram parts, the page-four notes text, and the reply topology before
grading. The requested three-way mutation is intentionally outside one
source-bound transaction, so the accepted result is a `failed_closed` audit
with `save: none`, no PPTX or flattened image, explicit
`PresentationFile.importPptx`/`PresentationFile.inspectPptx` shell inspection,
and a top-level diagnostic naming SmartArt, speaker notes, comment reply, and
the source-bound/fail-closed boundary. The audit must use the canonical
`provider.actual`/`provider.silentFallback`, `savePolicy.strategy`, source
hash, `validation.sourceUnchanged`/`noArtifact`, and three unexecuted
operations fields; aliases or prose-only explanations do not pass. The grader
rejects partial output, SmartArt flattening, untyped XML writes, and silent
fallback. Fixture SHA-256 is
`bcb469d5b586f4fd8f562b918c8d9f04ef500cd6289728683c10ee2ced7be367`. The
isolated snapshot matrix based on `d020cf0643411646cfc470f113dab59749356d8a`
used package `5857b4fe284dcb2778a2f5420eecfb130d7352b3e58b2fc95854fb99ed3f4891`,
input contract `97e88636595f5cb7e75aa716395aa19e7cf6b962c6f83a2bef2259e8a4b69c0c`,
and oracle `ff38a67ab2d424df453ba82cb07f9f25e88d4e7ec280e2f66c0b8694577bd4ba`.
The OfficeKit Skill passed candidate `3/3` at `100/100`; the copied reference
Skill passed `1/3` at `100/100` and scored `86.67` in its two other completed
trials, with no timeout. Candidate/reference Skill hashes were
`7950017fe932fe6276ae5221cc6f2ca87b44519eb5438b60981dfd47d7588f70` and
`fa9f6badd10fc4957b489ce10c9ccc2d4edc6357cf2396df38c447de6e7d0506`.
The reference misses were limited to its diagnostic wording and do not waive
the candidate contract. This is a safe-refusal and compatibility result, not
general SmartArt authoring.

The ready `docx-complex-table-topology-boundary` case locks
`evals/assets/documents/clinical-form.docx`, a self-authored DOCX whose second
table combines a custom table style, vertical merge restart/continuation,
one nested table, one tracked cell, and one table-cell SDT; the first ordinary
table is an immutable canary. The independent raw OPC oracle requires every
one of those topology signals and the exact `Medication`/`Dose`/`Route`/
`Status` header before grading the Agent's response. Because the requested
column insertion would require a topology-changing imported edit that the
bounded model cannot prove safe, the only accepted result is a source-bound
`failed_closed` refusal with no modified DOCX and `save: none`. It rejects
flattening, scratch XML rebuilds, and silent fallback. Fixture SHA-256 is
`6d40a614deab54eb67e4f8cd73cb141e323ddf4b947f34795308921e177e09cd`; the
independent safe-refusal oracle is exercised in repository smoke.

## Pilot findings

The first generated PDF pilots produced two useful product signals:

1. The overflow replacement correctly failed closed, preserved the source hash, and emitted no modified PDF.
2. The first bounded equal-length contract-ID replacement was semantically and visually correct, including five-page count and unchanged non-target pages. However, the shipped `pymupdf_edit.py replace_text` primitive rejected the replacement because its fit calculation exceeded the box by roughly `0.00002pt`; the Agent then bypassed the typed primitive with a direct content-stream replacement. The grader scores that historical result 90/100: machine, visual, and security pass, while save-policy, typed-primitive, and low-level-bypass trace checks fail. It is therefore not a passing task.
3. The typed primitive now preserves the source baseline/default Base14 style and tolerates only provider/search-box float quantization, capped at `0.0005pt`; genuine overflow and rotated/cross-span input still fail closed. A second trial exposed an underspecified audit envelope, so the Skill now publishes `office-kit.pdf-audit.v1`, a JSON Schema, and `pdf_audit.py validate` to bind canonical provider/save-policy/operation evidence to exact source/output bytes.
4. With those changes fixed, the later clean six-trial matrix passed candidate and reference subjects `3/3` at `100/100`. The reference Skill discovered and used the published candidate PDF tasks/scripts through the same byte-identical candidate tarball. The matrix used the explicit private managed PyMuPDF route recorded above; it supports package discoverability and reference-workflow compatibility, not candidate-Skill superiority or general PDF reflow.
5. The active-content case exposed a provider-contract gap: PyMuPDF's stock scrub left root actions, invisible text, a comment, and a populated form value in the generated file. The typed adapter now removes the bounded inert-publication surface, fails when invisible text overlaps visible text, and runs a structural `--require-inert` gate after the full rewrite. A first live trial safely refused because the Skill hard-coded system `python3` instead of the runner-configured provider interpreter. After making interpreter selection explicit, two fixed-candidate trials passed 100/100. A third repeat exposed a narrower gap: PyMuPDF represented removed `/OpenAction`, `/AA`, `/JavaScript`, and `/EmbeddedFiles` names as `null`; the Agent repaired the final bytes through direct object mutation, so the trace hard gate correctly reduced a 90 raw score to zero. The official scrub primitive now physically removes those null active-content names, reports their xrefs, and fails closed on unfamiliar object serialization. The repeat matrix must be restarted on the fixed tarball; the historical bypass is not counted as a pass.
6. The first two restarted candidate trials then safely refused because both Agents invoked system `python3` despite the non-empty provider-runtime variable and documented command expansion. This was not accepted as environmental noise: the scripts now make runtime selection executable. Every shipped Python entry point re-executes through `OFFICE_KIT_PDF_PROVIDER_PYTHON` before dependency probing/import, and an invalid configured interpreter fails closed. This keeps probe, plan, mutation, scan, and audit on one provider identity even when an Agent writes `python3 script.py`; the matrix must restart again on that fixed package.
7. The final fixed active-content matrix passed candidate `3/3` and reference Skill `3/3`, every run at 100/100 across machine, visual, security, and trace. All six records bind clean commit `39fa301dcb1005f2848282e6e63da1e934104821`, package SHA-256 `e78e18c0f8f1cffe301ae1f2ea17e882bc879b3044914033e24b0b11ac0e8b69`, prompt/input/oracle fingerprints, candidate Skill SHA-256 `8ff51b8081babac4ec5af6ba1a8e4ae5b1df6a4a783dc51691fd580fa236fe3e`, and distinct reference Skill SHA-256 `4a32786820fbadf1d6c528002555597cc3aac200ac96e17040056eb51846b79b`. Reference trial 3 initially received a false trace failure because it viewed `edit --help` before mutation and independently scheduled probe/plan so plan finished first; the evaluator now ignores help-only invocations and requires both successful preflights, in either order, before the real edit. The retained trace then scores 100 without weakening bypass or post-mutation gates.
8. The first AcroForm audit found that string-valued pypdf radio input updated the parent `/V` while every widget `/AS` remained `/Off`; metadata looked filled but Poppler showed no selected radio. The typed adapter now resolves radio/checkbox inputs against real `/AP /N` names, writes PDF Name values, validates `/V`/`/AS` and appearances before promoting the transaction, and fails closed for unknown states or non-fillable fields. The fixture now really includes five text widgets, two radio widgets, and a pre-checked checkbox. The fixed candidate Skill passes `3/3` at 100/100 across machine, visual, security, and trace; the reference Skill passes `2/3` at 100/100. Its third run produced structurally and visually correct bytes but wrote PDF objects manually, declared `manual-incremental-acroform-writer`, omitted the canonical output hash, and never ran provider check/plan, typed fill, Poppler-through-workflow, or audit validation. The 70 raw machine/visual score is therefore forced to zero by the safety hard gates and is retained as comparative evidence rather than rerun away.
9. All six AcroForm records bind clean commit `bffd35dbfdb94bb1183717703e7e55bfb83c3f3c`, package SHA-256 `7ab9e6a30035df5d0ef7ee9990f3a0445152877e58a7f0d065ede9ddc1db300b`, prompt SHA-256 `a73276f4109cf69dbfa2681c2ba1ad270d425a06b7e214a547e341fcc7c09d08`, input SHA-256 `0ebb7bca3cd52a185ec6b68fe5b0acf52de86718af6d02322714d827caf6424e`, and oracle SHA-256 `6c725d8f14ea0de33b858f2eec5d171c7aef2ad5a03a8b1569822b7af885f3e4`. Candidate Skill SHA-256 is `14e6dcf2dc7f827d3285ca0896fc217fe04227a8c67ecd98d393703f75f2c0f8`; reference Skill SHA-256 is `4a32786820fbadf1d6c528002555597cc3aac200ac96e17040056eb51846b79b`. The reference PDF Skill contains no private package-name occurrence, so its isolated patch list is empty; the installed npm tarball and prompt remain byte-identical between subjects.
10. The attachment-quarantine vertical slice adds a typed, read-only pypdf primitive for document-level and page-level FileSpecs. The fixture contains six payloads: duplicate display names across scopes, Unicode, an archive, an executable-like payload, and `../escape.exe`. The primitive neutralizes traversal/portable reserved names, resolves case-insensitive collisions without dropping bytes, enforces decoded count/per-file/total budgets, verifies every SHA-256, binds the immutable source and manifest, and never opens or executes payload contents. The hardened audit contract directly repeats source preservation, all-hash, contained-path, duplicate-separation, and no-open/no-execute evidence. The current candidate and copied-reference Skills each pass `3/3` at 100/100 in the latest repeat matrix.
11. The remaining reference trial correctly extracted all six payloads, kept them inside quarantine, preserved the source, and did not execute attachments, but used a custom Node parser instead of the typed pypdf route. Its alternate manifest omitted the canonical schema/source contract, and its audit did not bind the manifest output hash or report pypdf, `read-only`, no-fallback, provider preflight, typed extraction, or canonical validation. Every applicable category therefore failed and the run scored zero. It remains comparative workflow-discipline evidence rather than being rerun away. All six fixed records bind clean commit `748fbb1d81ccfa14a594d6fed9bc6601866bfa95`, package SHA-256 `9cff93494c5b32e16394ce3b4fcffa1daf76ad6df57326dab0ced47d2a45b5bf`, prompt SHA-256 `944a2c8045012ebe07f85d4ec6d549f44eb1d9cd050ab4f713d31446a707a3e5`, input SHA-256 `7e02892e252150d257ce41660f0c303a647792a57ba2c0ba22fe5988bc79feb4`, and oracle SHA-256 `b3bd43ea180310ab327f03d255e6715a085710c02c2dce5a58bd963fd92bbaa3`. Candidate Skill SHA-256 is `c6642317d808571e25c986d0ddf9778431bdbcf13321c2133cafcfcc2e4b9d73`; reference Skill SHA-256 remains `4a32786820fbadf1d6c528002555597cc3aac200ac96e17040056eb51846b79b`.
12. The greenfield accessibility vertical slice adds first-class running-text artifacts, meaningful URI Link annotations with StructParent/OBJR tagging, and a constrained `semanticId` that merges consecutive physical table segments into one logical Table while preserving per-page paint content. The shipped six-page CJK example performs modeled verification, file inspection, Poppler rendering, canonical byte-bound audit, optional veraPDF probing, and explicit human PDF/UA review. The clean fixed matrix passes candidate `3/3` and reference Skill `3/3`, every run at 100/100 across 9 machine, 4 visual, 5 security, and 7 trace checks. All six records bind clean commit `2323a70331b93781dee37aa05198e4a73a7ec533`, byte-identical package SHA-256 `cfbcf5c76ba5fdb929dae27f2a0295d6da12694eec2150a926cebeecedefccb9`, prompt SHA-256 `074d1793ded062b25ab6b0c8a8a2bf07c7422f7ac127f58e986eabbec463cf2b`, input SHA-256 `ab0b2047d986b373f471c081d765f47d482a604d126b3784e4d9fed9374ed4a6`, and oracle SHA-256 `fe1bc0483cedad572e7983ce640c82449ad140d479dc4c798a68cc3c8bef4ec4`; candidate Skill SHA-256 is `e29be9d186bf22d2990fffc905bbe802fe4c5f338f3820da54e33e53be681815` and reference Skill SHA-256 is `4a32786820fbadf1d6c528002555597cc3aac200ac96e17040056eb51846b79b`. The earlier dirty-candidate pair remains discovery evidence, not part of the fixed matrix.
13. The merge/reorder/selective-watermark slice adds a bounded pypdf manifest primitive that selects every source page exactly once, preserves page boxes plus outlines/named destinations/internal links, and binds all inputs in one audit. The independent oracle verifies exact page mapping, 20% watermark placement, one-revision structure, source provenance, and Poppler pixels. A first fixed attempt exposed Agent thumbnail hallucination: the output pages were over 98.7% white and byte-identical to successful renders, but the Agent called them black and deleted a correct PDF. The shipped `poppler_compare.py` now emits typed per-page stability, changed-bounds, blank-state, and dark-ratio evidence. The clean fixed matrix passes candidate `3/3` and reference Skill `3/3`, every run at 100/100 across 8 machine, 3 visual, 6 security, and 9 trace checks. All six records bind commit `90cbb9e0a5527a4620a28bb38aad8feeca895a3b`, package SHA-256 `c3962993aee732c7a8e60282159409dfe08ca2c7c4f0dd59eb80468c28630ff5`, prompt SHA-256 `243eb766f2354b7f7b1c288843baf659f7f2ff0e2aa4529e9dc660a8f2b0f52d`, oracle SHA-256 `ef0b1496c95bc7c8ef3207dfb5f5596037fa0dc3d29a251202165a38981fe89c`, and identical cover/report/appendix hashes `79a3d12b2e92f013fe4b626093d702c3e1719e1a40ed2b0bae9eaf4b99342bfe`, `9545062fdbfc4ec8c474c7a034af52fb448d5c0d1869028bcaa3a11a28c6477f`, and `a5f60d4957dcb49d14a5676e461bdc333e9111f8d0efe48ae5c6aa5c565c00b7`. Candidate Skill SHA-256 is `9b38a3e358478301a48d6147eb29e39fd6c0c7a3da9473bd289f16db22a203d6`; reference Skill SHA-256 is `4a32786820fbadf1d6c528002555597cc3aac200ac96e17040056eb51846b79b`.

An earlier pilot missed the exact output filename because the runner prompt omitted declared inputs/deliverables. `PROMPT.md` now includes both, so that historical naming failure is evaluator noise rather than a product defect.

Command-trace grading detects completed inline shell commands and shipped primitive invocations; it is not operating-system syscall attestation. A production benchmark that must resist a deliberately evasive Agent should add container-level process/filesystem tracing in addition to the current no-network mount boundary.
