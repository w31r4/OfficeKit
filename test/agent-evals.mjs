import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import { DocumentFile, FileBlob, PresentationFile, SpreadsheetFile } from "../src/index.mjs";
import {
  DOCX_CLASSIC_COMMENT_FIXTURE,
  DOCX_FOOTER_TEXT_FIXTURE,
  DOCX_HEADER_TEXT_FIXTURE,
  DOCX_SECTION_PAGE_NUMBERING_FIXTURE,
  PPTX_CLOSED_LEAF_CLONE_FIXTURE,
  PPTX_RICH_NOTES_FIXTURE,
  PPTX_SECTION_BOUNDARY_FIXTURE,
  PPTX_SLIDE_NAME_FIXTURE,
  PPTX_TITLE_NOTES_FIXTURE,
  XLSX_CONNECTION_REFRESH_FIXTURE,
  XLSX_GROWTH_UPDATE_FIXTURE,
  XLSX_PIVOT_REFRESH_FIXTURE,
  XLSX_THREADED_REVIEW_FIXTURE,
  generateOfficeInput,
} from "../scripts/agent-eval-office-fixtures.mjs";
import {
  gradeDocxClassicCommentEvidence,
  gradeDocxFooterTextEvidence,
  gradeDocxHeaderTextEvidence,
  gradeDocxSectionPageNumberingEvidence,
  inspectClassicCommentDocx,
  inspectFooterTextDocx,
  inspectHeaderTextDocx,
  inspectSectionPageNumberingDocx,
} from "../scripts/agent-eval-docx-graders.mjs";
import { gradeOfficeCase } from "../scripts/agent-eval-office-graders.mjs";
import {
  gradeXlsxConnectionRefreshEvidence,
  gradeXlsxGrowthUpdateEvidence,
  gradeXlsxPivotRefreshEvidence,
  gradeXlsxThreadedReplyEvidence,
  inspectConnectionRefreshWorkbook,
  inspectGrowthWorkbook,
  inspectPivotRefreshWorkbook,
  inspectThreadedWorkbook,
} from "../scripts/agent-eval-spreadsheet-graders.mjs";
import {
  gradePptxClosedLeafCloneEvidence,
  gradePptxRichNotesEvidence,
  gradePptxSectionBoundaryEvidence,
  gradePptxSlideNameEvidence,
  inspectClosedLeafClonePptx,
  inspectRichNotesPptx,
  inspectSectionBoundaryPptx,
  inspectTitleNotesPptx,
} from "../scripts/agent-eval-presentation-graders.mjs";
import { duplicatePptxSlide } from "../skills/presentations/skills/presentations/examples/officekit-slide-duplicate-workflow.mjs";
import { replacePptxSectionPartition } from "../skills/presentations/skills/presentations/examples/officekit-section-boundary-edit-workflow.mjs";
import { editImportedHeaderText } from "../skills/documents/skills/documents/examples/officekit-header-text-edit-workflow.mjs";
import { editImportedFooterText } from "../skills/documents/skills/documents/examples/officekit-footer-text-edit-workflow.mjs";
import { editImportedSectionPageNumbering } from "../skills/documents/skills/documents/examples/officekit-section-page-numbering-edit-workflow.mjs";
import { hardenXlsxConnectionRefreshOnOpen } from "../skills/spreadsheets/skills/spreadsheets/examples/officekit-connection-refresh-hardening-workflow.mjs";
import { hardenXlsxPivotRefreshOnLoad } from "../skills/spreadsheets/skills/spreadsheets/examples/officekit-pivot-refresh-hardening-workflow.mjs";
import {
  extractCompletedCommands,
  gradeAcroFormEvidence,
  gradeAccessibleReportEvidence,
  gradeActiveContentSanitizeEvidence,
  gradeAttachmentQuarantineEvidence,
  gradeBoundaryRefusalEvidence,
  gradeBoundedReplaceEvidence,
  gradeCertifiedDocMdpP2FillEvidence,
  gradeDamagedXrefRecoveryEvidence,
  gradeMergeStampEvidence,
  gradeOverflowRefusalEvidence,
  gradeSourceBoundHighlightEvidence,
  summarizeCaseScore,
} from "../scripts/agent-eval-pdf-graders.mjs";
import {
  fingerprintPath,
  loadSuite,
  makeReadOnly,
  oracleFingerprint,
  providerRuntimeInstruction,
  removePreparedTree,
  repositoryProvenance,
  scorePrepared,
  skillSource,
  MINIMUM_PDF_CASE_SHARE,
  validateSuite,
  verifiedLockedAsset,
  visibleCase,
} from "../scripts/run-agent-evals.mjs";

const { suite, cases } = await loadSuite();
const repoRoot = path.resolve(import.meta.dirname, "..");
assert.deepEqual(validateSuite(suite, cases), { cases: 41, pdfCases: 21, ready: 29 });
assert.equal(MINIMUM_PDF_CASE_SHARE, 0.5);
const escapedAssetCases = structuredClone(cases);
escapedAssetCases.find((item) => item.id === "pdf-encrypted-owner-policy-boundary").inputs[0].asset = "../outside.pdf";
assert.throws(() => validateSuite(suite, escapedAssetCases), /input\.asset escapes the workspace/);
assert.equal(cases.filter((item) => item.family === "pdf" && item.status === "ready").length, 17);
assert.equal(cases.filter((item) => item.family === "spreadsheets" && item.status === "ready").length, 4);
assert.equal(cases.filter((item) => item.family === "documents" && item.status === "ready").length, 4);
assert.equal(cases.filter((item) => item.family === "presentations" && item.status === "ready").length, 4);
const referenceDocumentSkill = skillSource({ family: "documents", skill: "documents" }, "reference");
assert.equal(referenceDocumentSkill, path.join(repoRoot, "reference", "office-artifact-tool", "skills", "documents", "skills", "documents"));
assert.doesNotMatch(referenceDocumentSkill, /handoff/);
assert.equal((await fs.stat(path.join(referenceDocumentSkill, "SKILL.md"))).isFile(), true);

const repository = repositoryProvenance();
assert.match(repository.head, /^[0-9a-f]{40}$/);
assert.equal(typeof repository.dirty, "boolean");
assert.match(repository.statusSha256, /^[0-9a-f]{64}$/);
assert.match(repository.trackedDiffSha256, /^[0-9a-f]{64}$/);

const visible = visibleCase(suite, cases.find((item) => item.id === "pdf-bounded-contract-id-replace"));
assert.match(visible.prompt, /outputs\/contract-updated\.pdf/);
assert.match(visible.prompt, /outputs\/audit\.json/);
assert.doesNotMatch(visible.prompt, /expectedOutcome|oracleSha256|pymupdf\.readthedocs|"grade"/i);
const runnerHelp = spawnSync(process.execPath, ["scripts/run-agent-evals.mjs", "help"], {
  cwd: path.resolve(import.meta.dirname, ".."),
  encoding: "utf8",
});
assert.equal(runnerHelp.status, 0, runnerHelp.stderr);
assert.match(runnerHelp.stdout, /four PPTX cases.*section-boundary edit.*closed-leaf slide clone/i);
assert.match(runnerHelp.stdout, /connection refresh-on-open/i);
assert.match(runnerHelp.stdout, /pivot refresh-on-open/i);
assert.match(runnerHelp.stdout, /source-bound DOCX header text/i);
assert.match(runnerHelp.stdout, /source-bound DOCX footer text/i);
assert.match(runnerHelp.stdout, /17 ready PDF cases include nine locked corpus signature\/boundary\/repair fixtures/i);
assert.match(runnerHelp.stdout, /remaining 12 asset-required cases/i);
const highlightVisible = visibleCase(suite, cases.find((item) => item.id === "pdf-source-bound-text-highlight"));
assert.match(highlightVisible.prompt, /add_text_highlight/);
assert.match(highlightVisible.prompt, /outputs\/review-highlighted\.pdf/);
assert.doesNotMatch(highlightVisible.prompt, /oracleSha256|outputHighlights|changedWithinAllowedMask/i);
const mixedScanVisible = visibleCase(suite, cases.find((item) => item.id === "pdf-mixed-scan-ocr-boundary"));
assert.match(mixedScanVisible.prompt, /倒置.*倾斜|倾斜.*倒置/);
assert.match(mixedScanVisible.prompt, /born-digital selectable-text canaries/i);
assert.doesNotMatch(mixedScanVisible.prompt, /scanImageCanariesMatch|imageOnlyPages|pixelEncodedDeskewDegrees/i);
const damagedXrefVisible = visibleCase(suite, cases.find((item) => item.id === "pdf-damaged-xref-recovery"));
assert.match(damagedXrefVisible.prompt, /qpdf repair|qpdf.*repair/i);
assert.match(damagedXrefVisible.prompt, /inputs\/recoverable\.pdf/);
assert.match(damagedXrefVisible.prompt, /inputs\/unrecoverable\.pdf/);
assert.doesNotMatch(damagedXrefVisible.prompt, /f3847151df00468cce51c49a00cec7b5c35c24421160e310cbde0e90148b5550|sourceRawStartxrefIsZero/i);

const corpusPython = process.env.OFFICE_KIT_AGENT_EVAL_PYTHON || process.env.OFFICE_KIT_PDF_PROVIDER_PYTHON || "python3";
const corpusVerification = spawnSync(corpusPython, ["scripts/agent-eval-corpus-fixtures.py", "verify"], {
  cwd: repoRoot,
  encoding: "utf8",
  env: process.env,
});
const corpusRuntimeAvailable = corpusVerification.status === 0;
if (!corpusRuntimeAvailable) {
  assert.match(corpusVerification.stderr, /(?:No module named|ModuleNotFoundError|cannot import name)/i);
  console.log("PromptBench corpus structural oracle smoke skipped (set OFFICE_KIT_AGENT_EVAL_PYTHON to the managed Python runtime)");
}
const lockedFixturePaths = [
  "pdf/encryption/owner-policy-aes256.pdf",
  "pdf/encryption/user-password.json",
  "pdf/annotations/reply-chain.pdf",
  "pdf/accessibility/untagged-complex-report.pdf",
  "pdf/ocr/mixed-bilingual-scan.pdf",
  "pdf/corrupt/recoverable.pdf",
  "pdf/corrupt/unrecoverable.pdf",
  "pdf/xfa/dynamic-dependents.pdf",
  "pdf/print/print-production-risk.pdf",
  "pdf/signing/docmdp-p1-final.pdf",
  "pdf/signing/test-pki/root.pem",
  "pdf/signing/docmdp-p2-form.pdf",
  "pdf/signing/test-pki/docmdp-p2-root.pem",
];
for (const relative of lockedFixturePaths) {
  assert.equal(await verifiedLockedAsset(relative), path.join(repoRoot, "evals", "assets", relative));
}
const ownerCredential = await fs.readFile(path.join(repoRoot, "evals", "assets", "pdf", "encryption", "user-password.json"), "utf8");
assert.match(ownerCredential, /fixture-user-password/);
assert.doesNotMatch(ownerCredential, /fixture-owner-password-not-for-agent/);
const encryptionVisible = visibleCase(suite, cases.find((item) => item.id === "pdf-encrypted-owner-policy-boundary"));
assert.match(encryptionVisible.prompt, /inputs\/credentials\/user-password\.json/);
assert.doesNotMatch(encryptionVisible.prompt, /fixture-owner-password-not-for-agent/);
const docmdpP2Item = cases.find((item) => item.id === "pdf-docmdp-allowed-field-fill");
assert.equal(docmdpP2Item?.status, "ready");
assert.equal(docmdpP2Item?.inputs?.some((input) => input.asset === "pdf/signing/test-pki/docmdp-p2-root.pem"), true);

if (corpusRuntimeAvailable) {
assert.deepEqual(JSON.parse(corpusVerification.stdout), { assets: 13, ok: true, root: path.join(repoRoot, "evals", "assets") });
function boundaryOracle(boundary, source, userPassword) {
  const result = spawnSync(corpusPython, ["scripts/agent-eval-pdf-oracle.py"], {
    cwd: repoRoot,
    encoding: "utf8",
    env: process.env,
    input: JSON.stringify({ kind: "boundary-refusal", boundary, source, userPassword }),
  });
  assert.equal(result.status, 0, result.stderr);
  return JSON.parse(result.stdout);
}

function docmdpRefusalAudit(sourceHash, { includeProof = true } = {}) {
  const audit = {
    status: "failed_closed",
    source: { sha256: sourceHash },
    output: null,
    provider: { actual: "pyhanko", version: "0.35.2", silentFallback: false },
    savePolicy: { strategy: "read-only", executed: false, sourcePreserved: true, artifactWritten: false },
    operation: { type: "replace-title-preserve-certification-signature", executed: false },
    validation: { no_partial_modified_pdf_in_outputs: true, outputDirectoryPolicy: { modifiedPdfPresent: false, partialArtifactPresent: false } },
    delivered_modified_pdf: false,
    reason: "DocMDP P=1 prohibits title modification",
  };
  if (includeProof) {
    audit.signaturePolicy = {
      certificationSignaturePresent: true,
      docMDP: { permissionCode: 1, requestedChangeAllowed: false },
      policyDecision: "refuse_without_mutation",
    };
    audit.validation.signatureVerification = {
      conclusion: "valid-under-selected-policy",
      cryptographicallyValid: true,
      intact: true,
      trusted: true,
      trustPolicy: "explicit-roots",
      trustRoot: { sha256: "a".repeat(64) },
      bottomLine: true,
      coverage: "entire-file",
      docMDPCompliantBeforeRequestedChange: true,
    };
  }
  return audit;
}

const boundaryCases = [
  { id: "pdf-encrypted-owner-policy-boundary", boundary: "encrypted-owner-policy", source: path.join(repoRoot, "evals", "assets", "pdf", "encryption", "owner-policy-aes256.pdf"), userPassword: "fixture-user-password", diagnostic: "owner password is unavailable" },
  { id: "pdf-annotation-reply-resolve-boundary", boundary: "annotation-reply-chain", source: path.join(repoRoot, "evals", "assets", "pdf", "annotations", "reply-chain.pdf"), diagnostic: "reply chain and resolved state are not supported" },
  { id: "pdf-auto-pdfua-overclaim-refusal", boundary: "pdfua-overclaim", source: path.join(repoRoot, "evals", "assets", "pdf", "accessibility", "untagged-complex-report.pdf"), diagnostic: "human author intent is required" },
  { id: "pdf-mixed-scan-ocr-boundary", boundary: "mixed-scan-ocr", source: path.join(repoRoot, "evals", "assets", "pdf", "ocr", "mixed-bilingual-scan.pdf"), diagnostic: "deskew and orientation preprocessing are not in the published OCR profile" },
  { id: "pdf-dynamic-xfa-boundary", boundary: "dynamic-xfa", source: path.join(repoRoot, "evals", "assets", "pdf", "xfa", "dynamic-dependents.pdf"), diagnostic: "Dynamic XFA runtime is not supported" },
  { id: "pdf-print-production-boundary", boundary: "print-production", source: path.join(repoRoot, "evals", "assets", "pdf", "print", "print-production-risk.pdf"), diagnostic: "DeviceN and overprint need preflight" },
  { id: "pdf-docmdp-forbidden-title-edit", boundary: "docmdp-p1", source: path.join(repoRoot, "evals", "assets", "pdf", "signing", "docmdp-p1-final.pdf"), diagnostic: "DocMDP P=1 prohibits title modification", provider: "pyhanko" },
];
for (const boundary of boundaryCases) {
  const item = cases.find((candidate) => candidate.id === boundary.id);
  const evidence = boundaryOracle(boundary.boundary, boundary.source, boundary.userPassword);
  const audit = boundary.id === "pdf-docmdp-forbidden-title-edit" ? docmdpRefusalAudit(evidence.source.sha256) : {
    status: "failed_closed",
    source: { sha256: evidence.source.sha256 },
    output: null,
    provider: { actual: boundary.provider || "qpdf", version: boundary.provider ? "0.35.2" : "12.3.2-oat.2", silentFallback: false },
    savePolicy: { strategy: "none", executed: false, sourcePreserved: true, artifactPublished: false },
    operation: { type: "inspect-and-refuse-unsupported-boundary", executed: false },
    validation: { no_partial_modified_pdf_in_outputs: true, outputDirectoryPolicy: { modifiedPdfPresent: false, partialArtifactPresent: false } },
    delivered_modified_pdf: false,
    reason: boundary.diagnostic,
  };
  const checks = gradeBoundaryRefusalEvidence({
    evidence,
    audit,
    commands: boundary.id === "pdf-docmdp-forbidden-title-edit"
      ? ["python .agents/skills/pdf/scripts/pyhanko_provider.py verify inputs/source.pdf --expected-sha256 deadbeef --trust-policy explicit-roots --trust-root inputs/credentials/test-root.pem --require-signature --require-all-integrity-valid --require-all-trusted --require-docmdp-compliant --require-all-bottom-line"]
      : ["python .agents/skills/pdf/scripts/pdf_provider.py check --provider qpdf"],
    finalMessage: boundary.diagnostic,
    item,
    outputEntries: ["audit.json"],
  });
  assert.equal(checks.every((check) => check.passed), true, boundary.id);
  assert.equal(summarizeCaseScore(checks, item.grade).rawScorePercent, 100, boundary.id);
}
const docmdpBoundary = boundaryCases.find((boundary) => boundary.id === "pdf-docmdp-forbidden-title-edit");
const docmdpItem = cases.find((item) => item.id === docmdpBoundary.id);
const docmdpEvidence = boundaryOracle(docmdpBoundary.boundary, docmdpBoundary.source);
const docmdpIncompleteValidation = gradeBoundaryRefusalEvidence({
  evidence: docmdpEvidence,
  audit: docmdpRefusalAudit(docmdpEvidence.source.sha256),
  commands: ["python .agents/skills/pdf/scripts/pyhanko_provider.py verify inputs/source.pdf --trust-policy explicit-roots --trust-root inputs/credentials/test-root.pem --require-signature --require-all-integrity-valid"],
  finalMessage: docmdpBoundary.diagnostic,
  item: docmdpItem,
  outputEntries: ["audit.json"],
});
assert.equal(docmdpIncompleteValidation.find((check) => check.id === "pdf-security:docmdp-explicit-root-validation")?.passed, false, "DocMDP refusal requires trust, integrity, policy, and bottom-line validation");
const docmdpUnboundAudit = gradeBoundaryRefusalEvidence({
  evidence: docmdpEvidence,
  audit: docmdpRefusalAudit(docmdpEvidence.source.sha256, { includeProof: false }),
  commands: ["python .agents/skills/pdf/scripts/pyhanko_provider.py verify inputs/source.pdf --expected-sha256 deadbeef --trust-policy explicit-roots --trust-root inputs/credentials/test-root.pem --require-signature --require-all-integrity-valid --require-all-trusted --require-docmdp-compliant --require-all-bottom-line"],
  finalMessage: docmdpBoundary.diagnostic,
  item: docmdpItem,
  outputEntries: ["audit.json"],
});
assert.equal(docmdpUnboundAudit.find((check) => check.id === "pdf-security:docmdp-explicit-root-validation")?.passed, false, "DocMDP refusal requires audit-bound verification evidence, not only a command trace");
const encryptedBoundary = boundaryCases.find((boundary) => boundary.id === "pdf-encrypted-owner-policy-boundary");
const encryptedItem = cases.find((item) => item.id === encryptedBoundary.id);
const encryptedEvidence = boundaryOracle(encryptedBoundary.boundary, encryptedBoundary.source, encryptedBoundary.userPassword);
const candidateRefusalAudit = {
  status: "failed_closed",
  source: { sha256: encryptedEvidence.source.sha256 },
  output: null,
  provider: { actual: "mupdf", version: "1.28.0", silentFallback: false },
  savePolicy: { strategy: "read-only", artifactPolicy: "no modified artifact on failed-closed refusal" },
  operation: { type: "change-copy-permission", mutationAttempted: false },
  validation: { sourceIdentity: { sourcePreserved: true }, artifactChecks: { modifiedPdfPresent: false, partialArtifactPresent: false } },
  reason: "owner password is unavailable",
};
const candidateRefusalChecks = gradeBoundaryRefusalEvidence({
  evidence: encryptedEvidence,
  audit: candidateRefusalAudit,
  commands: ["node_modules/.bin/officekit run .agents/skills/pdf/scripts/mupdf.mjs probe", "node_modules/.bin/officekit run .agents/skills/pdf/scripts/mupdf.mjs inspect inputs/owner-policy-aes256.pdf"],
  finalMessage: "owner password is unavailable",
  item: encryptedItem,
  outputEntries: ["audit.json"],
});
assert.equal(candidateRefusalChecks.every((check) => check.passed), true, "candidate-style refusal audit");
const proseOnlyRefusalAudit = {
  ...candidateRefusalAudit,
  operation: { type: "change-copy-permission", performed: "authorization reviewed" },
  validation: { sourceIdentity: { sourcePreserved: true } },
};
const proseOnlyRefusalChecks = gradeBoundaryRefusalEvidence({
  evidence: encryptedEvidence,
  audit: proseOnlyRefusalAudit,
  commands: ["node_modules/.bin/officekit run .agents/skills/pdf/scripts/mupdf.mjs probe"],
  finalMessage: "owner password is unavailable",
  item: encryptedItem,
  outputEntries: ["audit.json"],
});
assert.equal(proseOnlyRefusalChecks.find((check) => check.id === "pdf-security:no-partial-output")?.passed, false, "prose does not prove an absent output artifact");
const referenceRefusalAudit = {
  status: "failed_closed",
  source: { sha256: encryptedEvidence.source.sha256, preserved: true },
  provider: { actualProvider: "mupdf-js", actualProviderVersion: "1.28.0", fallbackUsed: false },
  saveStrategy: { selected: "none", publication: "refused" },
  operations: [
    { operation: "sha256", mode: "read_only", result: "passed" },
    { operation: "permission_change", result: "not_attempted" },
  ],
  verification: { modifiedArtifactPresent: false },
  reason: "owner password is unavailable",
};
const referenceRefusalChecks = gradeBoundaryRefusalEvidence({
  evidence: encryptedEvidence,
  audit: referenceRefusalAudit,
  commands: ["from pypdf import PdfReader"],
  finalMessage: "owner password is unavailable",
  item: encryptedItem,
  outputEntries: ["audit.json"],
});
assert.equal(referenceRefusalChecks.every((check) => check.passed), true, "reference-style refusal audit");
const partialOutputChecks = gradeBoundaryRefusalEvidence({
  evidence: encryptedEvidence,
  audit: candidateRefusalAudit,
  commands: ["python .agents/skills/pdf/scripts/pdf_provider.py check --provider qpdf"],
  finalMessage: "owner password is unavailable",
  item: encryptedItem,
  outputEntries: ["audit.json", "partial.pdf"],
});
assert.equal(partialOutputChecks.find((check) => check.id === "pdf-security:no-partial-output")?.passed, false);
const xfaBoundary = boundaryCases.find((boundary) => boundary.id === "pdf-dynamic-xfa-boundary");
const xfaItem = cases.find((candidate) => candidate.id === xfaBoundary.id);
const xfaEvidence = boundaryOracle(xfaBoundary.boundary, xfaBoundary.source);
xfaEvidence.hasXfaPackets = false;
const xfaFixtureFailure = gradeBoundaryRefusalEvidence({
  evidence: xfaEvidence,
  audit: {
    status: "failed_closed",
    source: { sha256: xfaEvidence.source.sha256 },
    output: null,
    provider: { actual: "qpdf", version: "12.3.2-oat.2", silentFallback: false },
    savePolicy: { strategy: "none", executed: false, sourcePreserved: true, artifactPublished: false },
    operation: { type: "inspect-and-refuse-unsupported-boundary", executed: false },
    validation: { no_partial_modified_pdf_in_outputs: true, outputDirectoryPolicy: { modifiedPdfPresent: false, partialArtifactPresent: false } },
    delivered_modified_pdf: false,
    reason: "Dynamic XFA runtime is not supported",
  },
  commands: ["python .agents/skills/pdf/scripts/pdf_provider.py check --provider qpdf"],
  finalMessage: "Dynamic XFA runtime is not supported",
  item: xfaItem,
  outputEntries: ["audit.json"],
});
assert.equal(xfaFixtureFailure.find((check) => check.id === "pdf-machine:source-boundary-fixture")?.passed, false);
const mixedScanBoundary = boundaryCases.find((boundary) => boundary.id === "pdf-mixed-scan-ocr-boundary");
const mixedScanItem = cases.find((candidate) => candidate.id === mixedScanBoundary.id);
const mixedScanEvidence = boundaryOracle(mixedScanBoundary.boundary, mixedScanBoundary.source);
mixedScanEvidence.scanImageCanariesMatch = false;
const mixedScanFixtureFailure = gradeBoundaryRefusalEvidence({
  evidence: mixedScanEvidence,
  audit: {
    status: "failed_closed",
    source: { sha256: mixedScanEvidence.source.sha256 },
    output: null,
    provider: { actual: "ocrmypdf", version: "17.8.1", silentFallback: false },
    savePolicy: { strategy: "none", executed: false, sourcePreserved: true, artifactPublished: false },
    operation: { type: "ocr-with-automatic-deskew", executed: false },
    validation: { no_partial_modified_pdf_in_outputs: true, outputDirectoryPolicy: { modifiedPdfPresent: false, partialArtifactPresent: false } },
    delivered_modified_pdf: false,
    reason: "deskew and orientation preprocessing are not in the published OCR profile",
  },
  commands: ["python .agents/skills/pdf/scripts/pdf_provider.py check --provider ocrmypdf"],
  finalMessage: "deskew and orientation preprocessing are not in the published OCR profile",
  item: mixedScanItem,
  outputEntries: ["audit.json"],
});
assert.equal(mixedScanFixtureFailure.find((check) => check.id === "pdf-machine:source-boundary-fixture")?.passed, false);

const damagedXrefItem = cases.find((item) => item.id === "pdf-damaged-xref-recovery");
const recoveryQpdf = process.env.OFFICE_KIT_AGENT_EVAL_QPDF || process.env.OFFICE_KIT_PDF_QPDF || "qpdf";
const recoveryPoppler = process.env.OFFICE_KIT_AGENT_EVAL_PDFTOPPM || "pdftoppm";
const recoveryQpdfAvailable = spawnSync(recoveryQpdf, ["--version"], { encoding: "utf8", env: process.env }).status === 0;
const recoveryPopplerAvailable = spawnSync(recoveryPoppler, ["-v"], { encoding: "utf8", env: process.env }).status === 0;
if (recoveryQpdfAvailable && recoveryPopplerAvailable) {
  const recoveryRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-agent-eval-damaged-xref-"));
  try {
    const recoverable = path.join(recoveryRoot, "inputs", "recoverable.pdf");
    const unrecoverable = path.join(recoveryRoot, "inputs", "unrecoverable.pdf");
    const recovered = path.join(recoveryRoot, "outputs", "recovered.pdf");
    await fs.mkdir(path.dirname(recoverable), { recursive: true });
    await fs.mkdir(path.dirname(recovered), { recursive: true });
    await fs.copyFile(path.join(repoRoot, "evals", "assets", "pdf", "corrupt", "recoverable.pdf"), recoverable);
    await fs.copyFile(path.join(repoRoot, "evals", "assets", "pdf", "corrupt", "unrecoverable.pdf"), unrecoverable);
    const qpdfProvider = path.join(repoRoot, "skills", "pdf", "skills", "pdf", "scripts", "qpdf_provider.py");
    const providerEnvironment = { ...process.env, OFFICE_KIT_PDF_QPDF: recoveryQpdf, PYTHONDONTWRITEBYTECODE: "1" };
    const inspect = spawnSync(corpusPython, [qpdfProvider, "inspect", recoverable], { cwd: repoRoot, encoding: "utf8", env: providerEnvironment });
    assert.equal(inspect.status, 0, inspect.stderr);
    const inspectAudit = JSON.parse(inspect.stdout);
    assert.equal(inspectAudit.check.status, "warnings");
    const rewrite = spawnSync(corpusPython, [qpdfProvider, "rewrite", recoverable, recovered, "--mode", "repair", "--expected-sha256", inspectAudit.source.sha256], { cwd: repoRoot, encoding: "utf8", env: providerEnvironment });
    assert.equal(rewrite.status, 0, rewrite.stderr);
    const repairAudit = JSON.parse(rewrite.stdout);
    const rejected = spawnSync(corpusPython, [qpdfProvider, "inspect", unrecoverable], { cwd: repoRoot, encoding: "utf8", env: providerEnvironment });
    assert.equal(rejected.status, 2, rejected.stdout || rejected.stderr);
    const recoverableHash = crypto.createHash("sha256").update(await fs.readFile(recoverable)).digest("hex");
    const unrecoverableHash = crypto.createHash("sha256").update(await fs.readFile(unrecoverable)).digest("hex");
    const audit = {
      ...repairAudit,
      status: "succeeded",
      controls: {
        unrecoverable: {
          status: "rejected",
          source: { sha256: unrecoverableHash },
          output: null,
          artifactWritten: false,
          diagnostics: String(rejected.stderr || rejected.stdout).split(/\r?\n/).filter(Boolean),
        },
      },
      validation: { sourceImmutable: true, render: { renderer: "pdftoppm", allPages: true } },
    };
    const oracle = spawnSync(corpusPython, ["scripts/agent-eval-pdf-oracle.py"], {
      cwd: repoRoot,
      encoding: "utf8",
      env: providerEnvironment,
      input: JSON.stringify({ kind: "damaged-xref-recovery", recoverable, unrecoverable, output: recovered, renderRoot: path.join(recoveryRoot, "evaluator", "render"), poppler: recoveryPoppler }),
    });
    assert.equal(oracle.status, 0, oracle.stderr);
    const evidence = JSON.parse(oracle.stdout);
    const commands = [
      "python .agents/skills/pdf/scripts/qpdf_provider.py probe",
      "python .agents/skills/pdf/scripts/qpdf_provider.py inspect inputs/recoverable.pdf > tmp/recoverable-inspect.json",
      "python .agents/skills/pdf/scripts/qpdf_provider.py inspect inputs/unrecoverable.pdf > tmp/unrecoverable-inspect.json || true",
      `python .agents/skills/pdf/scripts/qpdf_provider.py rewrite inputs/recoverable.pdf outputs/recovered.pdf --mode repair --expected-sha256 ${recoverableHash} > tmp/rewrite.json`,
      "python .agents/skills/pdf/scripts/qpdf_provider.py inspect outputs/recovered.pdf > tmp/recovered-inspect.json",
      "pdftoppm -png -r 144 outputs/recovered.pdf outputs/render/recovered",
    ];
    const checks = gradeDamagedXrefRecoveryEvidence({ evidence, audit, commands, item: damagedXrefItem });
    assert.equal(checks.every((check) => check.passed), true, JSON.stringify(checks.filter((check) => !check.passed), null, 2));
    assert.equal(summarizeCaseScore(checks, damagedXrefItem.grade).rawScorePercent, 100);
    const attachmentLoss = structuredClone(evidence);
    attachmentLoss.outputAttachments = [];
    const attachmentLossChecks = gradeDamagedXrefRecoveryEvidence({ evidence: attachmentLoss, audit, commands, item: damagedXrefItem });
    assert.equal(attachmentLossChecks.find((check) => check.id === "pdf-machine:document-attachment-preserved")?.passed, false, "a repaired page stream without its document attachment must fail");
  } finally {
    await fs.rm(recoveryRoot, { recursive: true, force: true });
  }
} else {
  console.log("PromptBench damaged-XRef qpdf execution smoke skipped (provide qpdf and pdftoppm)");
}

const p2ProviderPython = process.env.OFFICE_KIT_PYHANKO_TEST_PYTHON;
const p2Poppler = process.env.OFFICE_KIT_AGENT_EVAL_PDFTOPPM || "pdftoppm";
const p2PopplerAvailable = spawnSync(p2Poppler, ["-v"], { encoding: "utf8", env: process.env }).status === 0;
if (p2ProviderPython && p2PopplerAvailable) {
  const p2Root = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-agent-eval-docmdp-p2-"));
  try {
    const source = path.join(p2Root, "inputs", "source.pdf");
    const trustRoot = path.join(p2Root, "inputs", "credentials", "test-root.pem");
    const output = path.join(p2Root, "outputs", "approved-amount.pdf");
    const auditPath = path.join(p2Root, "outputs", "audit.json");
    await fs.mkdir(path.dirname(trustRoot), { recursive: true });
    await fs.mkdir(path.dirname(output), { recursive: true });
    await fs.copyFile(path.join(repoRoot, "evals", "assets", "pdf", "signing", "docmdp-p2-form.pdf"), source);
    await fs.copyFile(path.join(repoRoot, "evals", "assets", "pdf", "signing", "test-pki", "docmdp-p2-root.pem"), trustRoot);
    const sourceHash = crypto.createHash("sha256").update(await fs.readFile(source)).digest("hex");
    const filler = path.join(repoRoot, "skills", "pdf", "skills", "pdf", "scripts", "pyhanko_certified_form_fill.py");
    const filled = spawnSync(p2ProviderPython, [
      filler, "fill", source, output,
      "--expected-source-sha256", sourceHash,
      "--trust-root", trustRoot,
      "--field", "ApprovedAmount", "--value", "12500.00",
      "--expected-signature-field", "Certification",
      "--expected-locked-field", "LockedAmount", "--expected-locked-value", "LOCKED-9000",
      "--caller-isolated",
    ], {
      cwd: repoRoot,
      encoding: "utf8",
      env: { ...process.env, OFFICE_KIT_PDF_PROVIDER_PYTHON: p2ProviderPython, PYTHONDONTWRITEBYTECODE: "1" },
    });
    assert.equal(filled.status, 0, filled.stderr);
    const p2Audit = JSON.parse(filled.stdout);
    await fs.writeFile(auditPath, JSON.stringify(p2Audit, null, 2));
    const p2Oracle = spawnSync(corpusPython, ["scripts/agent-eval-pdf-oracle.py"], {
      cwd: repoRoot,
      encoding: "utf8",
      env: process.env,
      input: JSON.stringify({
        kind: "certified-docmdp-p2-fill",
        source,
        output,
        field: "ApprovedAmount",
        value: "12500.00",
        lockedField: "LockedAmount",
        renderRoot: path.join(p2Root, "evaluator", "render"),
        poppler: p2Poppler,
      }),
    });
    assert.equal(p2Oracle.status, 0, p2Oracle.stderr);
    const p2Evidence = JSON.parse(p2Oracle.stdout);
    const p2Commands = [
      "python .agents/skills/pdf/scripts/pyhanko_certified_form_fill.py probe",
      "python .agents/skills/pdf/scripts/pyhanko_certified_form_fill.py fill inputs/source.pdf outputs/approved-amount.pdf --expected-source-sha256 deadbeef --trust-root inputs/credentials/test-root.pem --field ApprovedAmount --value 12500.00 --expected-signature-field Certification --expected-locked-field LockedAmount --expected-locked-value LOCKED-9000 --caller-isolated",
      "python .agents/skills/pdf/scripts/pyhanko_provider.py verify outputs/approved-amount.pdf --trust-policy explicit-roots --trust-root inputs/credentials/test-root.pem --require-signature --require-all-integrity-valid --require-all-trusted --require-docmdp-compliant --require-all-bottom-line",
      "node .agents/skills/pdf/scripts/mupdf.mjs render outputs/approved-amount.pdf outputs/render",
    ];
    const p2Checks = gradeCertifiedDocMdpP2FillEvidence({ evidence: p2Evidence, audit: p2Audit, commands: p2Commands, item: docmdpP2Item });
    assert.equal(p2Checks.every((check) => check.passed), true, "DocMDP P=2 PromptBench fixture must pass every independent check");
    assert.equal(summarizeCaseScore(p2Checks, docmdpP2Item.grade).rawScorePercent, 100);
    const p2TamperedAudit = structuredClone(p2Audit);
    p2TamperedAudit.signature.postflight.changedFormFields = ["LockedAmount"];
    const p2TamperedChecks = gradeCertifiedDocMdpP2FillEvidence({ evidence: p2Evidence, audit: p2TamperedAudit, commands: p2Commands, item: docmdpP2Item });
    assert.equal(p2TamperedChecks.find((check) => check.id === "pdf-security:audit-provenance-and-policy")?.passed, false, "a generic success audit cannot replace an exact DocMDP/FieldMDP proof");
  } finally {
    await fs.rm(p2Root, { recursive: true, force: true });
  }
} else {
  console.log("PromptBench DocMDP P=2 execution smoke skipped (set OFFICE_KIT_PYHANKO_TEST_PYTHON and provide pdftoppm)");
}
}

const threadedReplyItem = cases.find((item) => item.id === "xlsx-threaded-reply-resolve");
const threadedReplyRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-eval-xlsx-threaded-"));
try {
  const threadedInput = path.join(threadedReplyRoot, "inputs", XLSX_THREADED_REVIEW_FIXTURE.workbookName);
  const threadedOutput = path.join(threadedReplyRoot, "outputs", "reviewed-budget-resolved.xlsx");
  await generateOfficeInput("xlsx-threaded-review", threadedInput);
  const threadedSource = await fs.readFile(threadedInput);
  const threadedWorkbook = await SpreadsheetFile.importXlsx(new FileBlob(threadedSource, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", name: XLSX_THREADED_REVIEW_FIXTURE.workbookName }));
  const threadedReview = threadedWorkbook.comments.threads.find((thread) => thread.target.sheetName === XLSX_THREADED_REVIEW_FIXTURE.sheetName && thread.target.address === XLSX_THREADED_REVIEW_FIXTURE.address);
  assert.ok(threadedReview);
  threadedReview.addReply(XLSX_THREADED_REVIEW_FIXTURE.requestedReply, { author: "Board secretary", date: "2026-07-17T10:00:00.000Z" });
  threadedReview.resolve();
  const threadedExport = await SpreadsheetFile.exportXlsx(threadedWorkbook, { recalculate: false });
  const threadedBytes = new Uint8Array(await threadedExport.arrayBuffer());
  await fs.mkdir(path.dirname(threadedOutput), { recursive: true });
  await fs.writeFile(threadedOutput, threadedBytes);
  const hash = (bytes) => crypto.createHash("sha256").update(bytes).digest("hex");
  const threadedAudit = {
    status: "succeeded",
    source: { sha256: hash(threadedSource) },
    output: { sha256: hash(threadedBytes) },
    provider: { actual: "office-kit", version: "test", silentFallback: false },
    savePolicy: { strategy: "rewrite" },
    operation: { type: "threaded-comment-direct-reply-resolve" },
    validation: { reimport: { ok: true } },
  };
  await fs.writeFile(path.join(threadedReplyRoot, "outputs", "audit.json"), JSON.stringify(threadedAudit, null, 2));
  const threadedEvidence = {
    source: await inspectThreadedWorkbook(threadedInput),
    output: await inspectThreadedWorkbook(threadedOutput),
    visual: {
      source: { available: true, ok: true, pageCount: 1, pages: [{ nonWhitePixels: 1 }] },
      output: { available: true, ok: true, pageCount: 1, pages: [{ nonWhitePixels: 1 }] },
    },
  };
  const threadedTrace = JSON.stringify({ type: "item.completed", item: { type: "command_execution", id: "xlsx-threaded", command: "node -e 'SpreadsheetFile.importXlsx(); SpreadsheetFile.exportXlsx()'" } });
  const threadedChecks = gradeXlsxThreadedReplyEvidence({ evidence: threadedEvidence, audit: threadedAudit, commands: extractCompletedCommands(threadedTrace), item: threadedReplyItem });
  assert.equal(threadedChecks.every((check) => check.passed), true);
  const publishedWorkflowChecks = gradeXlsxThreadedReplyEvidence({
    evidence: threadedEvidence,
    audit: threadedAudit,
    commands: ["node .agents/skills/spreadsheets/examples/officekit-threaded-comment-reply-workflow.mjs inputs/reviewed-budget.xlsx outputs/reviewed-budget-resolved.xlsx outputs/audit.json"],
    item: threadedReplyItem,
  });
  assert.equal(publishedWorkflowChecks.find((check) => check.id === "xlsx-trace:typed-roundtrip")?.passed, true);
  const untrustedWorkflowChecks = gradeXlsxThreadedReplyEvidence({
    evidence: threadedEvidence,
    audit: threadedAudit,
    commands: ["node scratch/threaded-comment-reply-workflow.mjs inputs/reviewed-budget.xlsx outputs/reviewed-budget-resolved.xlsx outputs/audit.json"],
    item: threadedReplyItem,
  });
  assert.equal(untrustedWorkflowChecks.find((check) => check.id === "xlsx-trace:typed-roundtrip")?.passed, false);
  const nativeThreadedResult = await gradeOfficeCase({ item: threadedReplyItem, workspace: threadedReplyRoot, evaluator: path.join(threadedReplyRoot, "evaluator"), finalMessage: "completed", trace: threadedTrace });
  if (nativeThreadedResult.graded) {
    assert.equal(nativeThreadedResult.rawScorePercent, 100);
    assert.equal(nativeThreadedResult.caseSpecificPassed, true);
  } else {
    assert.ok(nativeThreadedResult.infrastructureErrors?.length);
  }
} finally {
  await fs.rm(threadedReplyRoot, { recursive: true, force: true });
}

const growthUpdateItem = cases.find((item) => item.id === "xlsx-growth-assumption-update");
assert.ok(growthUpdateItem);
const growthUpdateRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-eval-xlsx-growth-"));
try {
  const growthInput = path.join(growthUpdateRoot, "inputs", XLSX_GROWTH_UPDATE_FIXTURE.workbookName);
  const growthOutput = path.join(growthUpdateRoot, "outputs", "operating-plan-updated.xlsx");
  await generateOfficeInput("xlsx-growth-update", growthInput);
  const growthSource = await fs.readFile(growthInput);
  const growthWorkbook = await SpreadsheetFile.importXlsx(new FileBlob(growthSource, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    name: XLSX_GROWTH_UPDATE_FIXTURE.workbookName,
  }));
  const forecast = growthWorkbook.worksheets.getItem(XLSX_GROWTH_UPDATE_FIXTURE.targetSheetName);
  const baseline = growthWorkbook.worksheets.getItem(XLSX_GROWTH_UPDATE_FIXTURE.canarySheetName);
  assert.ok(forecast);
  assert.ok(baseline);
  const baselineSnapshot = structuredClone(baseline.getRange("A1:C5").values);
  assert.equal(forecast.getRange(XLSX_GROWTH_UPDATE_FIXTURE.growthAddress).values[0][0], XLSX_GROWTH_UPDATE_FIXTURE.originalGrowth);
  assert.equal(forecast.getRange(XLSX_GROWTH_UPDATE_FIXTURE.marginAddress).values[0][0], XLSX_GROWTH_UPDATE_FIXTURE.grossMargin);
  assert.deepEqual(forecast.getRange("B5:B7").formulas.flat(), XLSX_GROWTH_UPDATE_FIXTURE.revenueFormulas);
  forecast.getRange(XLSX_GROWTH_UPDATE_FIXTURE.growthAddress).values = [[XLSX_GROWTH_UPDATE_FIXTURE.replacementGrowth]];
  growthWorkbook.recalculate();
  assert.ok(forecast.getRange("B5:B7").values.flat().every((value, index) => Math.abs(value - XLSX_GROWTH_UPDATE_FIXTURE.revisedRevenue[index]) < 1e-7));
  assert.deepEqual(baseline.getRange("A1:C5").values, baselineSnapshot);
  const growthExport = await SpreadsheetFile.exportXlsx(growthWorkbook, { recalculate: false });
  const growthBytes = new Uint8Array(await growthExport.arrayBuffer());
  await fs.mkdir(path.dirname(growthOutput), { recursive: true });
  await fs.writeFile(growthOutput, growthBytes);
  const hash = (bytes) => crypto.createHash("sha256").update(bytes).digest("hex");
  const growthAudit = {
    status: "succeeded",
    source: { sha256: hash(growthSource) },
    output: { sha256: hash(growthBytes) },
    provider: { actual: "office-kit", version: "test", silentFallback: false },
    savePolicy: { strategy: "rewrite" },
    operation: { type: "growth-assumption-update" },
    validation: { reimport: { ok: true } },
  };
  await fs.writeFile(path.join(growthUpdateRoot, "outputs", "audit.json"), JSON.stringify(growthAudit, null, 2));
  const growthEvidence = {
    source: await inspectGrowthWorkbook(growthInput),
    output: await inspectGrowthWorkbook(growthOutput),
    visual: {
      source: { available: true, ok: true, pageCount: 3, pages: [
        { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "growth-source-page-1" },
        { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "growth-source-page-2" },
        { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "baseline-stable" },
      ] },
      output: { available: true, ok: true, pageCount: 3, pages: [
        { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "growth-output-page-1" },
        { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "growth-output-page-2" },
        { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "baseline-stable" },
      ] },
    },
  };
  const growthTrace = JSON.stringify({ type: "item.completed", item: { type: "command_execution", id: "xlsx-growth", command: "node -e 'SpreadsheetFile.importXlsx(); SpreadsheetFile.exportXlsx()'" } });
  const growthChecks = gradeXlsxGrowthUpdateEvidence({
    evidence: growthEvidence,
    audit: growthAudit,
    commands: extractCompletedCommands(growthTrace),
    item: growthUpdateItem,
  });
  assert.equal(growthChecks.every((check) => check.passed), true);
  const publishedGrowthWorkflowChecks = gradeXlsxGrowthUpdateEvidence({
    evidence: growthEvidence,
    audit: growthAudit,
    commands: ["node .agents/skills/spreadsheets/examples/officekit-growth-assumption-edit-workflow.mjs inputs/operating-plan.xlsx outputs/operating-plan-updated.xlsx outputs/audit.json"],
    item: growthUpdateItem,
  });
  assert.equal(publishedGrowthWorkflowChecks.find((check) => check.id === "xlsx-growth-trace:typed-roundtrip")?.passed, true);
  const untrustedGrowthWorkflowChecks = gradeXlsxGrowthUpdateEvidence({
    evidence: growthEvidence,
    audit: growthAudit,
    commands: ["node scratch/growth-update.mjs inputs/operating-plan.xlsx outputs/operating-plan-updated.xlsx outputs/audit.json"],
    item: growthUpdateItem,
  });
  assert.equal(untrustedGrowthWorkflowChecks.find((check) => check.id === "xlsx-growth-trace:typed-roundtrip")?.passed, false);
  const nativeGrowthResult = await gradeOfficeCase({
    item: growthUpdateItem,
    workspace: growthUpdateRoot,
    evaluator: path.join(growthUpdateRoot, "evaluator"),
    finalMessage: "completed",
    trace: growthTrace,
  });
  if (nativeGrowthResult.graded) {
    assert.equal(nativeGrowthResult.rawScorePercent, 100);
    assert.equal(nativeGrowthResult.caseSpecificPassed, true);
  } else {
    assert.ok(nativeGrowthResult.infrastructureErrors?.length);
  }
} finally {
  await fs.rm(growthUpdateRoot, { recursive: true, force: true });
}

const connectionRefreshItem = cases.find((item) => item.id === "xlsx-connection-refresh-on-open");
assert.ok(connectionRefreshItem);
const connectionRefreshRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-eval-xlsx-connection-refresh-"));
try {
  const connectionInput = path.join(connectionRefreshRoot, "inputs", XLSX_CONNECTION_REFRESH_FIXTURE.workbookName);
  const connectionOutput = path.join(connectionRefreshRoot, "outputs", "external-sales-refresh-disabled.xlsx");
  const connectionAuditPath = path.join(connectionRefreshRoot, "outputs", "audit.json");
  await generateOfficeInput("xlsx-connection-refresh", connectionInput);
  const connectionSourceBytes = await fs.readFile(connectionInput);
  const connectionResult = await hardenXlsxConnectionRefreshOnOpen({
    inputPath: connectionInput,
    outputPath: connectionOutput,
    auditPath: connectionAuditPath,
    connectionId: XLSX_CONNECTION_REFRESH_FIXTURE.connectionId,
    expectedName: XLSX_CONNECTION_REFRESH_FIXTURE.connectionName,
  });
  assert.equal(connectionResult.audit.validation.reimport.ok, true);
  assert.deepEqual(await fs.readFile(connectionInput), connectionSourceBytes);
  const connectionAudit = JSON.parse(await fs.readFile(connectionAuditPath, "utf8"));
  const connectionEvidence = {
    source: await inspectConnectionRefreshWorkbook(connectionInput),
    output: await inspectConnectionRefreshWorkbook(connectionOutput),
    visual: {
      source: { available: true, ok: true, pageCount: 1, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "connection-refresh-stable" }] },
      output: { available: true, ok: true, pageCount: 1, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "connection-refresh-stable" }] },
    },
  };
  const connectionTrace = JSON.stringify({
    type: "item.completed",
    item: {
      type: "command_execution",
      id: "xlsx-connection-refresh",
      command: "node .agents/skills/spreadsheets/examples/officekit-connection-refresh-hardening-workflow.mjs inputs/external-sales-refresh-on-open.xlsx outputs/external-sales-refresh-disabled.xlsx outputs/audit.json 7 'Fixture warehouse'",
    },
  });
  const connectionChecks = gradeXlsxConnectionRefreshEvidence({
    evidence: connectionEvidence,
    audit: connectionAudit,
    commands: extractCompletedCommands(connectionTrace),
    item: connectionRefreshItem,
  });
  assert.equal(connectionChecks.every((check) => check.passed), true);
  const publishedConnectionWorkflowChecks = gradeXlsxConnectionRefreshEvidence({
    evidence: connectionEvidence,
    audit: connectionAudit,
    commands: ["node .agents/skills/spreadsheets/examples/officekit-connection-refresh-hardening-workflow.mjs inputs/external-sales-refresh-on-open.xlsx outputs/external-sales-refresh-disabled.xlsx outputs/audit.json 7 'Fixture warehouse'"],
    item: connectionRefreshItem,
  });
  assert.equal(publishedConnectionWorkflowChecks.find((check) => check.id === "xlsx-connection-trace:typed-roundtrip")?.passed, true);
  const rootDriftChecks = gradeXlsxConnectionRefreshEvidence({
    evidence: {
      ...connectionEvidence,
      output: {
        ...connectionEvidence.output,
        partHashes: {
          ...connectionEvidence.output.partHashes,
          "xl/workbook.xml": "root-drift",
        },
      },
    },
    audit: connectionAudit,
    commands: extractCompletedCommands(connectionTrace),
    item: connectionRefreshItem,
  });
  assert.equal(rootDriftChecks.find((check) => check.id === "xlsx-connection-machine:only-connections-part-changed")?.passed, false);
  assert.equal(rootDriftChecks.find((check) => check.id === "xlsx-connection-security:package-scope-and-source-provenance")?.passed, false);
  const connectionResidualDriftChecks = gradeXlsxConnectionRefreshEvidence({
    evidence: {
      ...connectionEvidence,
      output: {
        ...connectionEvidence.output,
        connectionXml: connectionEvidence.output.connectionXml.replace('value="kept"', 'value="drifted"'),
      },
    },
    audit: connectionAudit,
    commands: extractCompletedCommands(connectionTrace),
    item: connectionRefreshItem,
  });
  assert.equal(connectionResidualDriftChecks.find((check) => check.id === "xlsx-connection-machine:connection-residual-stable")?.passed, false);
  assert.equal(connectionResidualDriftChecks.find((check) => check.id === "xlsx-connection-security:package-scope-and-source-provenance")?.passed, false);
  const untrustedConnectionWorkflowChecks = gradeXlsxConnectionRefreshEvidence({
    evidence: connectionEvidence,
    audit: connectionAudit,
    commands: ["node scratch/patch-connections-xml.mjs inputs/external-sales-refresh-on-open.xlsx outputs/external-sales-refresh-disabled.xlsx"],
    item: connectionRefreshItem,
  });
  assert.equal(untrustedConnectionWorkflowChecks.find((check) => check.id === "xlsx-connection-trace:typed-roundtrip")?.passed, false);
  const nativeConnectionResult = await gradeOfficeCase({
    item: connectionRefreshItem,
    workspace: connectionRefreshRoot,
    evaluator: path.join(connectionRefreshRoot, "evaluator"),
    finalMessage: "completed",
    trace: connectionTrace,
  });
  if (nativeConnectionResult.graded) {
    assert.equal(nativeConnectionResult.rawScorePercent, 100);
    assert.equal(nativeConnectionResult.caseSpecificPassed, true);
  } else {
    assert.ok(nativeConnectionResult.infrastructureErrors?.length);
  }
} finally {
  await fs.rm(connectionRefreshRoot, { recursive: true, force: true });
}

const pivotRefreshItem = cases.find((item) => item.id === "xlsx-pivot-refresh-on-open");
assert.ok(pivotRefreshItem);
const pivotRefreshRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-eval-xlsx-pivot-refresh-"));
try {
  const pivotInput = path.join(pivotRefreshRoot, "inputs", XLSX_PIVOT_REFRESH_FIXTURE.workbookName);
  const pivotOutput = path.join(pivotRefreshRoot, "outputs", "regional-revenue-refresh-disabled.xlsx");
  const pivotAuditPath = path.join(pivotRefreshRoot, "outputs", "audit.json");
  await generateOfficeInput("xlsx-pivot-refresh", pivotInput);
  const pivotSourceBytes = await fs.readFile(pivotInput);
  const pivotResult = await hardenXlsxPivotRefreshOnLoad({
    inputPath: pivotInput,
    outputPath: pivotOutput,
    auditPath: pivotAuditPath,
    worksheetName: XLSX_PIVOT_REFRESH_FIXTURE.sheetName,
    pivotName: XLSX_PIVOT_REFRESH_FIXTURE.pivotName,
  });
  assert.equal(pivotResult.audit.validation.reimport.ok, true);
  assert.deepEqual(await fs.readFile(pivotInput), pivotSourceBytes);
  const pivotAudit = JSON.parse(await fs.readFile(pivotAuditPath, "utf8"));
  const pivotEvidence = {
    source: await inspectPivotRefreshWorkbook(pivotInput),
    output: await inspectPivotRefreshWorkbook(pivotOutput),
    visual: {
      source: { available: true, ok: true, pageCount: 1, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "pivot-refresh-stable" }] },
      output: { available: true, ok: true, pageCount: 1, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "pivot-refresh-stable" }] },
    },
  };
  const pivotTrace = JSON.stringify({
    type: "item.completed",
    item: {
      type: "command_execution",
      id: "xlsx-pivot-refresh",
      command: "node .agents/skills/spreadsheets/examples/officekit-pivot-refresh-hardening-workflow.mjs inputs/regional-revenue-refresh-on-open.xlsx outputs/regional-revenue-refresh-disabled.xlsx outputs/audit.json 'Pivot Summary' 'Revenue by region'",
    },
  });
  const pivotChecks = gradeXlsxPivotRefreshEvidence({
    evidence: pivotEvidence,
    audit: pivotAudit,
    commands: extractCompletedCommands(pivotTrace),
    item: pivotRefreshItem,
  });
  assert.equal(pivotChecks.every((check) => check.passed), true);
  const publishedPivotWorkflowChecks = gradeXlsxPivotRefreshEvidence({
    evidence: pivotEvidence,
    audit: pivotAudit,
    commands: ["node .agents/skills/spreadsheets/examples/officekit-pivot-refresh-hardening-workflow.mjs inputs/regional-revenue-refresh-on-open.xlsx outputs/regional-revenue-refresh-disabled.xlsx outputs/audit.json 'Pivot Summary' 'Revenue by region'"],
    item: pivotRefreshItem,
  });
  assert.equal(publishedPivotWorkflowChecks.find((check) => check.id === "xlsx-pivot-trace:typed-roundtrip")?.passed, true);
  const rootDriftChecks = gradeXlsxPivotRefreshEvidence({
    evidence: {
      ...pivotEvidence,
      output: {
        ...pivotEvidence.output,
        partHashes: {
          ...pivotEvidence.output.partHashes,
          "xl/workbook.xml": "root-drift",
        },
      },
    },
    audit: pivotAudit,
    commands: extractCompletedCommands(pivotTrace),
    item: pivotRefreshItem,
  });
  assert.equal(rootDriftChecks.find((check) => check.id === "xlsx-pivot-machine:only-cache-definition-changed")?.passed, false);
  assert.equal(rootDriftChecks.find((check) => check.id === "xlsx-pivot-security:package-scope-and-source-provenance")?.passed, false);
  const residualDriftChecks = gradeXlsxPivotRefreshEvidence({
    evidence: {
      ...pivotEvidence,
      output: {
        ...pivotEvidence.output,
        cache: {
          ...pivotEvidence.output.cache,
          normalized: `${pivotEvidence.output.cache.normalized}fixture-drift`,
        },
      },
    },
    audit: pivotAudit,
    commands: extractCompletedCommands(pivotTrace),
    item: pivotRefreshItem,
  });
  assert.equal(residualDriftChecks.find((check) => check.id === "xlsx-pivot-machine:cache-residual-stable")?.passed, false);
  assert.equal(residualDriftChecks.find((check) => check.id === "xlsx-pivot-security:package-scope-and-source-provenance")?.passed, false);
  const untrustedPivotWorkflowChecks = gradeXlsxPivotRefreshEvidence({
    evidence: pivotEvidence,
    audit: pivotAudit,
    commands: ["node scratch/patch-pivot-cache-xml.mjs inputs/regional-revenue-refresh-on-open.xlsx outputs/regional-revenue-refresh-disabled.xlsx"],
    item: pivotRefreshItem,
  });
  assert.equal(untrustedPivotWorkflowChecks.find((check) => check.id === "xlsx-pivot-trace:typed-roundtrip")?.passed, false);
  const nativePivotResult = await gradeOfficeCase({
    item: pivotRefreshItem,
    workspace: pivotRefreshRoot,
    evaluator: path.join(pivotRefreshRoot, "evaluator"),
    finalMessage: "completed",
    trace: pivotTrace,
  });
  if (nativePivotResult.graded) {
    assert.equal(nativePivotResult.rawScorePercent, 100);
    assert.equal(nativePivotResult.caseSpecificPassed, true);
  } else {
    assert.ok(nativePivotResult.infrastructureErrors?.length);
  }
} finally {
  await fs.rm(pivotRefreshRoot, { recursive: true, force: true });
}

const classicCommentItem = cases.find((item) => item.id === "docx-classic-comment-text-edit");
const classicCommentRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-eval-docx-comment-"));
try {
  const classicInput = path.join(classicCommentRoot, "inputs", DOCX_CLASSIC_COMMENT_FIXTURE.documentName);
  const classicOutput = path.join(classicCommentRoot, "outputs", "legal-review-updated.docx");
  await generateOfficeInput("docx-classic-comment-review", classicInput);
  const classicSource = await fs.readFile(classicInput);
  const classicDocument = await DocumentFile.importDocx(new FileBlob(classicSource, {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    name: DOCX_CLASSIC_COMMENT_FIXTURE.documentName,
  }));
  const classicTarget = classicDocument.blocks.find((block) => block.kind === "paragraph" && block.text === DOCX_CLASSIC_COMMENT_FIXTURE.anchorText);
  assert.ok(classicTarget);
  const classicComment = classicDocument.comments.find((comment) => comment.targetId === classicTarget.id);
  assert.ok(classicComment);
  assert.equal(classicComment.text, DOCX_CLASSIC_COMMENT_FIXTURE.comment.originalText);
  classicComment.text = DOCX_CLASSIC_COMMENT_FIXTURE.comment.replacementText;
  const classicExport = await DocumentFile.exportDocx(classicDocument);
  const classicBytes = new Uint8Array(await classicExport.arrayBuffer());
  await fs.mkdir(path.dirname(classicOutput), { recursive: true });
  await fs.writeFile(classicOutput, classicBytes);
  const hash = (bytes) => crypto.createHash("sha256").update(bytes).digest("hex");
  const classicAudit = {
    status: "succeeded",
    source: { sha256: hash(classicSource) },
    output: { sha256: hash(classicBytes) },
    provider: { actual: "office-kit", version: "test", silentFallback: false },
    savePolicy: { strategy: "rewrite" },
    operation: { type: "classic-comment-text-edit" },
    validation: { reimport: { ok: true } },
  };
  await fs.writeFile(path.join(classicCommentRoot, "outputs", "audit.json"), JSON.stringify(classicAudit, null, 2));
  const classicEvidence = {
    source: await inspectClassicCommentDocx(classicInput),
    output: await inspectClassicCommentDocx(classicOutput),
    visual: {
      source: { available: true, ok: true, pageCount: 1, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "stable" }] },
      output: { available: true, ok: true, pageCount: 1, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "stable" }] },
    },
  };
  const classicTrace = JSON.stringify({ type: "item.completed", item: { type: "command_execution", id: "docx-comment", command: "node -e 'DocumentFile.importDocx(); DocumentFile.exportDocx()'" } });
  const classicChecks = gradeDocxClassicCommentEvidence({
    evidence: classicEvidence,
    audit: classicAudit,
    commands: extractCompletedCommands(classicTrace),
    item: classicCommentItem,
  });
  assert.equal(classicChecks.every((check) => check.passed), true);
  const publishedClassicWorkflowChecks = gradeDocxClassicCommentEvidence({
    evidence: classicEvidence,
    audit: classicAudit,
    commands: ["node .agents/skills/documents/examples/officekit-classic-comment-edit-workflow.mjs inputs/legal-review.docx outputs/legal-review-updated.docx outputs/audit.json"],
    item: classicCommentItem,
  });
  assert.equal(publishedClassicWorkflowChecks.find((check) => check.id === "docx-trace:typed-roundtrip")?.passed, true);
  const untrustedClassicWorkflowChecks = gradeDocxClassicCommentEvidence({
    evidence: classicEvidence,
    audit: classicAudit,
    commands: ["node scratch/classic-comment-edit.mjs inputs/legal-review.docx outputs/legal-review-updated.docx outputs/audit.json"],
    item: classicCommentItem,
  });
  assert.equal(untrustedClassicWorkflowChecks.find((check) => check.id === "docx-trace:typed-roundtrip")?.passed, false);
  const nativeClassicResult = await gradeOfficeCase({
    item: classicCommentItem,
    workspace: classicCommentRoot,
    evaluator: path.join(classicCommentRoot, "evaluator"),
    finalMessage: "completed",
    trace: classicTrace,
  });
  if (nativeClassicResult.graded) {
    assert.equal(nativeClassicResult.rawScorePercent, 100);
    assert.equal(nativeClassicResult.caseSpecificPassed, true);
  } else {
    assert.ok(nativeClassicResult.infrastructureErrors?.length);
  }
} finally {
  await fs.rm(classicCommentRoot, { recursive: true, force: true });
}

const headerTextItem = cases.find((item) => item.id === "docx-header-text-edit");
assert.ok(headerTextItem);
const headerTextRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-eval-docx-header-"));
try {
  const headerInput = path.join(headerTextRoot, "inputs", DOCX_HEADER_TEXT_FIXTURE.documentName);
  const headerOutput = path.join(headerTextRoot, "outputs", "board-brief-header-reviewed.docx");
  const headerAuditPath = path.join(headerTextRoot, "outputs", "audit.json");
  await generateOfficeInput("docx-header-text-review", headerInput);
  const headerSourceBytes = await fs.readFile(headerInput);
  const headerResult = await editImportedHeaderText({
    inputPath: headerInput,
    outputPath: headerOutput,
    auditPath: headerAuditPath,
    expectedText: DOCX_HEADER_TEXT_FIXTURE.header.originalText,
    replacementText: DOCX_HEADER_TEXT_FIXTURE.header.replacementText,
    sectionIndex: DOCX_HEADER_TEXT_FIXTURE.header.sectionIndex,
    referenceType: DOCX_HEADER_TEXT_FIXTURE.header.referenceType,
  });
  assert.deepEqual(await fs.readFile(headerInput), headerSourceBytes);
  assert.deepEqual(headerResult.audit.validation.changedParts, ["word/header1.xml"]);
  const headerAudit = JSON.parse(await fs.readFile(headerAuditPath, "utf8"));
  const headerEvidence = {
    source: await inspectHeaderTextDocx(headerInput),
    output: await inspectHeaderTextDocx(headerOutput),
    visual: {
      source: { available: true, ok: true, pageCount: 1, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "header-source" }] },
      output: { available: true, ok: true, pageCount: 1, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "header-output" }] },
    },
  };
  const headerTrace = JSON.stringify({
    type: "item.completed",
    item: {
      type: "command_execution",
      id: "docx-header-text",
      command: "node .agents/skills/documents/examples/officekit-header-text-edit-workflow.mjs inputs/board-brief-header.docx outputs/board-brief-header-reviewed.docx outputs/audit.json 'Northwind | Internal' 'Northwind | Reviewed' 0 default",
    },
  });
  const headerChecks = gradeDocxHeaderTextEvidence({
    evidence: headerEvidence,
    audit: headerAudit,
    commands: extractCompletedCommands(headerTrace),
    item: headerTextItem,
  });
  assert.equal(headerChecks.every((check) => check.passed), true);
  const publishedHeaderWorkflowChecks = gradeDocxHeaderTextEvidence({
    evidence: headerEvidence,
    audit: headerAudit,
    commands: ["node .agents/skills/documents/examples/officekit-header-text-edit-workflow.mjs inputs/board-brief-header.docx outputs/board-brief-header-reviewed.docx outputs/audit.json"],
    item: headerTextItem,
  });
  assert.equal(publishedHeaderWorkflowChecks.find((check) => check.id === "docx-header-trace:typed-roundtrip")?.passed, true);
  const untrustedHeaderWorkflowChecks = gradeDocxHeaderTextEvidence({
    evidence: headerEvidence,
    audit: headerAudit,
    commands: ["node scratch/patch-header-xml.mjs inputs/board-brief-header.docx outputs/board-brief-header-reviewed.docx"],
    item: headerTextItem,
  });
  assert.equal(untrustedHeaderWorkflowChecks.find((check) => check.id === "docx-header-trace:typed-roundtrip")?.passed, false);
  const packageDriftEvidence = structuredClone(headerEvidence);
  packageDriftEvidence.output.partHashes["word/footer1.xml"] = "unexpected-footer-drift";
  const packageDriftChecks = gradeDocxHeaderTextEvidence({
    evidence: packageDriftEvidence,
    audit: headerAudit,
    commands: extractCompletedCommands(headerTrace),
    item: headerTextItem,
  });
  assert.equal(packageDriftChecks.find((check) => check.id === "docx-header-security:only-target-header-part-changed")?.passed, false);
  assert.equal(packageDriftChecks.find((check) => check.id === "docx-header-security:footer-field-and-package-inventory-preserved")?.passed, false);
  const residualDriftEvidence = structuredClone(headerEvidence);
  residualDriftEvidence.output.headerParts[0].xml = residualDriftEvidence.output.headerParts[0].xml
    .replace(DOCX_HEADER_TEXT_FIXTURE.header.companionText, "Unexpected companion drift.");
  const residualDriftChecks = gradeDocxHeaderTextEvidence({
    evidence: residualDriftEvidence,
    audit: headerAudit,
    commands: extractCompletedCommands(headerTrace),
    item: headerTextItem,
  });
  assert.equal(residualDriftChecks.find((check) => check.id === "docx-header-machine:header-residual-stable")?.passed, false);
  const nativeHeaderResult = await gradeOfficeCase({
    item: headerTextItem,
    workspace: headerTextRoot,
    evaluator: path.join(headerTextRoot, "evaluator"),
    finalMessage: "completed",
    trace: headerTrace,
  });
  if (nativeHeaderResult.graded) {
    assert.equal(nativeHeaderResult.rawScorePercent, 100);
    assert.equal(nativeHeaderResult.caseSpecificPassed, true);
  } else {
    assert.ok(nativeHeaderResult.infrastructureErrors?.length);
  }
} finally {
  await fs.rm(headerTextRoot, { recursive: true, force: true });
}

const footerTextItem = cases.find((item) => item.id === "docx-footer-text-edit");
assert.ok(footerTextItem);
const footerTextRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-eval-docx-footer-"));
try {
  const footerInput = path.join(footerTextRoot, "inputs", DOCX_FOOTER_TEXT_FIXTURE.documentName);
  const footerOutput = path.join(footerTextRoot, "outputs", "board-brief-footer-reviewed.docx");
  const footerAuditPath = path.join(footerTextRoot, "outputs", "audit.json");
  await generateOfficeInput("docx-footer-text-review", footerInput);
  const footerSourceBytes = await fs.readFile(footerInput);
  const footerResult = await editImportedFooterText({
    inputPath: footerInput,
    outputPath: footerOutput,
    auditPath: footerAuditPath,
    expectedText: DOCX_FOOTER_TEXT_FIXTURE.footer.originalText,
    replacementText: DOCX_FOOTER_TEXT_FIXTURE.footer.replacementText,
    sectionIndex: DOCX_FOOTER_TEXT_FIXTURE.footer.sectionIndex,
    referenceType: DOCX_FOOTER_TEXT_FIXTURE.footer.referenceType,
  });
  assert.deepEqual(await fs.readFile(footerInput), footerSourceBytes);
  assert.deepEqual(footerResult.audit.validation.changedParts, ["word/footer1.xml"]);
  const footerAudit = JSON.parse(await fs.readFile(footerAuditPath, "utf8"));
  const footerEvidence = {
    source: await inspectFooterTextDocx(footerInput),
    output: await inspectFooterTextDocx(footerOutput),
    visual: {
      source: { available: true, ok: true, pageCount: 1, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "footer-source" }] },
      output: { available: true, ok: true, pageCount: 1, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "footer-output" }] },
    },
  };
  const footerTrace = JSON.stringify({
    type: "item.completed",
    item: {
      type: "command_execution",
      id: "docx-footer-text",
      command: "node .agents/skills/documents/examples/officekit-footer-text-edit-workflow.mjs inputs/board-brief-footer.docx outputs/board-brief-footer-reviewed.docx outputs/audit.json 'Northwind | Internal' 'Northwind | Reviewed' 0 default",
    },
  });
  const footerChecks = gradeDocxFooterTextEvidence({
    evidence: footerEvidence,
    audit: footerAudit,
    commands: extractCompletedCommands(footerTrace),
    item: footerTextItem,
  });
  assert.equal(footerChecks.every((check) => check.passed), true);
  const publishedFooterWorkflowChecks = gradeDocxFooterTextEvidence({
    evidence: footerEvidence,
    audit: footerAudit,
    commands: ["node .agents/skills/documents/examples/officekit-footer-text-edit-workflow.mjs inputs/board-brief-footer.docx outputs/board-brief-footer-reviewed.docx outputs/audit.json"],
    item: footerTextItem,
  });
  assert.equal(publishedFooterWorkflowChecks.find((check) => check.id === "docx-footer-trace:typed-roundtrip")?.passed, true);
  const untrustedFooterWorkflowChecks = gradeDocxFooterTextEvidence({
    evidence: footerEvidence,
    audit: footerAudit,
    commands: ["node scratch/patch-footer-xml.mjs inputs/board-brief-footer.docx outputs/board-brief-footer-reviewed.docx"],
    item: footerTextItem,
  });
  assert.equal(untrustedFooterWorkflowChecks.find((check) => check.id === "docx-footer-trace:typed-roundtrip")?.passed, false);
  const packageDriftEvidence = structuredClone(footerEvidence);
  packageDriftEvidence.output.partHashes["word/header1.xml"] = "unexpected-header-drift";
  const packageDriftChecks = gradeDocxFooterTextEvidence({
    evidence: packageDriftEvidence,
    audit: footerAudit,
    commands: extractCompletedCommands(footerTrace),
    item: footerTextItem,
  });
  assert.equal(packageDriftChecks.find((check) => check.id === "docx-footer-security:only-target-footer-part-changed")?.passed, false);
  assert.equal(packageDriftChecks.find((check) => check.id === "docx-footer-security:header-field-and-package-inventory-preserved")?.passed, false);
  const residualDriftEvidence = structuredClone(footerEvidence);
  residualDriftEvidence.output.footerParts[0].xml = residualDriftEvidence.output.footerParts[0].xml
    .replace(DOCX_FOOTER_TEXT_FIXTURE.footer.companionText, "Unexpected companion drift.");
  const residualDriftChecks = gradeDocxFooterTextEvidence({
    evidence: residualDriftEvidence,
    audit: footerAudit,
    commands: extractCompletedCommands(footerTrace),
    item: footerTextItem,
  });
  assert.equal(residualDriftChecks.find((check) => check.id === "docx-footer-machine:footer-residual-stable")?.passed, false);
  const nativeFooterResult = await gradeOfficeCase({
    item: footerTextItem,
    workspace: footerTextRoot,
    evaluator: path.join(footerTextRoot, "evaluator"),
    finalMessage: "completed",
    trace: footerTrace,
  });
  if (nativeFooterResult.graded) {
    assert.equal(nativeFooterResult.rawScorePercent, 100);
    assert.equal(nativeFooterResult.caseSpecificPassed, true);
  } else {
    assert.ok(nativeFooterResult.infrastructureErrors?.length);
  }
} finally {
  await fs.rm(footerTextRoot, { recursive: true, force: true });
}

const sectionPageNumberingItem = cases.find((item) => item.id === "docx-section-page-numbering-edit");
assert.ok(sectionPageNumberingItem);
const sectionPageNumberingRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-eval-docx-page-numbering-"));
try {
  const fixture = DOCX_SECTION_PAGE_NUMBERING_FIXTURE;
  const sectionPageNumberingInput = path.join(sectionPageNumberingRoot, "inputs", fixture.documentName);
  const sectionPageNumberingOutput = path.join(sectionPageNumberingRoot, "outputs", "front-matter-page-numbering-reviewed.docx");
  const sectionPageNumberingAuditPath = path.join(sectionPageNumberingRoot, "outputs", "audit.json");
  await generateOfficeInput("docx-section-page-numbering-review", sectionPageNumberingInput);
  const sectionPageNumberingSourceBytes = await fs.readFile(sectionPageNumberingInput);
  const sectionPageNumberingResult = await editImportedSectionPageNumbering({
    inputPath: sectionPageNumberingInput,
    outputPath: sectionPageNumberingOutput,
    auditPath: sectionPageNumberingAuditPath,
    sectionBlockIndex: fixture.target.blockIndex,
    expectedPageNumbering: fixture.target.originalPageNumbering,
    replacementPageNumbering: fixture.target.replacementPageNumbering,
  });
  assert.deepEqual(await fs.readFile(sectionPageNumberingInput), sectionPageNumberingSourceBytes);
  assert.deepEqual(sectionPageNumberingResult.audit.validation.changedParts, ["word/document.xml"]);
  await assert.rejects(
    () => editImportedSectionPageNumbering({
      inputPath: sectionPageNumberingInput,
      outputPath: path.join(sectionPageNumberingRoot, "outputs", "wrong-source-value.docx"),
      auditPath: path.join(sectionPageNumberingRoot, "outputs", "wrong-source-value-audit.json"),
      sectionBlockIndex: fixture.target.blockIndex,
      expectedPageNumbering: { start: 1, format: "upperRoman" },
      replacementPageNumbering: fixture.target.replacementPageNumbering,
    }),
    /does not match the expected source value/i,
  );
  const sectionPageNumberingAudit = JSON.parse(await fs.readFile(sectionPageNumberingAuditPath, "utf8"));
  const sectionPageNumberingEvidence = {
    source: await inspectSectionPageNumberingDocx(sectionPageNumberingInput),
    output: await inspectSectionPageNumberingDocx(sectionPageNumberingOutput),
    visual: {
      source: {
        available: true,
        ok: true,
        pageCount: 3,
        pages: [
          { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "page-numbering-source-target" },
          { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "page-numbering-stable-1" },
          { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "page-numbering-stable-2" },
        ],
      },
      output: {
        available: true,
        ok: true,
        pageCount: 3,
        pages: [
          { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "page-numbering-output-target" },
          { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "page-numbering-stable-1" },
          { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "page-numbering-stable-2" },
        ],
      },
    },
  };
  const sectionPageNumberingTrace = JSON.stringify({
    type: "item.completed",
    item: {
      type: "command_execution",
      id: "docx-section-page-numbering",
      command: "node .agents/skills/documents/examples/officekit-section-page-numbering-edit-workflow.mjs inputs/front-matter-page-numbering.docx outputs/front-matter-page-numbering-reviewed.docx outputs/audit.json 1 '{\"start\":1,\"format\":\"lowerRoman\"}' '{\"start\":1,\"format\":\"decimal\"}'",
    },
  });
  const sectionPageNumberingChecks = gradeDocxSectionPageNumberingEvidence({
    evidence: sectionPageNumberingEvidence,
    audit: sectionPageNumberingAudit,
    commands: extractCompletedCommands(sectionPageNumberingTrace),
    item: sectionPageNumberingItem,
  });
  assert.equal(sectionPageNumberingChecks.every((check) => check.passed), true);
  const publishedSectionPageNumberingWorkflowChecks = gradeDocxSectionPageNumberingEvidence({
    evidence: sectionPageNumberingEvidence,
    audit: sectionPageNumberingAudit,
    commands: ["node .agents/skills/documents/examples/officekit-section-page-numbering-edit-workflow.mjs inputs/front-matter-page-numbering.docx outputs/front-matter-page-numbering-reviewed.docx outputs/audit.json"],
    item: sectionPageNumberingItem,
  });
  assert.equal(publishedSectionPageNumberingWorkflowChecks.find((check) => check.id === "docx-page-numbering-trace:typed-roundtrip")?.passed, true);
  const untrustedSectionPageNumberingWorkflowChecks = gradeDocxSectionPageNumberingEvidence({
    evidence: sectionPageNumberingEvidence,
    audit: sectionPageNumberingAudit,
    commands: ["node scratch/patch-pg-num-type.mjs inputs/front-matter-page-numbering.docx outputs/front-matter-page-numbering-reviewed.docx"],
    item: sectionPageNumberingItem,
  });
  assert.equal(untrustedSectionPageNumberingWorkflowChecks.find((check) => check.id === "docx-page-numbering-trace:typed-roundtrip")?.passed, false);
  const packageDriftEvidence = structuredClone(sectionPageNumberingEvidence);
  packageDriftEvidence.output.partHashes["word/footer1.xml"] = "unexpected-footer-drift";
  const packageDriftChecks = gradeDocxSectionPageNumberingEvidence({
    evidence: packageDriftEvidence,
    audit: sectionPageNumberingAudit,
    commands: extractCompletedCommands(sectionPageNumberingTrace),
    item: sectionPageNumberingItem,
  });
  assert.equal(packageDriftChecks.find((check) => check.id === "docx-page-numbering-security:only-document-part-changed")?.passed, false);
  assert.equal(packageDriftChecks.find((check) => check.id === "docx-page-numbering-security:sibling-sections-footers-and-inventory-preserved")?.passed, false);
  const residualDriftEvidence = structuredClone(sectionPageNumberingEvidence);
  residualDriftEvidence.output.documentXml = residualDriftEvidence.output.documentXml.replace('w:fmt="upperLetter"', 'w:fmt="lowerLetter"');
  const residualDriftChecks = gradeDocxSectionPageNumberingEvidence({
    evidence: residualDriftEvidence,
    audit: sectionPageNumberingAudit,
    commands: extractCompletedCommands(sectionPageNumberingTrace),
    item: sectionPageNumberingItem,
  });
  assert.equal(residualDriftChecks.find((check) => check.id === "docx-page-numbering-machine:document-xml-residual-stable")?.passed, false);
  const nativeSectionPageNumberingResult = await gradeOfficeCase({
    item: sectionPageNumberingItem,
    workspace: sectionPageNumberingRoot,
    evaluator: path.join(sectionPageNumberingRoot, "evaluator"),
    finalMessage: "completed",
    trace: sectionPageNumberingTrace,
  });
  if (nativeSectionPageNumberingResult.graded) {
    assert.equal(nativeSectionPageNumberingResult.rawScorePercent, 100);
    assert.equal(nativeSectionPageNumberingResult.caseSpecificPassed, true);
  } else {
    assert.ok(nativeSectionPageNumberingResult.infrastructureErrors?.length);
  }
} finally {
  await fs.rm(sectionPageNumberingRoot, { recursive: true, force: true });
}

const richNotesItem = cases.find((item) => item.id === "pptx-title-and-notes-edit");
assert.ok(richNotesItem);
const richNotesRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-eval-pptx-rich-notes-"));
try {
  const richNotesInput = path.join(richNotesRoot, "inputs", PPTX_RICH_NOTES_FIXTURE.presentationName);
  const richNotesOutput = path.join(richNotesRoot, "outputs", "rich-notes-review-updated.pptx");
  await generateOfficeInput("pptx-rich-notes-review", richNotesInput);
  const richNotesSource = await fs.readFile(richNotesInput);
  const richNotesPresentation = await PresentationFile.importPptx(new FileBlob(richNotesSource, {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: PPTX_RICH_NOTES_FIXTURE.presentationName,
  }));
  const richNotesSlide = richNotesPresentation.slides.items.find((slide) => slide.name === PPTX_RICH_NOTES_FIXTURE.targetSlideName);
  assert.ok(richNotesSlide);
  const richTitle = richNotesSlide.shapes.items.find((shape) => shape.name === PPTX_RICH_NOTES_FIXTURE.titleShapeName);
  assert.ok(richTitle);
  const richNotesId = richNotesSlide.speakerNotes.id;
  const sourceParagraphs = richNotesSlide.speakerNotes.textFrame.paragraphs;
  assert.equal(sourceParagraphs.length, 2);
  assert.equal(sourceParagraphs[0].runs.length, 2);
  assert.equal(sourceParagraphs[0].runs[1].text, PPTX_RICH_NOTES_FIXTURE.targetRun.expectedText);
  assert.deepEqual(sourceParagraphs[0].runs[1].style, PPTX_RICH_NOTES_FIXTURE.targetRun.expectedStyle);
  const replacementParagraphs = JSON.parse(JSON.stringify(sourceParagraphs));
  replacementParagraphs[PPTX_RICH_NOTES_FIXTURE.targetRun.paragraphIndex].runs[PPTX_RICH_NOTES_FIXTURE.targetRun.runIndex] = {
    ...replacementParagraphs[PPTX_RICH_NOTES_FIXTURE.targetRun.paragraphIndex].runs[PPTX_RICH_NOTES_FIXTURE.targetRun.runIndex],
    text: PPTX_RICH_NOTES_FIXTURE.targetRun.replacementText,
    style: { ...PPTX_RICH_NOTES_FIXTURE.targetRun.replacementStyle },
  };
  richTitle.text.set(PPTX_RICH_NOTES_FIXTURE.replacementTitle);
  richNotesSlide.speakerNotes.textFrame.paragraphs = replacementParagraphs;
  const richNotesExport = await PresentationFile.exportPptx(richNotesPresentation);
  const richNotesBytes = new Uint8Array(await richNotesExport.arrayBuffer());
  await fs.mkdir(path.dirname(richNotesOutput), { recursive: true });
  await fs.writeFile(richNotesOutput, richNotesBytes);
  const hash = (bytes) => crypto.createHash("sha256").update(bytes).digest("hex");
  const richNotesAudit = {
    status: "succeeded",
    source: { sha256: hash(richNotesSource) },
    output: { sha256: hash(richNotesBytes) },
    provider: { actual: "office-kit", version: "test", silentFallback: false },
    savePolicy: { strategy: "rewrite" },
    operation: {
      type: "title-and-rich-speaker-notes-run-edit",
      paragraphIndex: PPTX_RICH_NOTES_FIXTURE.targetRun.paragraphIndex,
      runIndex: PPTX_RICH_NOTES_FIXTURE.targetRun.runIndex,
      expectedRun: { text: PPTX_RICH_NOTES_FIXTURE.targetRun.expectedText },
      replacementRun: { text: PPTX_RICH_NOTES_FIXTURE.targetRun.replacementText },
    },
    validation: { reimport: { ok: true, richNotesFixedTopology: true, targetRunExact: true, notesIdPreserved: true, notesId: richNotesId } },
  };
  await fs.writeFile(path.join(richNotesRoot, "outputs", "audit.json"), JSON.stringify(richNotesAudit, null, 2));
  const richNotesEvidence = {
    source: await inspectRichNotesPptx(richNotesInput),
    output: await inspectRichNotesPptx(richNotesOutput),
    visual: {
      source: { available: true, ok: true, pageCount: 2, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "source" }, { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "stable" }] },
      output: { available: true, ok: true, pageCount: 2, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "output" }, { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "stable" }] },
    },
  };
  const richNotesTrace = JSON.stringify({ type: "item.completed", item: { type: "command_execution", id: "pptx-rich-notes", command: "node -e 'PresentationFile.importPptx(); PresentationFile.exportPptx()'" } });
  const richNotesChecks = gradePptxRichNotesEvidence({
    evidence: richNotesEvidence,
    audit: richNotesAudit,
    commands: extractCompletedCommands(richNotesTrace),
    item: richNotesItem,
  });
  assert.equal(richNotesChecks.every((check) => check.passed), true);
  const siblingDriftPresentation = await PresentationFile.importPptx(new FileBlob(richNotesSource, {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: PPTX_RICH_NOTES_FIXTURE.presentationName,
  }));
  const siblingDriftSlide = siblingDriftPresentation.slides.getItem(0);
  const siblingDriftTitle = siblingDriftSlide.shapes.items.find((shape) => shape.name === PPTX_RICH_NOTES_FIXTURE.titleShapeName);
  assert.ok(siblingDriftTitle);
  siblingDriftTitle.text.set(PPTX_RICH_NOTES_FIXTURE.replacementTitle);
  const siblingDriftParagraphs = siblingDriftSlide.speakerNotes.textFrame.paragraphs;
  siblingDriftParagraphs[0].runs[0].text = "Changed sibling run.";
  siblingDriftParagraphs[PPTX_RICH_NOTES_FIXTURE.targetRun.paragraphIndex].runs[PPTX_RICH_NOTES_FIXTURE.targetRun.runIndex] = {
    ...siblingDriftParagraphs[PPTX_RICH_NOTES_FIXTURE.targetRun.paragraphIndex].runs[PPTX_RICH_NOTES_FIXTURE.targetRun.runIndex],
    text: PPTX_RICH_NOTES_FIXTURE.targetRun.replacementText,
    style: { ...PPTX_RICH_NOTES_FIXTURE.targetRun.replacementStyle },
  };
  siblingDriftSlide.speakerNotes.textFrame.paragraphs = siblingDriftParagraphs;
  const siblingDriftOutput = path.join(richNotesRoot, "outputs", "rich-notes-sibling-drift.pptx");
  await (await PresentationFile.exportPptx(siblingDriftPresentation)).save(siblingDriftOutput);
  const siblingDriftChecks = gradePptxRichNotesEvidence({
    evidence: {
      ...richNotesEvidence,
      output: await inspectRichNotesPptx(siblingDriftOutput),
    },
    audit: richNotesAudit,
    commands: extractCompletedCommands(richNotesTrace),
    item: richNotesItem,
  });
  assert.equal(siblingDriftChecks.find((check) => check.id === "pptx-rich-notes-machine:fixed-topology-and-siblings-preserved")?.passed, false);
  assert.equal(siblingDriftChecks.find((check) => check.id === "pptx-rich-notes-security:fixed-topology-and-package-preservation")?.passed, false);
  const notesOutsideBodyDriftOutput = path.join(richNotesRoot, "outputs", "rich-notes-outside-body-drift.pptx");
  const notesOutsideBodyDriftZip = await JSZip.loadAsync(await fs.readFile(richNotesOutput));
  const notesPath = "ppt/notesSlides/notesSlide1.xml";
  const notesXml = await notesOutsideBodyDriftZip.file(notesPath)?.async("text");
  assert.equal(typeof notesXml, "string");
  const driftedNotesXml = notesXml.replace(/(<p:cNvPr\b[^>]*\bname=")[^"]*(")/, "$1outside-body-drift$2");
  assert.notEqual(driftedNotesXml, notesXml);
  notesOutsideBodyDriftZip.file(notesPath, driftedNotesXml);
  await fs.writeFile(notesOutsideBodyDriftOutput, await notesOutsideBodyDriftZip.generateAsync({ type: "nodebuffer" }));
  const notesOutsideBodyDriftChecks = gradePptxRichNotesEvidence({
    evidence: {
      ...richNotesEvidence,
      output: await inspectRichNotesPptx(notesOutsideBodyDriftOutput),
    },
    audit: richNotesAudit,
    commands: extractCompletedCommands(richNotesTrace),
    item: richNotesItem,
  });
  assert.equal(notesOutsideBodyDriftChecks.find((check) => check.id === "pptx-rich-notes-machine:fixed-topology-and-siblings-preserved")?.passed, false);
  assert.equal(notesOutsideBodyDriftChecks.find((check) => check.id === "pptx-rich-notes-security:fixed-topology-and-package-preservation")?.passed, false);
  const publishedRichNotesWorkflowChecks = gradePptxRichNotesEvidence({
    evidence: richNotesEvidence,
    audit: richNotesAudit,
    commands: ["node .agents/skills/presentations/examples/officekit-rich-speaker-notes-edit-workflow.mjs inputs/rich-notes-review.pptx outputs/rich-notes-review-updated.pptx outputs/audit.json"],
    item: richNotesItem,
  });
  assert.equal(publishedRichNotesWorkflowChecks.find((check) => check.id === "pptx-rich-notes-trace:typed-roundtrip")?.passed, true);
  const untrustedRichNotesWorkflowChecks = gradePptxRichNotesEvidence({
    evidence: richNotesEvidence,
    audit: richNotesAudit,
    commands: ["node scratch/rich-notes-edit.mjs inputs/rich-notes-review.pptx outputs/rich-notes-review-updated.pptx outputs/audit.json"],
    item: richNotesItem,
  });
  assert.equal(untrustedRichNotesWorkflowChecks.find((check) => check.id === "pptx-rich-notes-trace:typed-roundtrip")?.passed, false);
  const nativeRichNotesResult = await gradeOfficeCase({
    item: richNotesItem,
    workspace: richNotesRoot,
    evaluator: path.join(richNotesRoot, "evaluator"),
    finalMessage: "completed",
    trace: richNotesTrace,
  });
  if (nativeRichNotesResult.graded) {
    assert.equal(nativeRichNotesResult.rawScorePercent, 100);
    assert.equal(nativeRichNotesResult.caseSpecificPassed, true);
  } else {
    assert.ok(nativeRichNotesResult.infrastructureErrors?.length);
  }
} finally {
  await fs.rm(richNotesRoot, { recursive: true, force: true });
}

const slideNameItem = cases.find((item) => item.id === "pptx-source-bound-slide-name-edit");
assert.ok(slideNameItem);
const slideNameRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-eval-pptx-slide-name-"));
try {
  const slideNameInput = path.join(slideNameRoot, "inputs", PPTX_SLIDE_NAME_FIXTURE.presentationName);
  const slideNameOutput = path.join(slideNameRoot, "outputs", "launch-review-renamed.pptx");
  await generateOfficeInput("pptx-slide-name-review", slideNameInput);
  const slideNameSource = await fs.readFile(slideNameInput);
  const slideNamePresentation = await PresentationFile.importPptx(new FileBlob(slideNameSource, {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: PPTX_SLIDE_NAME_FIXTURE.presentationName,
  }));
  const slideNameTarget = slideNamePresentation.slides.items.find((slide) => slide.name === PPTX_SLIDE_NAME_FIXTURE.expectedName);
  assert.ok(slideNameTarget);
  slideNameTarget.name = PPTX_SLIDE_NAME_FIXTURE.replacementName;
  const slideNameExport = await PresentationFile.exportPptx(slideNamePresentation);
  const slideNameBytes = new Uint8Array(await slideNameExport.arrayBuffer());
  await fs.mkdir(path.dirname(slideNameOutput), { recursive: true });
  await fs.writeFile(slideNameOutput, slideNameBytes);
  const hash = (bytes) => crypto.createHash("sha256").update(bytes).digest("hex");
  const slideNameSourceEvidence = await inspectTitleNotesPptx(slideNameInput);
  const sourceSlidePart = slideNameSourceEvidence.slides.find((slide) => slide.name === PPTX_SLIDE_NAME_FIXTURE.expectedName)?.path;
  assert.ok(sourceSlidePart);
  const slideNameAudit = {
    status: "succeeded",
    source: { sha256: hash(slideNameSource) },
    output: { sha256: hash(slideNameBytes) },
    provider: { actual: "office-kit", version: "test", silentFallback: false },
    savePolicy: { strategy: "rewrite" },
    operation: {
      type: "source-bound-slide-name-edit",
      sourcePart: sourceSlidePart,
      expectedName: PPTX_SLIDE_NAME_FIXTURE.expectedName,
      replacementName: PPTX_SLIDE_NAME_FIXTURE.replacementName,
      nativeAttribute: "p:cSld/@name",
    },
    validation: { reimport: { ok: true } },
  };
  await fs.writeFile(path.join(slideNameRoot, "outputs", "audit.json"), JSON.stringify(slideNameAudit, null, 2));
  const slideNameEvidence = {
    source: slideNameSourceEvidence,
    output: await inspectTitleNotesPptx(slideNameOutput),
    visual: {
      source: { available: true, ok: true, pageCount: 2, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "target-stable" }, { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "appendix-stable" }] },
      output: { available: true, ok: true, pageCount: 2, pages: [{ width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "target-stable" }, { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "appendix-stable" }] },
    },
  };
  const slideNameTrace = JSON.stringify({ type: "item.completed", item: { type: "command_execution", id: "pptx-slide-name", command: "node -e 'PresentationFile.importPptx(); PresentationFile.exportPptx()'" } });
  const slideNameChecks = gradePptxSlideNameEvidence({
    evidence: slideNameEvidence,
    audit: slideNameAudit,
    commands: extractCompletedCommands(slideNameTrace),
    item: slideNameItem,
  });
  assert.equal(slideNameChecks.every((check) => check.passed), true);
  const semanticDriftEvidence = structuredClone(slideNameEvidence);
  semanticDriftEvidence.output.slides.find((slide) => slide.path === sourceSlidePart).title = "Unexpected visible title change";
  const semanticDriftChecks = gradePptxSlideNameEvidence({
    evidence: semanticDriftEvidence,
    audit: slideNameAudit,
    commands: extractCompletedCommands(slideNameTrace),
    item: slideNameItem,
  });
  assert.equal(semanticDriftChecks.find((check) => check.id === "pptx-name-machine:semantic-content-and-order-preserved")?.passed, false);
  const packageDriftEvidence = structuredClone(slideNameEvidence);
  packageDriftEvidence.output.partHashes["ppt/slides/slide2.xml"] = "unexpected-appendix-part-change";
  const packageDriftChecks = gradePptxSlideNameEvidence({
    evidence: packageDriftEvidence,
    audit: slideNameAudit,
    commands: extractCompletedCommands(slideNameTrace),
    item: slideNameItem,
  });
  assert.equal(packageDriftChecks.find((check) => check.id === "pptx-name-security:fixed-topology-and-non-target-byte-preservation")?.passed, false);
  const publishedSlideNameWorkflowChecks = gradePptxSlideNameEvidence({
    evidence: slideNameEvidence,
    audit: slideNameAudit,
    commands: ["node .agents/skills/presentations/examples/officekit-slide-name-edit-workflow.mjs inputs/launch-review.pptx outputs/launch-review-renamed.pptx outputs/audit.json"],
    item: slideNameItem,
  });
  assert.equal(publishedSlideNameWorkflowChecks.find((check) => check.id === "pptx-name-trace:typed-roundtrip")?.passed, true);
  const untrustedSlideNameWorkflowChecks = gradePptxSlideNameEvidence({
    evidence: slideNameEvidence,
    audit: slideNameAudit,
    commands: ["node scratch/slide-name-edit.mjs inputs/launch-review.pptx outputs/launch-review-renamed.pptx outputs/audit.json"],
    item: slideNameItem,
  });
  assert.equal(untrustedSlideNameWorkflowChecks.find((check) => check.id === "pptx-name-trace:typed-roundtrip")?.passed, false);
  const nativeSlideNameResult = await gradeOfficeCase({
    item: slideNameItem,
    workspace: slideNameRoot,
    evaluator: path.join(slideNameRoot, "evaluator"),
    finalMessage: "completed",
    trace: slideNameTrace,
  });
  if (nativeSlideNameResult.graded) {
    assert.equal(nativeSlideNameResult.rawScorePercent, 100);
    assert.equal(nativeSlideNameResult.caseSpecificPassed, true);
  } else {
    assert.ok(nativeSlideNameResult.infrastructureErrors?.length);
  }
} finally {
  await fs.rm(slideNameRoot, { recursive: true, force: true });
}

const sectionBoundaryItem = cases.find((item) => item.id === "pptx-source-bound-section-boundary-edit");
assert.ok(sectionBoundaryItem);
assert.equal(sectionBoundaryItem.grade.machine.fixedSectionCount, 3);
assert.equal(sectionBoundaryItem.grade.machine.onlyPresentationPartChanged, true);
assert.equal(sectionBoundaryItem.grade.security.fixedSectionIdentityPreserved, true);
assert.match(sectionBoundaryItem.prompt, /目标完整 partition/i);
assert.match(sectionBoundaryItem.prompt, /ppt\/presentation\.xml/);
const sectionBoundaryRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-eval-pptx-section-boundary-"));
try {
  const sectionBoundaryInput = path.join(sectionBoundaryRoot, "inputs", PPTX_SECTION_BOUNDARY_FIXTURE.presentationName);
  const sectionBoundaryOutput = path.join(sectionBoundaryRoot, "outputs", "section-boundary-review-updated.pptx");
  const sectionBoundaryAuditPath = path.join(sectionBoundaryRoot, "outputs", "audit.json");
  await generateOfficeInput("pptx-section-boundary-review", sectionBoundaryInput);
  const sectionBoundarySource = await fs.readFile(sectionBoundaryInput);
  const sectionBoundaryResult = await replacePptxSectionPartition({
    inputPath: sectionBoundaryInput,
    outputPath: sectionBoundaryOutput,
    auditPath: sectionBoundaryAuditPath,
    expectedSections: structuredClone(PPTX_SECTION_BOUNDARY_FIXTURE.sourceSections),
    replacementSections: structuredClone(PPTX_SECTION_BOUNDARY_FIXTURE.replacementSections),
  });
  assert.equal(sectionBoundaryResult.audit.operation.type, "source-bound-section-boundary-edit");
  assert.equal(sectionBoundaryResult.audit.operation.changedSections.length, 2);
  assert.equal(sectionBoundaryResult.audit.validation.package.targetPart, "ppt/presentation.xml");
  assert.equal(sectionBoundaryResult.audit.validation.reimport.exactFixedIdentityPartitionRetained, true);
  assert.deepEqual(await fs.readFile(sectionBoundaryInput), sectionBoundarySource);
  const sectionBoundarySourceEvidence = await inspectSectionBoundaryPptx(sectionBoundaryInput);
  const sectionBoundaryOutputEvidence = await inspectSectionBoundaryPptx(sectionBoundaryOutput);
  const stablePages = PPTX_SECTION_BOUNDARY_FIXTURE.slides.map((_, index) => ({
    width: 1,
    height: 1,
    nonWhitePixels: 1,
    pixelSha256: `section-boundary-page-${index + 1}`,
  }));
  const sectionBoundaryEvidence = {
    source: sectionBoundarySourceEvidence,
    output: sectionBoundaryOutputEvidence,
    visual: {
      source: { available: true, ok: true, pageCount: stablePages.length, pages: stablePages },
      output: { available: true, ok: true, pageCount: stablePages.length, pages: structuredClone(stablePages) },
    },
  };
  const sectionBoundaryTrace = JSON.stringify({ type: "item.completed", item: {
    type: "command_execution",
    id: "pptx-section-boundary",
    command: "node .agents/skills/presentations/examples/officekit-section-boundary-edit-workflow.mjs inputs/section-boundary-review.pptx outputs/section-boundary-review-updated.pptx outputs/audit.json @expected-sections.json @replacement-sections.json",
  } });
  const sectionBoundaryChecks = gradePptxSectionBoundaryEvidence({
    evidence: sectionBoundaryEvidence,
    audit: sectionBoundaryResult.audit,
    commands: extractCompletedCommands(sectionBoundaryTrace),
    item: sectionBoundaryItem,
  });
  assert.equal(sectionBoundaryChecks.every((check) => check.passed), true);
  const partitionDriftEvidence = structuredClone(sectionBoundaryEvidence);
  partitionDriftEvidence.output.sections[1].slidePaths = [
    sectionBoundarySourceEvidence.slideParts[2],
    sectionBoundarySourceEvidence.slideParts[1],
  ];
  const partitionDriftChecks = gradePptxSectionBoundaryEvidence({
    evidence: partitionDriftEvidence,
    audit: sectionBoundaryResult.audit,
    commands: extractCompletedCommands(sectionBoundaryTrace),
    item: sectionBoundaryItem,
  });
  assert.equal(partitionDriftChecks.find((check) => check.id === "pptx-section-boundary-machine:requested-fixed-identity-partition")?.passed, false);
  const packageDriftEvidence = structuredClone(sectionBoundaryEvidence);
  packageDriftEvidence.output.partHashes["ppt/slides/slide4.xml"] = "unexpected-appendix-part-change";
  const packageDriftChecks = gradePptxSectionBoundaryEvidence({
    evidence: packageDriftEvidence,
    audit: sectionBoundaryResult.audit,
    commands: extractCompletedCommands(sectionBoundaryTrace),
    item: sectionBoundaryItem,
  });
  assert.equal(packageDriftChecks.find((check) => check.id === "pptx-section-boundary-security:fixed-topology-and-non-target-byte-preservation")?.passed, false);
  const malformedSectionPath = path.join(sectionBoundaryRoot, "outputs", "section-boundary-malformed.pptx");
  const malformedSectionZip = await JSZip.loadAsync(await fs.readFile(sectionBoundaryOutput));
  const outputPresentationXml = await malformedSectionZip.file("ppt/presentation.xml").async("text");
  const reorderedMembershipXml = outputPresentationXml.replace(
    '<p14:sldId id="257" /><p14:sldId id="258" />',
    '<p14:sldId id="258" /><p14:sldId id="257" />',
  );
  assert.notEqual(reorderedMembershipXml, outputPresentationXml);
  malformedSectionZip.file("ppt/presentation.xml", reorderedMembershipXml);
  await fs.writeFile(malformedSectionPath, await malformedSectionZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => inspectSectionBoundaryPptx(malformedSectionPath),
    /sections must partition every SlidePart once in presentation order/i,
  );
  const publishedSectionBoundaryWorkflowChecks = gradePptxSectionBoundaryEvidence({
    evidence: sectionBoundaryEvidence,
    audit: sectionBoundaryResult.audit,
    commands: ["node .agents/skills/presentations/examples/officekit-section-boundary-edit-workflow.mjs inputs/section-boundary-review.pptx outputs/section-boundary-review-updated.pptx outputs/audit.json @expected-sections.json @replacement-sections.json"],
    item: sectionBoundaryItem,
  });
  assert.equal(publishedSectionBoundaryWorkflowChecks.find((check) => check.id === "pptx-section-boundary-trace:typed-roundtrip")?.passed, true);
  const untrustedSectionBoundaryWorkflowChecks = gradePptxSectionBoundaryEvidence({
    evidence: sectionBoundaryEvidence,
    audit: sectionBoundaryResult.audit,
    commands: ["node scratch/section-boundary-edit.mjs inputs/section-boundary-review.pptx outputs/section-boundary-review-updated.pptx outputs/audit.json"],
    item: sectionBoundaryItem,
  });
  assert.equal(untrustedSectionBoundaryWorkflowChecks.find((check) => check.id === "pptx-section-boundary-trace:typed-roundtrip")?.passed, false);
  const nativeSectionBoundaryResult = await gradeOfficeCase({
    item: sectionBoundaryItem,
    workspace: sectionBoundaryRoot,
    evaluator: path.join(sectionBoundaryRoot, "evaluator"),
    finalMessage: "completed",
    trace: sectionBoundaryTrace,
  });
  if (nativeSectionBoundaryResult.graded) {
    assert.equal(nativeSectionBoundaryResult.rawScorePercent, 100);
    assert.equal(nativeSectionBoundaryResult.caseSpecificPassed, true);
  } else {
    assert.ok(nativeSectionBoundaryResult.infrastructureErrors?.length);
  }
} finally {
  await fs.rm(sectionBoundaryRoot, { recursive: true, force: true });
}

const closedLeafCloneItem = cases.find((item) => item.id === "pptx-closed-leaf-slide-clone");
assert.ok(closedLeafCloneItem);
assert.equal(closedLeafCloneItem.grade.machine.chartParts, 1);
assert.equal(closedLeafCloneItem.grade.machine.oleWorkbookParts, 1);
assert.equal(closedLeafCloneItem.grade.machine.customShowRunLink, true);
assert.equal(closedLeafCloneItem.grade.security.independentChartPart, true);
assert.equal(closedLeafCloneItem.grade.security.independentOleWorkbookPart, true);
assert.equal(closedLeafCloneItem.grade.security.sharedOlePreviewPart, true);
assert.equal(closedLeafCloneItem.grade.security.customShowMembershipStable, true);
assert.match(closedLeafCloneItem.prompt, /独立 ChartPart/);
assert.match(closedLeafCloneItem.prompt, /独立.*XLSX EmbeddedPackagePart/);
assert.match(closedLeafCloneItem.prompt, /custom show/i);
const closedLeafCloneRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-eval-pptx-closed-leaf-clone-"));
try {
  const closedLeafInput = path.join(closedLeafCloneRoot, "inputs", PPTX_CLOSED_LEAF_CLONE_FIXTURE.presentationName);
  const closedLeafOutput = path.join(closedLeafCloneRoot, "outputs", "release-review-with-copy.pptx");
  const closedLeafAuditPath = path.join(closedLeafCloneRoot, "outputs", "audit.json");
  await generateOfficeInput("pptx-closed-leaf-clone", closedLeafInput);
  const closedLeafSource = await fs.readFile(closedLeafInput);
  const closedLeafResult = await duplicatePptxSlide({
    inputPath: closedLeafInput,
    outputPath: closedLeafOutput,
    auditPath: closedLeafAuditPath,
    expectedName: PPTX_CLOSED_LEAF_CLONE_FIXTURE.sourceSlideName,
    allowClosedLeaves: true,
  });
  const closedLeafOutputBytes = await fs.readFile(closedLeafOutput);
  const closedLeafReimport = await PresentationFile.importPptx(new FileBlob(closedLeafOutputBytes, {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "release-review-with-copy.pptx",
  }));
  assert.deepEqual(closedLeafReimport.slides.items.map((slide) => slide.name), [
    PPTX_CLOSED_LEAF_CLONE_FIXTURE.sourceSlideName,
    PPTX_CLOSED_LEAF_CLONE_FIXTURE.sourceSlideName,
    PPTX_CLOSED_LEAF_CLONE_FIXTURE.appendixSlideName,
  ]);
  assert.deepEqual(closedLeafReimport.slides.items.slice(0, 2).map((slide) => slide.speakerNotes.text), [
    PPTX_CLOSED_LEAF_CLONE_FIXTURE.sourceNotes,
    PPTX_CLOSED_LEAF_CLONE_FIXTURE.sourceNotes,
  ]);
  assert.deepEqual(closedLeafReimport.slides.items.slice(0, 2).map((slide) => slide.comments.items[0]?.comments[0]?.text), [
    PPTX_CLOSED_LEAF_CLONE_FIXTURE.sourceComment,
    PPTX_CLOSED_LEAF_CLONE_FIXTURE.sourceComment,
  ]);
  assert.deepEqual(closedLeafReimport.slides.items.slice(0, 2).map((slide) => ({
    title: slide.charts.items[0]?.title,
    categories: slide.charts.items[0]?.categories,
    values: slide.charts.items[0]?.series[0]?.values,
  })), [
    {
      title: PPTX_CLOSED_LEAF_CLONE_FIXTURE.chartTitle,
      categories: [...PPTX_CLOSED_LEAF_CLONE_FIXTURE.chartCategories],
      values: [...PPTX_CLOSED_LEAF_CLONE_FIXTURE.chartValues],
    },
    {
      title: PPTX_CLOSED_LEAF_CLONE_FIXTURE.chartTitle,
      categories: [...PPTX_CLOSED_LEAF_CLONE_FIXTURE.chartCategories],
      values: [...PPTX_CLOSED_LEAF_CLONE_FIXTURE.chartValues],
    },
  ]);
  const sourceOleObject = closedLeafReimport.slides.getItem(0).nativeObjects.items.find((object) => object.name === PPTX_CLOSED_LEAF_CLONE_FIXTURE.oleObjectName);
  const cloneOleObject = closedLeafReimport.slides.getItem(1).nativeObjects.items.find((object) => object.name === PPTX_CLOSED_LEAF_CLONE_FIXTURE.oleObjectName);
  assert.ok(sourceOleObject?.oleWorkbook);
  assert.ok(cloneOleObject?.oleWorkbook);
  assert.notEqual(sourceOleObject.oleWorkbook.partPath, cloneOleObject.oleWorkbook.partPath);
  assert.equal(sourceOleObject.oleWorkbook.sourceSha256, cloneOleObject.oleWorkbook.sourceSha256);
  assert.deepEqual(closedLeafReimport.customShows.getItem(PPTX_CLOSED_LEAF_CLONE_FIXTURE.customShowName).slideIds, [
    closedLeafReimport.slides.getItem(0).id,
    closedLeafReimport.slides.getItem(2).id,
  ]);
  assert.ok(!closedLeafReimport.customShows.getItem(PPTX_CLOSED_LEAF_CLONE_FIXTURE.customShowName).slideIds.includes(closedLeafReimport.slides.getItem(1).id));
  assert.deepEqual(
    closedLeafReimport.slides.getItem(1).shapes.items.find((shape) => shape.name === "board-route-link").text.paragraphs[0].runs[0].link,
    { customShow: PPTX_CLOSED_LEAF_CLONE_FIXTURE.customShowName, returnToSlide: true, tooltip: "Open the board route" },
  );
  assert.equal(closedLeafResult.audit.operation.chartParts.count, 1);
  assert.equal(closedLeafResult.audit.operation.oleWorkbookParts.count, 1);
  assert.equal(closedLeafResult.audit.operation.runHyperlinks.customShowCount, 1);
  assert.equal(closedLeafResult.audit.operation.customShows.count, 1);
  assert.equal(closedLeafResult.audit.validation.package.chartParts.independentParts, true);
  assert.equal(closedLeafResult.audit.validation.package.chartParts.allPayloadsByteIdentical, true);
  assert.equal(closedLeafResult.audit.validation.package.oleWorkbookParts.independentParts, true);
  assert.equal(closedLeafResult.audit.validation.package.oleWorkbookParts.allPayloadsByteIdentical, true);
  assert.equal(closedLeafResult.audit.validation.package.oleWorkbookParts.previewPartsShared, true);
  assert.equal(closedLeafResult.audit.validation.package.customShows.exactSourceMembershipRetained, true);
  const closedLeafEvidence = {
    source: await inspectClosedLeafClonePptx(closedLeafInput),
    output: await inspectClosedLeafClonePptx(closedLeafOutput),
    visual: {
      source: { available: true, ok: true, pageCount: 2, pages: [
        { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "source-slide" },
        { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "appendix-slide" },
      ] },
      output: { available: true, ok: true, pageCount: 3, pages: [
        { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "source-slide" },
        { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "source-slide" },
        { width: 1, height: 1, nonWhitePixels: 1, pixelSha256: "appendix-slide" },
      ] },
    },
  };
  const closedLeafTrace = JSON.stringify({ type: "item.completed", item: { type: "command_execution", id: "pptx-closed-leaf-clone", command: "node .agents/skills/presentations/examples/officekit-slide-duplicate-workflow.mjs inputs/release-review.pptx outputs/release-review-with-copy.pptx outputs/audit.json Release decision --allow-closed-leaves" } });
  const closedLeafChecks = gradePptxClosedLeafCloneEvidence({
    evidence: closedLeafEvidence,
    audit: closedLeafResult.audit,
    commands: extractCompletedCommands(closedLeafTrace),
    item: closedLeafCloneItem,
  });
  assert.equal(closedLeafChecks.every((check) => check.passed), true);
  const irregularCustomShowPath = path.join(closedLeafCloneRoot, "outputs", "irregular-custom-show.pptx");
  const irregularCustomShowZip = await JSZip.loadAsync(closedLeafOutputBytes);
  const irregularPresentationXml = await irregularCustomShowZip.file("ppt/presentation.xml").async("text");
  assert.match(irregularPresentationXml, /<p:custShow\b/);
  irregularCustomShowZip.file(
    "ppt/presentation.xml",
    irregularPresentationXml.replace("</p:custShow>", "<p:extLst /></p:custShow>"),
  );
  await fs.writeFile(irregularCustomShowPath, await irregularCustomShowZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(() => inspectClosedLeafClonePptx(irregularCustomShowPath), /non-canonical custom show/);
  const rootSharedOlePath = path.join(closedLeafCloneRoot, "outputs", "root-shared-ole-workbook.pptx");
  const rootSharedOleZip = await JSZip.loadAsync(closedLeafOutputBytes);
  const rootSharedRelationships = await rootSharedOleZip.file("_rels/.rels").async("text");
  const rootSharedTarget = closedLeafEvidence.output.slides[0].oleWorkbooks[0].part;
  rootSharedOleZip.file(
    "_rels/.rels",
    rootSharedRelationships.replace(
      "</Relationships>",
      `<Relationship Id="rIdRootSharedOleWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="/${rootSharedTarget}"/></Relationships>`,
    ),
  );
  await fs.writeFile(rootSharedOlePath, await rootSharedOleZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(() => inspectClosedLeafClonePptx(rootSharedOlePath), /missing, shared, or owns a child relationship graph/);
  const closedLeafDriftEvidence = structuredClone(closedLeafEvidence);
  closedLeafDriftEvidence.output.partHashes[closedLeafDriftEvidence.output.commentAuthors.part] = "unexpected-comment-authors-change";
  const closedLeafDriftChecks = gradePptxClosedLeafCloneEvidence({
    evidence: closedLeafDriftEvidence,
    audit: closedLeafResult.audit,
    commands: extractCompletedCommands(closedLeafTrace),
    item: closedLeafCloneItem,
  });
  assert.equal(closedLeafDriftChecks.find((check) => check.id === "pptx-clone-security:source-parts-byte-preserved-and-graph-bounded")?.passed, false);
  const chartAliasingEvidence = structuredClone(closedLeafEvidence);
  chartAliasingEvidence.output.slides[1].charts[0].part = chartAliasingEvidence.output.slides[0].charts[0].part;
  const chartAliasingChecks = gradePptxClosedLeafCloneEvidence({
    evidence: chartAliasingEvidence,
    audit: closedLeafResult.audit,
    commands: extractCompletedCommands(closedLeafTrace),
    item: closedLeafCloneItem,
  });
  assert.equal(chartAliasingChecks.find((check) => check.id === "pptx-clone-machine:chart-part-copied-to-independent-leaf")?.passed, false);
  assert.equal(chartAliasingChecks.find((check) => check.id === "pptx-clone-security:source-parts-byte-preserved-and-graph-bounded")?.passed, false);
  const oleAliasingEvidence = structuredClone(closedLeafEvidence);
  oleAliasingEvidence.output.slides[1].oleWorkbooks[0].part = oleAliasingEvidence.output.slides[0].oleWorkbooks[0].part;
  const oleAliasingChecks = gradePptxClosedLeafCloneEvidence({
    evidence: oleAliasingEvidence,
    audit: closedLeafResult.audit,
    commands: extractCompletedCommands(closedLeafTrace),
    item: closedLeafCloneItem,
  });
  assert.equal(oleAliasingChecks.find((check) => check.id === "pptx-clone-machine:ole-workbook-copied-to-independent-package")?.passed, false);
  assert.equal(oleAliasingChecks.find((check) => check.id === "pptx-clone-security:source-parts-byte-preserved-and-graph-bounded")?.passed, false);
  const customShowMembershipDriftEvidence = structuredClone(closedLeafEvidence);
  customShowMembershipDriftEvidence.output.customShows[0].slideParts.splice(1, 0, customShowMembershipDriftEvidence.output.slides[1].part);
  const customShowMembershipDriftChecks = gradePptxClosedLeafCloneEvidence({
    evidence: customShowMembershipDriftEvidence,
    audit: closedLeafResult.audit,
    commands: extractCompletedCommands(closedLeafTrace),
    item: closedLeafCloneItem,
  });
  assert.equal(customShowMembershipDriftChecks.find((check) => check.id === "pptx-clone-machine:custom-show-action-retained-without-membership-drift")?.passed, false);
  assert.equal(customShowMembershipDriftChecks.find((check) => check.id === "pptx-clone-security:source-parts-byte-preserved-and-graph-bounded")?.passed, false);
  const missingOptInChecks = gradePptxClosedLeafCloneEvidence({
    evidence: closedLeafEvidence,
    audit: closedLeafResult.audit,
    commands: ["node .agents/skills/presentations/examples/officekit-slide-duplicate-workflow.mjs inputs/release-review.pptx outputs/release-review-with-copy.pptx outputs/audit.json Release decision"],
    item: closedLeafCloneItem,
  });
  assert.equal(missingOptInChecks.find((check) => check.id === "pptx-clone-trace:typed-roundtrip")?.passed, false);
  const nativeClosedLeafResult = await gradeOfficeCase({
    item: closedLeafCloneItem,
    workspace: closedLeafCloneRoot,
    evaluator: path.join(closedLeafCloneRoot, "evaluator"),
    finalMessage: "completed",
    trace: closedLeafTrace,
  });
  if (nativeClosedLeafResult.graded) {
    assert.equal(nativeClosedLeafResult.rawScorePercent, 100);
    assert.equal(nativeClosedLeafResult.caseSpecificPassed, true);
  } else {
    assert.ok(nativeClosedLeafResult.infrastructureErrors?.length);
  }
  assert.match(closedLeafResult.audit.output.sha256, /^[0-9a-f]{64}$/);
  assert.notEqual(closedLeafResult.audit.source.sha256, closedLeafResult.audit.output.sha256);
  assert.notDeepEqual(closedLeafSource, closedLeafOutputBytes);
} finally {
  await fs.rm(closedLeafCloneRoot, { recursive: true, force: true });
}

const accessibleItem = cases.find((item) => item.id === "pdf-greenfield-accessible-report");
const accessiblePages = Array.from({ length: 6 }, (_, index) => ({ page: index + 1, width: 1224, height: 1584, nonBlank: true, inkBBox: [50, 50, 1100, 1500], touchesEdge: false, bytes: 20_000 }));
const accessibleEvidence = {
  source: { sha256: "accessible-source-sha" },
  output: { sha256: "accessible-output-sha", pageCount: 6 },
  structure: {
    tagged: true,
    language: "zh-CN",
    title: "Agent Artifact Readiness",
    roles: { H1: 1, H2: 4, H3: 7, Table: 1, TR: 5, TH: 6, TD: 9, Figure: 1, Link: 1 },
    tables: [{ id: "risk-register", pages: [3, 4], rows: 5, headers: 6, dataCells: 9 }],
    figuresWithAlt: 1,
    links: [{ page: 6, uri: "https://www.w3.org/WAI/", structParent: 6 }],
    linkObjrAssociations: 1,
    artifactMarkers: 12,
    rootIds: ["board-page-1/text", "summary-h2", "summary-h3", "format-pass-rate-chart", "risks-h2", "risks-h3", "risk-register", "mitigation-h3", "validation-h2", "modeled-h3", "machine-h3", "human-h3", "conclusion-h2", "conclusion-h3", "wai-guidance-link"],
    pageText: ["封面", "摘要", "风险 级别 缓解措施", "风险 级别 缓解措施", "验证", "结论"],
  },
  visual: { renderer: "poppler-pdftoppm", pageCount: 6, pages: accessiblePages },
};
const accessibleAudit = {
  status: "succeeded",
  source: { sha256: "accessible-source-sha" },
  output: { sha256: "accessible-output-sha" },
  provider: { actual: "artifact-tool", version: "0.2.0", silentFallback: false },
  savePolicy: { strategy: "rewrite" },
  operation: { type: "create-accessible-report" },
  validation: {
    modeledVerify: { status: "passed", scope: "PdfArtifact modeled invariants" },
    poppler: { status: "passed", pages: accessiblePages },
    veraPdfMachine: { available: false, status: "not-run", claim: "No veraPDF machine validation was performed." },
    humanPdfUa: { status: "required", claim: "No complete PDF/UA certification is claimed." },
  },
};
const accessibleChecks = gradeAccessibleReportEvidence({
  evidence: accessibleEvidence,
  audit: accessibleAudit,
  commands: ["node .agents/skills/pdf/examples/accessible-board-report.mjs inputs/report-data.json outputs/readiness-report.pdf outputs/audit.json"],
  finalMessage: "已创建；不声明完整 PDF/UA 认证。",
  item: accessibleItem,
});
assert.equal(accessibleChecks.every((entry) => entry.passed), true);
assert.equal(summarizeCaseScore(accessibleChecks, accessibleItem.grade).rawScorePercent, 100);
const overclaimChecks = gradeAccessibleReportEvidence({
  evidence: accessibleEvidence,
  audit: accessibleAudit,
  commands: ["node .agents/skills/pdf/examples/accessible-board-report.mjs inputs/report-data.json outputs/readiness-report.pdf outputs/audit.json"],
  finalMessage: "The report is PDF/UA certified.",
  item: accessibleItem,
});
assert.equal(overclaimChecks.find((entry) => entry.id === "pdf-security:no-pdfua-overclaim")?.passed, false);

const mergeItem = cases.find((item) => item.id === "pdf-merge-reorder-stamp-links");
const mergeSequence = ["cover:1", "appendix:3", "report:1", "report:2", "appendix:1", "appendix:2"];
const geometryBySource = {
  cover: { boxes: { mediabox: [0, 0, 612, 792] }, rotation: 0 },
  report: { boxes: { mediabox: [0, 0, 792, 612] }, rotation: 0 },
  appendix: { boxes: { mediabox: [0, 0, 595.2756, 841.8898] }, rotation: 0 },
};
const mergePageMap = mergeSequence.map((identity, index) => {
  const [source, sourcePage] = identity.split(":");
  return {
    outputPage: index + 1,
    source,
    sourcePage: Number(sourcePage),
    sourceGeometry: geometryBySource[source],
    outputGeometry: structuredClone(geometryBySource[source]),
    watermarkCount: [3, 4].includes(index + 1) ? 1 : 0,
    opacities: [3, 4].includes(index + 1) ? [0.2] : [],
  };
});
const mergeOutlines = mergeSequence.map((identity, index) => {
  const [source, sourcePage] = identity.split(":");
  return { title: `${source}-${sourcePage}`, page: index + 1, parentPath: [] };
});
const mergeNamed = Object.fromEntries(mergeOutlines.map((entry) => [`${entry.title}-named`, entry.page]));
const mergeLinks = mergeOutlines.map((entry, index) => ({ page: entry.page, targetPage: mergeOutlines[(index + 1) % mergeOutlines.length].page, rect: [68, 100, 280, 124] }));
const mergePathRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-eval-merge-paths-"));
const mergePathAlias = `${mergePathRoot}-alias`;
await fs.symlink(mergePathRoot, mergePathAlias);
const mergeSources = Object.fromEntries([
  ["cover", 1], ["report", 2], ["appendix", 3],
].map(([id, pageCount]) => [id, { path: path.join(mergePathRoot, `${id}.pdf`), bytes: 100 + pageCount, sha256: `${id}-sha`, pageCount, termCounts: { CONFIDENTIAL: 0 } }]));
await Promise.all(Object.values(mergeSources).map((source) => fs.writeFile(source.path, "fixture", "utf8")));
const mergeEvidence = {
  manifest: {
    path: "/tmp/merge-stamp.json",
    bytes: 400,
    sha256: "manifest-sha",
    value: {
      schema: "office-kit.pdf-merge-stamp.v1",
      sources: [{ id: "cover" }, { id: "report" }, { id: "appendix" }],
      sequence: [
        { source: "cover", pages: "all" },
        { source: "appendix", pages: [3] },
        { source: "report", pages: "all" },
        { source: "appendix", pages: [1, 2] },
      ],
      watermarks: [{ source: "report", text: "CONFIDENTIAL", opacity: 0.2 }],
    },
  },
  sources: mergeSources,
  output: { sha256: "merge-output-sha", pageCount: 6, startxrefCount: 1, eofCount: 1, decodedStreamErrors: [] },
  pageMap: mergePageMap,
  navigation: { expected: { outlines: mergeOutlines, namedDestinations: mergeNamed, internalLinks: mergeLinks }, actual: { outlines: structuredClone(mergeOutlines), namedDestinations: Object.fromEntries(Object.entries(mergeNamed).reverse()), internalLinks: structuredClone(mergeLinks) } },
  visual: { renderer: "poppler-pdftoppm", pageCount: 6, pages: mergePageMap.map((entry) => ({ page: entry.outputPage, source: entry.source, sourcePage: entry.sourcePage, sameDimensions: true, nonBlank: true, pixelStable: ![3, 4].includes(entry.outputPage), changedPixelsBBox: [3, 4].includes(entry.outputPage) ? [400, 300, 800, 700] : null, watermarkExpected: [3, 4].includes(entry.outputPage) })) },
};
const mergeAudit = {
  status: "succeeded",
  source: { sha256: "manifest-sha" },
  inputs: Object.values(mergeSources).map(({ path: sourcePath, bytes, sha256 }) => ({ path: path.join(mergePathAlias, path.basename(sourcePath)), bytes, sha256 })),
  output: { sha256: "merge-output-sha" },
  provider: { actual: "pypdf", version: "6.10.0", silentFallback: false },
  savePolicy: { strategy: "rewrite" },
  preflight: { probeCompleted: true, planCompleted: true },
  operation: { type: "merge-stamp" },
};
const mergeCommands = [
  "python -c 'from reportlab.pdfgen import canvas; print(canvas)'",
  `"$PYTHON_BIN" "$S/pdf_provider.py" check --provider pypdf --require`,
  `"$PYTHON_BIN" "$S/pdf_provider.py" plan --task merge-stamp --provider pypdf --strategy rewrite --input merge-stamp.json --output outputs/merged.pdf --require-provider`,
  `"$PYTHON_BIN" "$S/pypdf_edit.py" merge-stamp merge-stamp.json outputs/merged.pdf --strategy rewrite`,
  `"$PYTHON_BIN" "$S/poppler_compare.py" merge-stamp merge-stamp.json outputs/merged.pdf --report merge-visual.json --render-dir merge-rendered`,
  `"$PYTHON_BIN" "$S/pdf_audit.py" validate outputs/audit.json --source merge-stamp.json --input inputs/cover.pdf --input inputs/report.pdf --input inputs/appendix.pdf --artifact outputs/merged.pdf --require-operation merge-stamp`,
];
const mergeChecks = gradeMergeStampEvidence({ evidence: mergeEvidence, audit: mergeAudit, commands: mergeCommands, item: mergeItem });
assert.equal(mergeChecks.every((entry) => entry.passed), true);
assert.equal(summarizeCaseScore(mergeChecks, mergeItem.grade).rawScorePercent, 100);
const danglingMergeEvidence = structuredClone(mergeEvidence);
danglingMergeEvidence.navigation.actual.internalLinks[0].targetPage = 99;
const danglingMergeChecks = gradeMergeStampEvidence({ evidence: danglingMergeEvidence, audit: mergeAudit, commands: mergeCommands, item: mergeItem });
assert.equal(danglingMergeChecks.find((entry) => entry.id === "pdf-security:navigation-resolved")?.passed, false);
const adHocMergeChecks = gradeMergeStampEvidence({ evidence: mergeEvidence, audit: mergeAudit, commands: [...mergeCommands, "python -c 'from pypdf import PdfWriter; PdfWriter()'"], item: mergeItem });
assert.equal(adHocMergeChecks.find((entry) => entry.id === "pdf-trace:no-ad-hoc-pdf-writer")?.passed, false);
await fs.rm(mergePathAlias);
await fs.rm(mergePathRoot, { recursive: true, force: true });

const providerInstruction = providerRuntimeInstruction(mergeItem, { OFFICE_KIT_AGENT_EVAL_PYTHON: "/opt/eval python/bin/python3" });
assert.match(providerInstruction, /OFFICE_KIT_PDF_PROVIDER_PYTHON="\/opt\/eval python\/bin\/python3"/);
assert.match(providerInstruction, /Do not replace it/);
const splitRuntimeInstruction = providerRuntimeInstruction(mergeItem, {
  OFFICE_KIT_AGENT_EVAL_PYTHON: "/opt/oracle/bin/python3",
  OFFICE_KIT_AGENT_EVAL_PROVIDER_PYTHON: "/opt/provider/bin/python3",
  OFFICE_KIT_PDF_PROVIDER_PYTHON: "/opt/legacy-provider/bin/python3",
});
assert.match(splitRuntimeInstruction, /OFFICE_KIT_PDF_PROVIDER_PYTHON="\/opt\/provider\/bin\/python3"/);
assert.doesNotMatch(splitRuntimeInstruction, /oracle\/bin\/python3|legacy-provider\/bin\/python3/);
assert.equal(providerRuntimeInstruction({ family: "xlsx" }, { OFFICE_KIT_AGENT_EVAL_PYTHON: "/opt/python" }), "");

const validationSampleId = "pdf-bounded-contract-id-replace";
const badNetwork = structuredClone(cases);
badNetwork.find((item) => item.id === validationSampleId).policy.network = true;
assert.throws(() => validateSuite(suite, badNetwork), /network must be false/);
const badPath = structuredClone(cases);
badPath.find((item) => item.id === validationSampleId).inputs[0].to = "../source.pdf";
assert.throws(() => validateSuite(suite, badPath), /escapes the workspace|under inputs/);
const normalizedEscape = structuredClone(cases);
normalizedEscape.find((item) => item.id === validationSampleId).deliverables[0].path = "outputs/../inputs/source.pdf";
assert.throws(() => validateSuite(suite, normalizedEscape), /deliverable must stay under outputs/);

const boundedItem = cases.find((item) => item.id === "pdf-bounded-contract-id-replace");
const oldContractId = "ACME-2025-041";
const newContractId = "ACME-2026-041";
const sourcePages = Array.from({ length: 5 }, (_, index) => ({
  page: index + 1,
  width: 612,
  height: 792,
  rotation: 0,
  termCounts: { [oldContractId]: index === 2 ? 1 : 0, [newContractId]: 0 },
}));
const outputPages = sourcePages.map((page, index) => ({ ...page, termCounts: { [oldContractId]: 0, [newContractId]: index === 2 ? 1 : 0 } }));
const boundedEvidence = {
  source: { sha256: "source-sha", pageCount: 5, pages: sourcePages, termCounts: { [oldContractId]: 1, [newContractId]: 0 }, decodedStreamErrors: [] },
  output: {
    sha256: "output-sha",
    pageCount: 5,
    pages: outputPages,
    termCounts: { [oldContractId]: 0, [newContractId]: 1 },
    rawTermCounts: { [oldContractId]: 0 },
    decodedStreamTermCounts: { [oldContractId]: 0 },
    metadataTermCounts: { [oldContractId]: 0 },
    decodedStreamErrors: [],
    startxrefCount: 1,
    eofCount: 1,
  },
  sourceStyle: { found: true, fonts: ["Helvetica"], sizes: [11], bbox: [172, 152, 254, 167] },
  outputStyle: { found: true, fonts: ["Helvetica"], sizes: [11], bbox: [172, 152, 254, 167] },
  visual: {
    renderer: "poppler-pdftoppm",
    sourcePageCount: 5,
    outputPageCount: 5,
    allowedMask: { page: 3, bboxPx: [330, 296, 470, 344] },
    pages: Array.from({ length: 5 }, (_, index) => ({
      page: index + 1,
      sameDimensions: true,
      nonBlank: true,
      changedPixelsBBox: index === 2 ? [399, 312, 411, 329] : null,
      changedWithinAllowedMask: index === 2,
    })),
  },
};
const boundedAudit = {
  status: "succeeded",
  source: { sha256: "source-sha" },
  output: { sha256: "output-sha" },
  provider: { actual: "pymupdf", version: "1.27.2.3", silentFallback: false },
  savePolicy: { strategy: "sanitize" },
  preflight: { probeCompleted: true, planCompleted: true },
  operation: { type: "replace_text" },
};
const boundedCommands = [
  "python .agents/skills/pdf/scripts/pymupdf_edit.py probe --accept-license agpl",
  "python .agents/skills/pdf/scripts/pymupdf_edit.py edit inputs/source.pdf outputs/contract-updated.pdf",
];
const boundedChecks = gradeBoundedReplaceEvidence({ evidence: boundedEvidence, audit: boundedAudit, commands: boundedCommands, item: boundedItem });
assert.equal(boundedChecks.every((entry) => entry.passed), true);
assert.deepEqual(summarizeCaseScore(boundedChecks, boundedItem.grade), {
  categoryScores: {
    machine: { applicable: true, weight: 45, passed: true, checks: 7 },
    visual: { applicable: true, weight: 25, passed: true, checks: 3 },
    security: { applicable: true, weight: 20, passed: true, checks: 4 },
    trace: { applicable: true, weight: 10, passed: true, checks: 7 },
  },
  rawScorePercent: 100,
  scorePercent: 100,
  caseSpecificPassed: true,
});
const undecodableEvidence = structuredClone(boundedEvidence);
undecodableEvidence.output.decodedStreamErrors.push({ object: 9, generation: 0, error: "test decode failure" });
const undecodableChecks = gradeBoundedReplaceEvidence({ evidence: undecodableEvidence, audit: boundedAudit, commands: boundedCommands, item: boundedItem });
assert.equal(undecodableChecks.find((entry) => entry.id === "pdf-security:all-streams-decodable")?.passed, false);
const shiftedEvidence = structuredClone(boundedEvidence);
shiftedEvidence.outputStyle.bbox[0] += 1;
const shiftedChecks = gradeBoundedReplaceEvidence({ evidence: shiftedEvidence, audit: boundedAudit, commands: boundedCommands, item: boundedItem });
assert.equal(shiftedChecks.find((entry) => entry.id === "pdf-machine:font-geometry-unchanged")?.passed, false);

const bypassAudit = structuredClone(boundedAudit);
bypassAudit.savePolicy.strategy = "rewrite";
bypassAudit.operation.type = "direct_content_stream_equal_length_replace";
const bypassChecks = gradeBoundedReplaceEvidence({ evidence: boundedEvidence, audit: bypassAudit, commands: [...boundedCommands, "doc.update_stream(xref, data.replace(old, new))"], item: boundedItem });
assert.equal(bypassChecks.filter((entry) => entry.category !== "trace").every((entry) => entry.passed), true);
assert.equal(bypassChecks.find((entry) => entry.id === "pdf-trace:no-content-stream-bypass")?.passed, false);
assert.equal(summarizeCaseScore(bypassChecks, boundedItem.grade).rawScorePercent, 90);
const postHocProbeChecks = gradeBoundedReplaceEvidence({
  evidence: boundedEvidence,
  audit: boundedAudit,
  commands: [boundedCommands[1], boundedCommands[0]],
  item: boundedItem,
});
assert.equal(postHocProbeChecks.find((entry) => entry.id === "pdf-trace:capability-probe")?.passed, false);

const highlightItem = cases.find((item) => item.id === "pdf-source-bound-text-highlight");
const highlightText = highlightItem.grade.machine.text;
const highlightSourcePages = Array.from({ length: 3 }, (_, index) => ({
  page: index + 1,
  width: 612,
  height: 792,
  rotation: 0,
  termCounts: { [highlightText]: index === 1 ? 1 : 0 },
}));
const highlightEvidence = {
  source: { sha256: "highlight-source-sha", pageCount: 3, pages: highlightSourcePages, termCounts: { [highlightText]: 1 }, decodedStreamErrors: [] },
  output: {
    sha256: "highlight-output-sha",
    pageCount: 3,
    pages: highlightSourcePages.map((page) => structuredClone(page)),
    termCounts: { [highlightText]: 1 },
    decodedStreamErrors: [],
    startxrefCount: 1,
    eofCount: 1,
  },
  sourceTarget: { found: true, term: highlightText, page: 2, bbox: [84, 152, 294, 166], fonts: ["Helvetica"], sizes: [11] },
  outputTarget: { found: true, term: highlightText, page: 2, bbox: [84, 152, 294, 166], fonts: ["Helvetica"], sizes: [11] },
  sourceHighlights: [],
  outputHighlights: [{
    page: 2,
    contents: highlightItem.grade.machine.contents,
    author: highlightItem.grade.machine.author,
    subject: highlightItem.grade.machine.subject,
    color: [1, 0.92000000298, 0.20000000298],
    quadPoints: [84, 640, 294, 640, 84, 626, 294, 626],
    rect: [84, 626, 294, 640],
  }],
  visual: {
    renderer: "poppler-pdftoppm",
    sourcePageCount: 3,
    outputPageCount: 3,
    allowedMask: { page: 2, bboxPx: [160, 296, 596, 340] },
    pages: [
      { page: 1, sameDimensions: true, nonBlank: true, changedPixelsBBox: null, changedWithinAllowedMask: false },
      { page: 2, sameDimensions: true, nonBlank: true, changedPixelsBBox: [168, 301, 588, 338], changedWithinAllowedMask: true },
      { page: 3, sameDimensions: true, nonBlank: true, changedPixelsBBox: null, changedWithinAllowedMask: false },
    ],
  },
};
const highlightAudit = {
  status: "succeeded",
  source: { sha256: "highlight-source-sha" },
  output: { sha256: "highlight-output-sha" },
  provider: { actual: "mupdf", version: "1.28.0", silentFallback: false },
  savePolicy: { strategy: "rewrite" },
  operation: {
    type: "add_text_highlight",
    page: 2,
    sourceSha256: "highlight-source-sha",
    expectedPage: { bbox: [0, 0, 612, 792], rotation: 0 },
    text: highlightText,
    color: [1, 0.92, 0.2],
    contents: highlightItem.grade.machine.contents,
    author: highlightItem.grade.machine.author,
    subject: highlightItem.grade.machine.subject,
  },
};
const highlightCommands = [
  "node .agents/skills/pdf/scripts/mupdf.mjs probe",
  "node .agents/skills/pdf/scripts/mupdf.mjs inspect inputs/source.pdf > tmp/source-inspection.json",
  "node .agents/skills/pdf/scripts/mupdf.mjs edit inputs/source.pdf tmp/highlight-operation.json outputs/review-highlighted.pdf --save-policy rewrite",
  "node .agents/skills/pdf/scripts/mupdf.mjs inspect outputs/review-highlighted.pdf > tmp/output-inspection.json",
  "node .agents/skills/pdf/scripts/mupdf.mjs render outputs/review-highlighted.pdf outputs/review-highlighted.png --page 2",
];
const highlightChecks = gradeSourceBoundHighlightEvidence({ evidence: highlightEvidence, audit: highlightAudit, commands: highlightCommands, item: highlightItem });
assert.equal(highlightChecks.every((entry) => entry.passed), true);
assert.deepEqual(summarizeCaseScore(highlightChecks, highlightItem.grade), {
  categoryScores: {
    machine: { applicable: true, weight: 45, passed: true, checks: 6 },
    visual: { applicable: true, weight: 25, passed: true, checks: 3 },
    security: { applicable: true, weight: 20, passed: true, checks: 3 },
    trace: { applicable: true, weight: 10, passed: true, checks: 8 },
  },
  rawScorePercent: 100,
  scorePercent: 100,
  caseSpecificPassed: true,
});
const staleHighlightAudit = structuredClone(highlightAudit);
staleHighlightAudit.operation.expectedPage.rotation = 90;
const staleHighlightChecks = gradeSourceBoundHighlightEvidence({ evidence: highlightEvidence, audit: staleHighlightAudit, commands: highlightCommands, item: highlightItem });
assert.equal(staleHighlightChecks.find((entry) => entry.id === "pdf-security:source-bound-operation")?.passed, false);
const directHighlightChecks = gradeSourceBoundHighlightEvidence({ evidence: highlightEvidence, audit: highlightAudit, commands: [...highlightCommands, "page.addAnnotation('Highlight').setQuadPoints(points)"], item: highlightItem });
assert.equal(directHighlightChecks.find((entry) => entry.id === "pdf-trace:no-caller-coordinate-or-direct-mutation")?.passed, false);

const overflowItem = cases.find((item) => item.id === "pdf-overflow-replace-refusal");
const overflowEvidence = {
  source: { sha256: "overflow-source-sha", termCounts: { Approved: 1 }, pages: [{ termCounts: { Approved: 1 } }] },
  geometry: { targetRectangle: { width: 70 }, replacementWidth: 291.027, availableWidth: 64, fits: false },
};
const overflowAudit = {
  status: "failed_closed",
  delivered_modified_pdf: false,
  reason: "replacement overflows the available box",
  source: { sha256: "overflow-source-sha" },
  provider: { selected: "PyMuPDF", version: "1.27.2.3", fallback_used: false },
  operation: { performed: "no PDF mutation; failed before edit" },
  geometry_evidence: { replacement_text_width_pt: 291.027, available_width_from_text_origin_pt: 64, fits_outer_cell_without_wrap: false },
  validation: { no_partial_modified_pdf_in_outputs: true },
};
const overflowChecks = gradeOverflowRefusalEvidence({
  evidence: overflowEvidence,
  audit: overflowAudit,
  commands: ["python -c 'import fitz; print(fitz.VersionBind)'"],
  finalMessage: "The replacement would overflow and does not fit.",
  item: overflowItem,
});
assert.equal(overflowChecks.every((entry) => entry.passed), true);
assert.equal(summarizeCaseScore(overflowChecks, overflowItem.grade).rawScorePercent, 100);
const mutatedOverflowChecks = gradeOverflowRefusalEvidence({
  evidence: overflowEvidence,
  audit: overflowAudit,
  commands: ["python -c 'import fitz; doc.save(\"outputs/partial.pdf\")'"],
  finalMessage: "The replacement would overflow and does not fit.",
  item: overflowItem,
});
assert.equal(mutatedOverflowChecks.find((entry) => entry.id === "pdf-trace:no-mutation-after-failed-preflight")?.passed, false);

const formItem = cases.find((item) => item.id === "pdf-acroform-visible-preserved");
const formFieldNames = Object.keys(formItem.grade.machine.fields);
const formSourceFields = Object.fromEntries(formFieldNames.map((name) => [name, {
  fieldType: name === "company_type" ? "/Btn" : "/Tx",
  value: "",
  defaultValue: "",
  readOnly: false,
  states: name === "company_type" ? ["/LLC", "/Corporation"] : [],
}]));
formSourceFields.terms_ack = { fieldType: "/Btn", value: "/Yes", defaultValue: "", readOnly: false, states: ["/Off", "/Yes"] };
const formOutputFields = structuredClone(formSourceFields);
for (const [name, value] of Object.entries(formItem.grade.machine.fields)) {
  formOutputFields[name].value = name === "company_type" ? `/${value}` : value;
}
const formWidgets = [
  ["full_name", "/Tx", [190, 88, 450, 112], []],
  ["address", "/Tx", [190, 133, 450, 157], []],
  ["effective_date", "/Tx", [190, 178, 450, 202], []],
  ["tin", "/Tx", [190, 223, 450, 247], []],
  ["signature", "/Tx", [190, 268, 450, 292], []],
  ["company_type", "/Btn", [190, 317, 210, 337], ["/LLC", "/Off"]],
  ["company_type", "/Btn", [260, 317, 280, 337], ["/Corporation", "/Off"]],
  ["terms_ack", "/Btn", [260, 362, 280, 382], ["/Off", "/Yes"]],
].map(([name, fieldType, rect, appearanceStates]) => ({
  page: 1,
  name,
  fieldType,
  rect,
  appearancePresent: true,
  appearanceStates,
  selectedState: "/Off",
  readOnly: false,
}));
formWidgets[7].selectedState = "/Yes";
const outputFormWidgets = structuredClone(formWidgets);
outputFormWidgets[5].selectedState = "/LLC";
const formWidgetChanges = formWidgets.map((widget, index) => ({
  name: widget.name,
  page: 1,
  fieldType: widget.fieldType,
  appearanceStates: widget.appearanceStates,
  expectedChange: [0, 1, 2, 5].includes(index),
  changedPixelsBBox: [0, 1, 2, 5].includes(index) ? [4, 4, 20, 18] : null,
  changedInteriorPixelsBBox: [0, 1, 2, 5].includes(index) ? [1, 1, 16, 14] : null,
}));
const formEvidence = {
  source: { sha256: "form-source-sha", pageCount: 1, pages: [{ page: 1, width: 612, height: 792, rotation: 0 }], decodedStreamErrors: [], startxrefCount: 1, eofCount: 1 },
  output: { sha256: "form-output-sha", pageCount: 1, pages: [{ page: 1, width: 612, height: 792, rotation: 0 }], decodedStreamErrors: [], startxrefCount: 2, eofCount: 2 },
  sourceForm: { acroFormPresent: true, needAppearances: false, fieldTreeRoots: 7, fields: formSourceFields, widgets: formWidgets },
  outputForm: { acroFormPresent: true, needAppearances: false, fieldTreeRoots: 7, fields: formOutputFields, widgets: outputFormWidgets },
  originalPrefixPreserved: true,
  visual: {
    renderer: "poppler-pdftoppm",
    sourcePageCount: 1,
    outputPageCount: 1,
    pages: [{ page: 1, sameDimensions: true, nonBlank: true, changedOnlyWithinAllowedMasks: true, changedOutsideAllowedMasksBBox: null }],
    widgetChanges: formWidgetChanges,
  },
};
const formAudit = {
  status: "succeeded",
  source: { sha256: "form-source-sha" },
  output: { sha256: "form-output-sha" },
  provider: { actual: "pypdf", version: "6.10.0", silentFallback: false },
  savePolicy: { strategy: "incremental" },
  preflight: { probeCompleted: true, planCompleted: true },
  operation: { type: "fill-form" },
};
const formCommands = [
  "python pypdf_edit.py inspect inputs/source.pdf --output tmp/inspect.json",
  "python pdf_provider.py check --provider pypdf --require",
  "python pdf_provider.py plan --task fill-form --provider pypdf --strategy incremental",
  "python pypdf_edit.py fill-form inputs/source.pdf outputs/form-filled.pdf --strategy incremental",
  "pdftoppm -png outputs/form-filled.pdf tmp/form-page",
  "python pdf_audit.py validate outputs/audit.json --artifact outputs/form-filled.pdf",
];
const formChecks = gradeAcroFormEvidence({ evidence: formEvidence, audit: formAudit, commands: formCommands, item: formItem });
assert.equal(formChecks.every((entry) => entry.passed), true);
assert.equal(summarizeCaseScore(formChecks, formItem.grade).rawScorePercent, 100);
const brokenRadioEvidence = structuredClone(formEvidence);
brokenRadioEvidence.outputForm.widgets[5].selectedState = "/Off";
const brokenRadioChecks = gradeAcroFormEvidence({ evidence: brokenRadioEvidence, audit: formAudit, commands: formCommands, item: formItem });
assert.equal(brokenRadioChecks.find((entry) => entry.id === "pdf-machine:radio-appearance-state")?.passed, false);
const formBypassChecks = gradeAcroFormEvidence({
  evidence: formEvidence,
  audit: formAudit,
  commands: [...formCommands, "writer = PdfWriter(reader, incremental=True); writer.update_page_form_field_values(None, values)"],
  item: formItem,
});
assert.equal(formBypassChecks.find((entry) => entry.id === "pdf-trace:no-ad-hoc-pypdf-writer")?.passed, false);

const attachmentItem = cases.find((item) => item.id === "pdf-attachment-quarantine-inventory");
const attachmentPayloads = [
  ["../escape.exe", "document", null, null, "application/vnd.microsoft.portable-executable", 17, "hash-escape", "escape.exe"],
  ["archive.zip", "document", null, null, "application/zip", 13, "hash-archive", "archive.zip"],
  ["report.txt", "document", null, null, "text/plain", 5, "hash-first", "report.txt"],
  ["report.txt", "document", null, null, "text/plain", 6, "hash-second", "report__2.txt"],
  ["unicode-测试.txt", "document", null, null, "text/plain", 7, "hash-unicode", "unicode-测试.txt"],
  ["report.txt", "page", 1, 0, "text/plain", 17, "hash-page", "report__3.txt"],
];
const expectedAttachments = attachmentPayloads.map(([displayName, scope, page, annotationIndex, mime, bytes, sha256], index) => ({
  scope,
  page,
  annotationIndex,
  internalKey: displayName,
  displayName,
  mime,
  bytes,
  sha256,
  ordinal: index + 1,
}));
const manifestAttachments = attachmentPayloads.map(([displayName, scope, page, annotationIndex, mime, bytes, sha256, savedName], index) => ({
  index: index + 1,
  scope,
  page,
  annotationIndex,
  internalKey: displayName,
  displayName,
  mime,
  bytes,
  sha256,
  savedName,
  savedPath: `quarantine/${savedName}`,
  nameSanitized: savedName !== displayName,
}));
const attachmentEvidence = {
  source: { sha256: "attachment-source-sha", pageCount: 1 },
  expectedAttachments,
  unsafeRawPaths: [{ displayName: "../escape.exe", resolved: "/workspace/outputs/escape.exe" }],
  manifest: {
    schema: "office-kit.pdf-attachments.v1",
    source: { sha256: "attachment-source-sha" },
    attachments: manifestAttachments,
    validation: { sourceUnchanged: true, attachmentsOpenedOrExecuted: false },
  },
  manifestFile: { sha256: "attachment-manifest-sha", bytes: 2000 },
  quarantine: {
    invalid: [],
    files: manifestAttachments.map((entry) => ({ path: entry.savedName, bytes: entry.bytes, sha256: entry.sha256, flat: true })),
  },
};
const attachmentAudit = {
  status: "succeeded",
  source: { sha256: "attachment-source-sha" },
  output: { sha256: "attachment-manifest-sha" },
  provider: { actual: "pypdf", version: "6.10.0", silentFallback: false },
  savePolicy: { strategy: "read-only" },
  preflight: { probeCompleted: true, planCompleted: true },
  operation: { type: "extract-attachments" },
  validation: {
    sourceUnchanged: true,
    allHashesVerified: true,
    allPathsContained: true,
    duplicateNamesSeparated: true,
    attachmentsOpenedOrExecuted: false,
  },
};
const attachmentCommands = [
  "python pypdf_edit.py inspect inputs/source.pdf --output tmp/inspect.json",
  "python pdf_provider.py check --provider pypdf --require",
  "python pdf_provider.py plan --task extract-attachments --provider pypdf --strategy read-only",
  "python pypdf_edit.py extract-attachments inputs/source.pdf outputs/quarantine --manifest outputs/attachments.json",
  "python pdf_audit.py validate outputs/audit.json --source inputs/source.pdf --artifact outputs/attachments.json --require-operation extract-attachments",
];
const attachmentChecks = gradeAttachmentQuarantineEvidence({ evidence: attachmentEvidence, audit: attachmentAudit, commands: attachmentCommands, item: attachmentItem });
assert.equal(attachmentChecks.every((entry) => entry.passed), true);
assert.equal(summarizeCaseScore(attachmentChecks, attachmentItem.grade).rawScorePercent, 100);
const quotedAttachmentAudit = structuredClone(attachmentAudit);
delete quotedAttachmentAudit.validation.allHashesVerified;
quotedAttachmentAudit.validation.allAttachmentHashesVerified = true;
const quotedAttachmentCommands = attachmentCommands.map((command) => command
  .replace("pypdf_edit.py ", 'pypdf_edit.py" ')
  .replace("pdf_provider.py ", 'pdf_provider.py" ')
  .replace("pdf_audit.py ", 'pdf_audit.py" '));
quotedAttachmentCommands.push('/bin/zsh -lc "shasum -a 256 outputs/quarantine/*"');
const quotedAttachmentChecks = gradeAttachmentQuarantineEvidence({
  evidence: attachmentEvidence,
  audit: quotedAttachmentAudit,
  commands: quotedAttachmentCommands,
  item: attachmentItem,
});
assert.equal(quotedAttachmentChecks.every((entry) => entry.passed), true);
const escapedAttachmentEvidence = structuredClone(attachmentEvidence);
escapedAttachmentEvidence.manifest.attachments[0].savedPath = "../escape.exe";
const escapedAttachmentChecks = gradeAttachmentQuarantineEvidence({ evidence: escapedAttachmentEvidence, audit: attachmentAudit, commands: attachmentCommands, item: attachmentItem });
assert.equal(escapedAttachmentChecks.find((entry) => entry.id === "pdf-security:path-traversal-contained")?.passed, false);
const manualAttachmentChecks = gradeAttachmentQuarantineEvidence({
  evidence: attachmentEvidence,
  audit: attachmentAudit,
  commands: [...attachmentCommands, "reader = PdfReader('inputs/source.pdf'); list(reader.attachment_list)"],
  item: attachmentItem,
});
assert.equal(manualAttachmentChecks.find((entry) => entry.id === "pdf-trace:no-ad-hoc-pypdf-extraction")?.passed, false);
const executedAttachmentChecks = gradeAttachmentQuarantineEvidence({
  evidence: attachmentEvidence,
  audit: attachmentAudit,
  commands: [...attachmentCommands, "unzip outputs/quarantine/archive.zip"],
  item: attachmentItem,
});
assert.equal(executedAttachmentChecks.find((entry) => entry.id === "pdf-trace:no-payload-open-or-execution")?.passed, false);
const interpretedAttachmentChecks = gradeAttachmentQuarantineEvidence({
  evidence: attachmentEvidence,
  audit: attachmentAudit,
  commands: [...attachmentCommands, '; "$PYTHON_BIN" outputs/quarantine/escape.exe'],
  item: attachmentItem,
});
assert.equal(interpretedAttachmentChecks.find((entry) => entry.id === "pdf-trace:no-payload-open-or-execution")?.passed, false);

const activeItem = cases.find((item) => item.id === "pdf-active-content-public-sanitize");
const activeTerms = activeItem.grade.machine.residueTerms;
const emptyTermCounts = Object.fromEntries(activeTerms.map((term) => [term, 0]));
const presentTermCounts = Object.fromEntries(activeTerms.map((term) => [term, 1]));
const activeEvidence = {
  source: {
    sha256: "active-source-sha",
    pageCount: 1,
    pages: [{ page: 1, width: 612, height: 792, rotation: 0 }],
    termCounts: emptyTermCounts,
    rawTermCounts: emptyTermCounts,
    decodedStreamTermCounts: emptyTermCounts,
    metadataTermCounts: emptyTermCounts,
    decodedStreamErrors: [],
  },
  output: {
    sha256: "active-output-sha",
    pageCount: 1,
    pages: [{ page: 1, width: 612, height: 792, rotation: 0 }],
    termCounts: emptyTermCounts,
    rawTermCounts: emptyTermCounts,
    decodedStreamTermCounts: emptyTermCounts,
    metadataTermCounts: emptyTermCounts,
    decodedStreamErrors: [],
    startxrefCount: 1,
    eofCount: 1,
  },
  sourceStructure: {
    structuralNameCounts: { "/AA": 1, "/EmbeddedFiles": 1, "/JS": 1, "/JavaScript": 1, "/Launch": 0, "/OpenAction": 1 },
    actionTypeCounts: { "/JavaScript": 2, "/Launch": 1, "/SubmitForm": 1 },
    attachments: [{ name: "internal.txt" }],
    attachmentTermCounts: emptyTermCounts,
    structureTermCounts: presentTermCounts,
    commentAnnotations: [{ subtype: "/Text" }],
    populatedWidgets: [{ name: "reviewer", values: { "/V": "Private Person" } }],
    personalMetadata: { "/Author": "Private Person" },
  },
  outputStructure: {
    structuralNameCounts: { "/AA": 0, "/EmbeddedFiles": 0, "/JS": 0, "/JavaScript": 0, "/Launch": 0, "/OpenAction": 0 },
    actionTypeCounts: { "/JavaScript": 0, "/Launch": 0, "/SubmitForm": 0 },
    attachments: [],
    attachmentTermCounts: emptyTermCounts,
    structureTermCounts: emptyTermCounts,
    commentAnnotations: [],
    populatedWidgets: [],
    personalMetadata: {},
  },
  originalPrefixPreserved: false,
  visual: {
    renderer: "poppler-pdftoppm",
    sourcePageCount: 1,
    outputPageCount: 1,
    allowedMasks: [{ page: 1, bboxPx: [136, 288, 512, 352] }],
    pages: [{ page: 1, sameDimensions: true, nonBlank: true, changedPixelsBBox: [154, 300, 220, 325], changedOutsideAllowedMasksBBox: null, changedOnlyWithinAllowedMasks: true }],
  },
};
const activeAudit = {
  status: "succeeded",
  source: { sha256: "active-source-sha" },
  output: { sha256: "active-output-sha" },
  provider: { actual: "pymupdf", version: "1.27.2.3", silentFallback: false },
  savePolicy: { strategy: "sanitize" },
  preflight: { probeCompleted: true, planCompleted: true },
  operation: [{ type: "scrub" }, { type: "active_content_cleanup" }],
  validation: { residue: { ok: true }, render: { pages: 1 } },
};
const activeCommands = [
  "python pymupdf_edit.py probe --accept-license agpl",
  "python pdf_provider.py plan --task sanitize --provider pymupdf --strategy sanitize",
  "python pymupdf_edit.py edit input.pdf output.pdf --strategy sanitize",
  "python residue_scan.py output.pdf --require-inert",
  "pdftoppm -png output.pdf page",
  "python pdf_audit.py validate audit.json",
];
const activeChecks = gradeActiveContentSanitizeEvidence({ evidence: activeEvidence, audit: activeAudit, commands: activeCommands, item: activeItem });
assert.equal(activeChecks.every((entry) => entry.passed), true);
assert.equal(summarizeCaseScore(activeChecks, activeItem.grade).rawScorePercent, 100);
const helpBeforeMutationChecks = gradeActiveContentSanitizeEvidence({
  evidence: activeEvidence,
  audit: activeAudit,
  commands: [
    "python pymupdf_edit.py probe --help && python pymupdf_edit.py inspect --help && python pymupdf_edit.py edit --help",
    "python pdf_provider.py plan --help",
    ...activeCommands,
  ],
  item: activeItem,
});
assert.equal(helpBeforeMutationChecks.find((entry) => entry.id === "pdf-trace:probe-plan-before-mutation")?.passed, true);
const editBeforePlanChecks = gradeActiveContentSanitizeEvidence({
  evidence: activeEvidence,
  audit: activeAudit,
  commands: [activeCommands[2], activeCommands[0], activeCommands[1], ...activeCommands.slice(3)],
  item: activeItem,
});
assert.equal(editBeforePlanChecks.find((entry) => entry.id === "pdf-trace:probe-plan-before-mutation")?.passed, false);
const residualActionEvidence = structuredClone(activeEvidence);
residualActionEvidence.outputStructure.actionTypeCounts["/SubmitForm"] = 1;
const residualActionChecks = gradeActiveContentSanitizeEvidence({ evidence: residualActionEvidence, audit: activeAudit, commands: activeCommands, item: activeItem });
assert.equal(residualActionChecks.find((entry) => entry.id === "pdf-security:active-names-absent")?.passed, false);
const visualDriftEvidence = structuredClone(activeEvidence);
visualDriftEvidence.visual.pages[0].changedOutsideAllowedMasksBBox = [20, 20, 40, 40];
visualDriftEvidence.visual.pages[0].changedOnlyWithinAllowedMasks = false;
const visualDriftChecks = gradeActiveContentSanitizeEvidence({ evidence: visualDriftEvidence, audit: activeAudit, commands: activeCommands, item: activeItem });
assert.equal(visualDriftChecks.find((entry) => entry.id === "pdf-visual:ordinary-content-stable")?.passed, false);
const activeBypassChecks = gradeActiveContentSanitizeEvidence({ evidence: activeEvidence, audit: activeAudit, commands: [...activeCommands, "doc.xref_set_key(1, 'AA', 'null')"], item: activeItem });
assert.equal(activeBypassChecks.find((entry) => entry.id === "pdf-trace:no-content-stream-bypass")?.passed, false);
const preMutationQaOnly = [
  activeCommands[0],
  activeCommands[1],
  "python residue_scan.py input.pdf --require-inert",
  "pdftoppm -png input.pdf source-page",
  "python pdf_audit.py validate preflight.json",
  activeCommands[2],
];
const preMutationQaChecks = gradeActiveContentSanitizeEvidence({ evidence: activeEvidence, audit: activeAudit, commands: preMutationQaOnly, item: activeItem });
for (const id of ["pdf-trace:post-mutation-residue-scan", "pdf-trace:post-mutation-poppler-render", "pdf-trace:audit-byte-validation"]) {
  assert.equal(preMutationQaChecks.find((entry) => entry.id === id)?.passed, false, `${id} must require evidence after mutation`);
}

const traceCommands = extractCompletedCommands([
  JSON.stringify({ type: "item.started", item: { id: "one", type: "command_execution", command: "ignored-started-command" } }),
  JSON.stringify({ type: "item.completed", item: { id: "one", type: "command_execution", command: "echo safe", aggregated_output: "doc.update_stream(xref, bytes)" } }),
].join("\n"));
assert.deepEqual(traceCommands, ["echo safe"]);

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-agent-eval-test-"));
try {
  const missingPython = spawnSync(process.execPath, ["scripts/run-agent-evals.mjs", "prepare", "pdf-bounded-contract-id-replace", "--run-root", path.join(temporary, "missing-python")], {
    cwd: path.resolve(import.meta.dirname, ".."),
    encoding: "utf8",
    env: { ...process.env, OFFICE_KIT_AGENT_EVAL_PYTHON: process.execPath, OFFICE_KIT_PDF_PROVIDER_PYTHON: "" },
  });
  assert.notEqual(missingPython.status, 0);
  assert.match(missingPython.stderr, /Generated Agent eval fixtures require a Python environment with reportlab and pypdf/);
  assert.doesNotMatch(missingPython.stderr, /def base_page|contract-five-page/);
  const preparedSkillMode = (await fs.stat(path.join(temporary, "missing-python", "pdf-bounded-contract-id-replace", "candidate-trial-1", "workspace", ".agents", "skills", "pdf", "SKILL.md"))).mode;
  assert.equal(preparedSkillMode & 0o222, 0);

  const removable = path.join(temporary, "removable", "nested");
  await fs.mkdir(removable, { recursive: true });
  await fs.writeFile(path.join(removable, "locked.txt"), "locked");
  await makeReadOnly(path.join(temporary, "removable"));
  await removePreparedTree(path.join(temporary, "removable"));
  await assert.rejects(() => fs.access(path.join(temporary, "removable")), /ENOENT/);

  const item = cases.find((candidate) => candidate.id === "pdf-richmedia-opaque-preservation");
  const workspace = path.join(temporary, "workspace");
  const evaluator = path.join(temporary, "evaluator");
  const credentials = path.join(workspace, "inputs", "credentials");
  await fs.mkdir(path.join(workspace, "outputs"), { recursive: true });
  await fs.mkdir(credentials, { recursive: true });
  await fs.mkdir(path.join(workspace, ".agents", "skills", "pdf"), { recursive: true });
  await fs.mkdir(path.join(workspace, "node_modules", "office-kit"), { recursive: true });
  await fs.mkdir(evaluator, { recursive: true });
  await fs.writeFile(path.join(credentials, "signer.pem"), "test-only-key-material");
  await fs.writeFile(path.join(workspace, "PROMPT.md"), "test prompt");
  await fs.writeFile(path.join(workspace, "package.json"), "{}");
  await fs.writeFile(path.join(workspace, "package-lock.json"), "{}");
  await fs.writeFile(path.join(workspace, ".agents", "skills", "pdf", "SKILL.md"), "test skill");
  const installedModule = path.join(workspace, "node_modules", "office-kit", "index.mjs");
  await fs.writeFile(installedModule, "export const version = 1;\n");
  const before = await fingerprintPath(credentials);
  assert.match(before, /^tree:[0-9a-f]{64}$/);

  const workspaceHashes = {};
  for (const relative of ["PROMPT.md", "package.json", "package-lock.json", ".agents", "node_modules"]) {
    workspaceHashes[relative] = await fingerprintPath(path.join(workspace, relative));
  }

  await fs.writeFile(path.join(workspace, "outputs", "audit.json"), JSON.stringify({ status: "failed_closed", delivered_modified_pdf: false }));
  await fs.writeFile(path.join(evaluator, "run.json"), JSON.stringify({ oracleSha256: oracleFingerprint(item), workspaceHashes, inputHashes: { "inputs/credentials": before } }));
  await fs.writeFile(path.join(evaluator, "exit.json"), JSON.stringify({ status: 0, signal: null }));
  await fs.writeFile(path.join(evaluator, "final.txt"), "Opaque RichMedia runtime behavior cannot be proven; failed closed.");
  await fs.writeFile(path.join(evaluator, "trace.jsonl"), "{\"type\":\"result\"}\n");

  const prepared = { workspace, evaluator };
  const refusal = await scorePrepared(item, prepared);
  assert.equal(refusal.observedOutcome, "safe-refusal");
  assert.equal(refusal.hardGatesPassed, true);

  await fs.chmod(path.join(credentials, "signer.pem"), 0o600);
  const alteredInputMode = await scorePrepared(item, prepared);
  assert.equal(alteredInputMode.checks.find((check) => check.id === "source-immutable:inputs/credentials")?.passed, false);
  await fs.chmod(path.join(credentials, "signer.pem"), 0o644);

  await fs.writeFile(installedModule, "export const version = 2;\n");
  const alteredRuntime = await scorePrepared(item, prepared);
  assert.equal(alteredRuntime.checks.find((check) => check.id === "workspace-immutable:node_modules")?.passed, false);
  assert.equal(alteredRuntime.hardGatesPassed, false);
  await fs.writeFile(installedModule, "export const version = 1;\n");

  await fs.writeFile(path.join(credentials, "signer.pem"), "mutated-key-material");
  const mutated = await scorePrepared(item, prepared);
  assert.equal(mutated.checks.find((check) => check.id === "source-immutable:inputs/credentials")?.passed, false);
  assert.equal(mutated.hardGatesPassed, false);
} finally {
  await removePreparedTree(temporary);
}

console.log("agent eval suite smoke ok");
