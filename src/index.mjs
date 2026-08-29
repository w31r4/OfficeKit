import { DocumentFile, DocumentModel } from "./document/index.mjs";
import { PdfArtifact, PdfFile } from "./pdf/index.mjs";
import { Range, SpreadsheetFile, Workbook, Worksheet, WorksheetDataTableCollection } from "./spreadsheet/index.mjs";
import { queryHelpRecords } from "./help/index.mjs";
import { createArtifactVisualQaApi } from "./qa/artifact-visual.mjs";
import { ndjson, verificationIssue, verificationResult } from "./shared/inspection.mjs";

export { FileBlob } from "./shared/file-blob.mjs";
export { reviewArtifact } from "./review/index.mjs";
export {
  clearOfficeFontDesignMetrics,
  registerScopedOfficeFontDesignMetrics,
  resolveOfficeFontDesignMetrics,
  setOfficeFontDesignMetrics,
  skiaPaintBaselineCompensationPx,
} from "./shared/font-design-metrics.mjs";
export { DocumentFile, DocumentModel, PdfArtifact, PdfFile, Range, SpreadsheetFile, Workbook, Worksheet, WorksheetDataTableCollection };

function inferArtifactKind(artifact) {
  if (artifact instanceof Workbook) return "workbook";
  if (artifact instanceof DocumentModel) return "document";
  if (artifact instanceof PdfArtifact) return "pdf";
  // The Presentation object model is an internal codec implementation in 2.0.
  // Keep generic QA useful for repository-internal presentation values without
  // importing or re-exporting that authoring surface from the public root.
  if (artifact?.slides?.items && artifact?.slideSize && typeof artifact?.toProto === "function") return "presentation";
  return "unknown";
}

export function verifyArtifact(artifact, options = {}) {
  if (!artifact || typeof artifact.verify !== "function") {
    return verificationResult("unknown", [verificationIssue("unknown", "unsupportedArtifact", "Artifact does not expose a verify() method.")], options);
  }
  return artifact.verify(options);
}

const artifactVisualQaApi = createArtifactVisualQaApi({ inferArtifactKind });

export async function renderArtifact(artifact, options = {}) {
  return artifactVisualQaApi.renderArtifact(artifact, options);
}

export async function visualQaArtifact(artifact, options = {}) {
  return artifactVisualQaApi.visualQaArtifact(artifact, options);
}

export function helpArtifact(artifactOrKind = "*", query = "*", options = {}) {
  const artifactKind = typeof artifactOrKind === "string" ? artifactOrKind : inferArtifactKind(artifactOrKind);
  return ndjson(queryHelpRecords(artifactKind, query, options), options.maxChars ?? Infinity);
}
