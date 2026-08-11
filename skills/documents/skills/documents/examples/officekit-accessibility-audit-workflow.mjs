import crypto from "node:crypto";
import { constants as FS_CONSTANTS } from "node:fs";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { DocumentFile, FileBlob } from "office-kit";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function packageVersion() {
  const entry = require.resolve("office-kit");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

function requiredPath(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(label + " must be a non-empty path.");
  return path.resolve(value.trim());
}

function boundedMaxChars(value) {
  const numeric = value === undefined ? 200_000 : Number(value);
  if (!Number.isSafeInteger(numeric) || numeric < 1 || numeric > 2_000_000) {
    throw new RangeError("maxChars must be an integer from 1 through 2000000.");
  }
  return numeric;
}

async function assertAbsent(filePath, label) {
  try {
    await fs.lstat(filePath);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  throw new Error(label + " already exists; refusing to overwrite it.");
}

async function publishNoReplace(temporaryPath, finalPath) {
  try {
    await fs.link(temporaryPath, finalPath);
  } catch (error) {
    if (error?.code === "EEXIST") throw new Error("reportPath already exists; refusing to overwrite it.");
    if (!["EPERM", "EXDEV", "ENOTSUP", "EOPNOTSUPP"].includes(error?.code)) throw error;
    try {
      await fs.copyFile(temporaryPath, finalPath, FS_CONSTANTS.COPYFILE_EXCL);
    } catch (copyError) {
      if (copyError?.code === "EEXIST") throw new Error("reportPath already exists; refusing to overwrite it.");
      throw copyError;
    }
  }
  await fs.rm(temporaryPath, { force: true });
}

export async function auditDocxAccessibility({ inputPath, reportPath, maxChars } = {}) {
  const sourcePath = requiredPath(inputPath, "inputPath");
  const finalReportPath = requiredPath(reportPath, "reportPath");
  const reportLimit = boundedMaxChars(maxChars);
  if (sourcePath === finalReportPath) throw new Error("reportPath must be distinct from inputPath.");

  const sourceStat = await fs.lstat(sourcePath);
  if (!sourceStat.isFile() || sourceStat.isSymbolicLink()) {
    throw new Error("inputPath must be a regular, non-symlink DOCX file.");
  }
  await fs.mkdir(path.dirname(finalReportPath), { recursive: true });
  await assertAbsent(finalReportPath, "reportPath");

  const source = await fs.readFile(sourcePath);
  const sourceSha256 = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, {
    type: DOCX_MIME,
    name: path.basename(sourcePath),
  }));
  const accessibility = document.auditAccessibility({ maxChars: reportLimit });
  const verification = document.verify({ maxChars: reportLimit });
  const sourceAfter = await fs.readFile(sourcePath);
  if (sha256(sourceAfter) !== sourceSha256) throw new Error("Source DOCX changed during the read-only accessibility audit.");

  const report = {
    schema: "office-kit.docx-accessibility-audit.v1",
    status: "succeeded",
    source: {
      path: sourcePath,
      sha256: sourceSha256,
      bytes: source.length,
      immutable: true,
    },
    provider: {
      requested: "office-kit",
      actual: "office-kit",
      version: await packageVersion(),
      silentFallback: false,
    },
    savePolicy: {
      strategy: "none",
      sourceMutation: false,
      artifactProduced: false,
    },
    operation: {
      type: "document-accessibility-audit",
      maxChars: reportLimit,
    },
    accessibility,
    validation: {
      sourceUnchanged: true,
      documentVerify: {
        ok: verification.ok,
        issueCount: verification.issues.length,
      },
    },
    boundaries: {
      headingIntent: "modeled-outline-and-heading-style-check-only",
      tableAndLinkPurpose: "manual-author-review",
      opaqueContent: "manual-source-or-native-host-review",
      conformanceClaimed: false,
      note: "The machine result covers modeled headings, images, tables, and hyperlinks only; it is not Word Accessibility Checker, WCAG, or whole-document conformance evidence.",
    },
  };

  const temporaryPath = path.join(
    path.dirname(finalReportPath),
    `.${path.basename(finalReportPath)}.${process.pid}.${crypto.randomBytes(8).toString("hex")}.tmp`,
  );
  try {
    const handle = await fs.open(temporaryPath, "wx", 0o600);
    try {
      await handle.writeFile(JSON.stringify(report, null, 2) + "\n");
      await handle.sync();
    } finally {
      await handle.close();
    }
    await publishNoReplace(temporaryPath, finalReportPath);
  } catch (error) {
    await fs.rm(temporaryPath, { force: true }).catch(() => {});
    throw error;
  }

  return { reportPath: finalReportPath, report };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const [inputPath, reportPath, maxChars] = process.argv.slice(2);
  const result = await auditDocxAccessibility({ inputPath, reportPath, maxChars });
  console.log(JSON.stringify({
    reportPath: result.reportPath,
    sourceSha256: result.report.source.sha256,
    machineCheckPassed: result.report.accessibility.machineCheckPassed,
    manualReviewRequired: result.report.accessibility.manualReviewRequired,
  }));
}
