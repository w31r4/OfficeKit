import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import sharp from "sharp";
import JSZip from "jszip";

import {
  DocumentFile,
  DocumentModel,
  SpreadsheetFile,
  Workbook,
} from "office-kit";

const packageRoot = path.resolve(import.meta.dirname, "..");
const creatorPath = path.join(
  packageRoot,
  "skills/template-creator/skills/template-creator/scripts/create-template-skill.mjs",
);
const presentationCreatorPath = path.join(
  packageRoot,
  "skills/presentation-template-creator/skills/presentation-template-creator/scripts/package-presentation-template.mjs",
);

try {
  await Promise.all([fs.access(creatorPath), fs.access(presentationCreatorPath)]);
} catch (error) {
  if (error?.code === "ENOENT") {
    console.log("template creator smoke skipped: repository-only skills are not packaged");
    process.exit(0);
  }
  throw error;
}

const tempRoot = await fs.mkdtemp(
  path.join(os.tmpdir(), "office-kit-template-creator-"),
);
const home = path.join(tempRoot, "neutral-home");
const fixturesDirectory = path.join(tempRoot, "fixtures");

function runScript(scriptPath, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [scriptPath, ...args], {
      env: { ...process.env, OFFICE_KIT_HOME: home },
      stdio: ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });
    child.once("error", reject);
    child.once("close", (code) => {
      resolve({ code, stderr, stdout });
    });
  });
}

function runCreator(args) {
  return runScript(creatorPath, args);
}

function runPresentationCreator(args) {
  return runScript(presentationCreatorPath, args);
}

async function runSuccessfulCreator(args) {
  const result = await runCreator(args);
  if (result.code !== 0) {
    throw new Error(`Template creator failed (${result.code}): ${result.stderr}`);
  }
  try {
    return JSON.parse(result.stdout);
  } catch (error) {
    throw new Error(`Template creator did not return JSON: ${result.stdout}\n${error}`);
  }
}

async function runSuccessfulPresentationCreator(args) {
  const result = await runPresentationCreator([...args, "--json"]);
  assert.equal(result.code, 0, result.stderr);
  return JSON.parse(result.stdout);
}

async function assertRejectedOfficeReference(referencePath, label) {
  const result = await runCreator([
    "--reference-path", referencePath,
    "--preview-path", path.join(fixturesDirectory, "preview.png"),
    "--display-name", `${label} fixture`,
    "--description", "This malformed Office reference must not be retained.",
  ]);
  assert.notEqual(result.code, 0, `Template creator accepted ${label}.`);
  assert.match(
    result.stderr,
    /structurally valid Office Open XML package/i,
    `Template creator did not explain the Office-package rejection for ${label}: ${result.stderr}`,
  );
}

async function assertBytesEqual(actualPath, expectedPath, label) {
  const [actual, expected] = await Promise.all([
    fs.readFile(actualPath),
    fs.readFile(expectedPath),
  ]);
  if (!actual.equals(expected)) {
    throw new Error(`${label} did not retain the exact source bytes.`);
  }
}

async function assertGeneratedTemplate(
  result,
  {
    kind,
    referencePath,
    visualCommitment = "opinionated",
    editLevel = "copy-only",
    provenanceSource = "local-user-reference",
  },
) {
  const skillsRoot = path.join(home, "skills");
  if (path.dirname(result.skillPath) !== skillsRoot) {
    throw new Error(`Template was not written below OFFICE_KIT_HOME: ${result.skillPath}`);
  }
  if (result.kind !== kind || !result.skillName.startsWith("artifact-template-")) {
    throw new Error(`Unexpected template result: ${JSON.stringify(result)}`);
  }

  const extension = path.extname(referencePath).toLowerCase();
  const referenceName = `reference${extension}`;
  const agentPath = path.join(result.skillPath, "agents/agent.yaml");
  const legacyAgentPath = path.join(result.skillPath, "agents/openai.yaml");
  const sidecarPath = path.join(result.skillPath, "artifact-template.json");
  const skillPath = path.join(result.skillPath, "SKILL.md");
  const previewPath = path.join(result.skillPath, "assets/preview.png");
  const retainedReferencePath = path.join(result.skillPath, "assets", referenceName);

  await Promise.all([
    fs.access(agentPath),
    fs.access(sidecarPath),
    fs.access(skillPath),
    fs.access(previewPath),
    fs.access(retainedReferencePath),
  ]);
  if (await fs.access(legacyAgentPath).then(() => true).catch(() => false)) {
    throw new Error("Generated template retained the legacy agent metadata filename.");
  }

  const [sidecar, skillText, previewBytes, retainedReferenceBytes] = await Promise.all([
    fs.readFile(sidecarPath, "utf8").then(JSON.parse),
    fs.readFile(skillPath, "utf8"),
    fs.readFile(previewPath),
    fs.readFile(retainedReferencePath),
  ]);
  if (
    result.schemaVersion !== 2 ||
    sidecar.schemaVersion !== 2 ||
    sidecar.id !== result.skillName ||
    sidecar.displayName !== result.displayName ||
    sidecar.kind !== kind ||
    sidecar.reference !== `assets/${referenceName}` ||
    sidecar.preview !== "assets/preview.png" ||
    !Array.isArray(sidecar.useWhen) ||
    sidecar.useWhen.length === 0 ||
    sidecar.visualCommitment !== visualCommitment ||
    sidecar.editProfile?.level !== editLevel ||
    (editLevel === "copy-only" && sidecar.editProfile?.verifiedOperations?.length !== 0) ||
    sidecar.provenance?.license !== "user-provided" ||
    sidecar.provenance?.source !== provenanceSource ||
    sidecar.provenance?.referenceSha256 !== sha256(retainedReferenceBytes) ||
    sidecar.provenance?.previewSha256 !== sha256(previewBytes)
  ) {
    throw new Error(`Generated sidecar is invalid: ${JSON.stringify(sidecar)}`);
  }
  if (/codex|openai|plugin:\/\//iu.test(skillText)) {
    throw new Error("Generated template skill contains a product-specific reference.");
  }

  await Promise.all([
    assertBytesEqual(retainedReferencePath, referencePath, `${kind} reference`),
    assertBytesEqual(previewPath, path.join(fixturesDirectory, "preview.png"), `${kind} preview`),
  ]);
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

async function assertNoTransactionalResidue() {
  const [skillEntries, homeEntries] = await Promise.all([
    fs.readdir(path.join(home, "skills")),
    fs.readdir(home),
  ]);
  const skillResidue = skillEntries.filter(
    (entry) => entry.includes("-stage-") || entry.includes(".backup-"),
  );
  const lockResidue = homeEntries.filter(
    (entry) =>
      entry.startsWith(".artifact-template-write-lock.pending-") ||
      entry.startsWith(".artifact-template-write-lock.stale-"),
  );
  const residue = [...skillResidue, ...lockResidue];
  if (residue.length > 0) {
    throw new Error(`Template creator left transactional residue: ${residue.join(", ")}`);
  }
}

async function writePresentationFixture(filePath, slideCount) {
  assert.equal(slideCount, 1);
  await fs.copyFile(
    path.join(packageRoot, "skills/presentation-template-library/skills/artifact-template-evidence-ledger/assets/references/reference.pptx"),
    filePath,
  );
}

async function writePngFixture(filePath, background = { r: 15, g: 118, b: 110, alpha: 1 }) {
  await sharp({
    create: {
      width: 320,
      height: 180,
      channels: 4,
      background,
    },
  }).png().toFile(filePath);
}

async function writeDocumentFixture(filePath) {
  const document = DocumentModel.create();
  const file = await DocumentFile.exportDocx(document);
  await file.save(filePath);
}

async function writeSpreadsheetFixture(filePath) {
  const workbook = Workbook.create();
  const worksheet = workbook.worksheets.add("Fixture");
  worksheet.getRange("A1:B2").values = [
    ["Kind", "Value"],
    ["Fixture", 1],
  ];
  const file = await SpreadsheetFile.exportXlsx(workbook);
  await file.save(filePath);
}

try {
  await fs.mkdir(fixturesDirectory, { recursive: true });
  const pptxPath = path.join(fixturesDirectory, "reference.pptx");
  const docxPath = path.join(fixturesDirectory, "reference.docx");
  const xlsxPath = path.join(fixturesDirectory, "reference.xlsx");
  const renamedDocxPath = path.join(fixturesDirectory, "renamed-not-office.docx");
  const renamedXlsxPath = path.join(fixturesDirectory, "renamed-not-office.xlsx");
  const crossFamilyDocxPath = path.join(fixturesDirectory, "presentation-renamed.docx");
  const invalidRootRelationshipDocxPath = path.join(fixturesDirectory, "invalid-root-relationship.docx");
  const previewPath = path.join(fixturesDirectory, "preview.png");

  await Promise.all([
    writePresentationFixture(pptxPath, 1),
    writeSpreadsheetFixture(xlsxPath),
    writeDocumentFixture(docxPath),
    fs.writeFile(renamedDocxPath, "not an Office package\n", "utf8"),
    fs.writeFile(renamedXlsxPath, "not an Office package\n", "utf8"),
    writePngFixture(previewPath),
    fs.mkdir(home, { recursive: true }),
  ]);
  await fs.copyFile(pptxPath, crossFamilyDocxPath);
  const invalidRootRelationshipZip = await JSZip.loadAsync(
    await fs.readFile(docxPath),
  );
  const rootRelationships = await invalidRootRelationshipZip.file("_rels/.rels").async("text");
  invalidRootRelationshipZip.file(
  "_rels/.rels",
  rootRelationships.replace(
    /Target="\/?word\/document\.xml"/u,
    'Target="/word/not-the-main-document.xml"',
  ),
  );
  await fs.writeFile(
    invalidRootRelationshipDocxPath,
    await invalidRootRelationshipZip.generateAsync({ type: "nodebuffer" }),
  );

  for (const [referencePath, label] of [
    [renamedDocxPath, "renamed DOCX text"],
    [renamedXlsxPath, "renamed XLSX text"],
    [crossFamilyDocxPath, "PPTX bytes renamed as DOCX"],
    [invalidRootRelationshipDocxPath, "DOCX with invalid root relationship"],
  ]) {
    await assertRejectedOfficeReference(referencePath, label);
  }
  assert.deepEqual(
    await fs.readdir(home),
    [],
    "Template creator must reject an invalid Office reference before it creates locks or writes a template tree.",
  );

  await fs.writeFile(
    path.join(home, ".artifact-template-write-lock"),
    "999999999\n",
  );

  const genericPptx = await runCreator([
    "--reference-path", pptxPath,
    "--preview-path", previewPath,
    "--display-name", "Presentation fixture",
    "--description", "This input must route to the presentation specialist.",
  ]);
  assert.notEqual(genericPptx.code, 0);
  assert.match(genericPptx.stderr, /presentation-template-creator/);

  const presentationInputRoot = path.join(fixturesDirectory, "presentation-style");
  const presentationOutputRoot = path.join(home, "presentation-skills");
  await fs.mkdir(presentationInputRoot, { recursive: true });
  const examplePaths = Array.from({ length: 4 }, (_, index) =>
    path.join(presentationInputRoot, `example-${index + 1}.png`));
  await Promise.all(examplePaths.map((entry, index) => writePngFixture(entry, {
    r: 30 + index * 35,
    g: 70 + index * 20,
    b: 120 + index * 15,
    alpha: 1,
  })));
  const guidePath = path.join(presentationInputRoot, "guide.md");
  await fs.writeFile(
    guidePath,
    "# Visual direction\n\nUse a disciplined editorial rhythm, a strong evidence hierarchy, restrained color roles, and varied page silhouettes. Build every slide freely for the current content. Treat the examples as visual evidence rather than geometry to trace.\n",
  );
  const referenceProgramPath = path.join(presentationInputRoot, "reference.ppj");
  await fs.writeFile(referenceProgramPath, `${JSON.stringify({
    schema: "office-kit/ppj/v1",
    meta: { id: "template-fixture", title: "Template fixture", language: "en-US", version: 1 },
    intent: {},
    design: {},
    pages: [],
  }, null, 2)}\n`);
  const specPath = path.join(presentationInputRoot, "spec.json");
  const presentationSpec = {
    id: "artifact-template-presentation-fixture",
    displayName: "Presentation Fixture",
    description: "Use an original editorial presentation style with free composition.",
    guidePath,
    useWhen: ["editorial evidence presentation"],
    avoidWhen: ["playful consumer launch"],
    audiences: ["executives"],
    contentShapes: ["evidence narrative"],
    visualTraits: {
      tone: ["disciplined", "editorial"],
      density: "medium",
      colorMode: "light",
      structure: ["asymmetric", "claim led"],
    },
    visualCommitment: "opinionated",
    examples: examplePaths.map((entry, index) => ({
      path: entry,
      role: ["cover", "analysis", "data", "closing"][index],
    })),
    referenceProgram: {
      path: referenceProgramPath,
      license: "AGPL-3.0-or-later",
      source: "OfficeKit original clean-room fixture",
    },
    referencePptx: {
      path: pptxPath,
      license: "AGPL-3.0-or-later",
      source: "Compiled OfficeKit fixture",
    },
    provenance: {
      license: "user-provided",
      source: "unrelated calibration pages created for this smoke",
    },
  };
  await fs.writeFile(specPath, `${JSON.stringify(presentationSpec, null, 2)}\n`);
  const presentationTemplate = await runSuccessfulPresentationCreator([
    "--spec", specPath,
    "--output-root", presentationOutputRoot,
  ]);
  assert.equal(presentationTemplate.schemaVersion, 3);
  assert.equal(presentationTemplate.updated, false);
  assert.equal(presentationTemplate.examplePaths.length, 4);
  const presentationMetadata = JSON.parse(
    await fs.readFile(path.join(presentationTemplate.skillPath, "artifact-template.json"), "utf8"),
  );
  assert.equal(presentationMetadata.kind, "presentation");
  assert.equal(presentationMetadata.schemaVersion, 3);
  assert.equal(Object.hasOwn(presentationMetadata, "reference"), false);
  assert.equal(Object.hasOwn(presentationMetadata, "editProfile"), false);
  assert.equal(presentationMetadata.referenceProgram.path, "assets/references/reference.ppj");
  assert.equal(presentationMetadata.referencePptx.path, "assets/references/reference.pptx");
  assert.equal(presentationMetadata.referenceProgram.sha256, sha256(await fs.readFile(referenceProgramPath)));
  assert.equal(presentationMetadata.referencePptx.sha256, sha256(await fs.readFile(pptxPath)));
  assert.equal(presentationTemplate.referenceProgramPath, path.join(presentationTemplate.skillPath, "assets", "references", "reference.ppj"));
  assert.equal(presentationTemplate.referencePptxPath, path.join(presentationTemplate.skillPath, "assets", "references", "reference.pptx"));
  assert.deepEqual(
    (await fs.readdir(presentationTemplate.skillPath)).sort(),
    ["SKILL.md", "agents", "artifact-template.json", "assets"],
  );
  const updatedPresentationTemplate = await runSuccessfulPresentationCreator([
    "--spec", specPath,
    "--output-root", presentationOutputRoot,
    "--expected-sha256", presentationTemplate.sidecarSha256,
  ]);
  assert.equal(updatedPresentationTemplate.updated, true);
  assert.equal(updatedPresentationTemplate.sidecarSha256, presentationTemplate.sidecarSha256);

  const docxTemplate = await runSuccessfulCreator([
    "--reference-path", docxPath,
    "--preview-path", previewPath,
    "--display-name", "Document fixture",
    "--description", "Create documents from the fixture layout.",
  ]);
  await assertGeneratedTemplate(docxTemplate, { kind: "document", referencePath: docxPath });

  const xlsxTemplate = await runSuccessfulCreator([
    "--reference-path", xlsxPath,
    "--preview-path", previewPath,
    "--display-name", "Spreadsheet fixture",
    "--description", "Create spreadsheets from the fixture layout.",
  ]);
  await assertGeneratedTemplate(xlsxTemplate, { kind: "spreadsheet", referencePath: xlsxPath });

  const longDescription =
    "Create a detailed planning workbook for a recurring operating review with assumptions, " +
    "owners, milestones, risks, decisions, supporting evidence, and a concise executive summary.";
  const longDescriptionTemplate = await runSuccessfulCreator([
    "--reference-path", xlsxPath,
    "--preview-path", previewPath,
    "--display-name", "Long description fixture",
    "--description", longDescription,
  ]);
  await assertGeneratedTemplate(longDescriptionTemplate, {
    kind: "spreadsheet",
    referencePath: xlsxPath,
  });
  const longDescriptionMetadata = JSON.parse(
    await fs.readFile(
      path.join(longDescriptionTemplate.skillPath, "artifact-template.json"),
      "utf8",
    ),
  );
  assert.ok(longDescriptionMetadata.useWhen[0].length <= 120);
  assert.equal(longDescriptionMetadata.useWhen[0], longDescription.slice(0, 120).trimEnd());
  assert.match(
    await fs.readFile(path.join(longDescriptionTemplate.skillPath, "SKILL.md"), "utf8"),
    new RegExp(longDescription.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&")),
  );

  const kindChange = await runCreator([
    "--mode", "update",
    "--skill-name", docxTemplate.skillName,
    "--reference-path", xlsxPath,
    "--preview-path", previewPath,
    "--display-name", "Document fixture",
    "--description", "Attempt to change the document fixture kind.",
  ]);
  if (kindChange.code === 0) {
    throw new Error("Template creator accepted an artifact-kind-changing update.");
  }
  await assertGeneratedTemplate(docxTemplate, { kind: "document", referencePath: docxPath });

  await assertNoTransactionalResidue();

  const linkedTemplatePath = path.join(home, "skills", "artifact-template-linked");
  const outsidePath = path.join(tempRoot, "outside-assets");
  await fs.mkdir(linkedTemplatePath, { recursive: true });
  await fs.mkdir(outsidePath, { recursive: true });
  await fs.writeFile(path.join(linkedTemplatePath, "artifact-template.json"), JSON.stringify({ schemaVersion: 1, kind: "document" }));
  await fs.symlink(outsidePath, path.join(linkedTemplatePath, "assets"), "dir");
  const linkedUpdate = await runCreator([
    "--mode", "update",
    "--skill-name", "artifact-template-linked",
    "--reference-path", docxPath,
    "--preview-path", previewPath,
    "--display-name", "Linked template",
    "--description", "This update must fail before following a template-owned symlink.",
  ]);
  if (linkedUpdate.code === 0 || !/reject symbolic links/i.test(linkedUpdate.stderr)) {
    throw new Error(`Template creator did not fail closed on a template-owned symlink: ${linkedUpdate.stderr}`);
  }
  if ((await fs.readdir(outsidePath)).length !== 0) {
    throw new Error("Template creator wrote through a template-owned symbolic link.");
  }
  await fs.rm(linkedTemplatePath, { recursive: true, force: true });
  await assertNoTransactionalResidue();

  const oversizedReferencePath = path.join(fixturesDirectory, "oversized.docx");
  const oversizedHandle = await fs.open(oversizedReferencePath, "w");
  await oversizedHandle.truncate(512 * 1024 * 1024 + 1);
  await oversizedHandle.close();
  const oversized = await runCreator([
    "--reference-path", oversizedReferencePath,
    "--preview-path", previewPath,
    "--display-name", "Oversized reference",
    "--description", "This input must be rejected before it is copied.",
  ]);
  if (oversized.code === 0 || !/input budget/i.test(oversized.stderr)) {
    throw new Error(`Template creator did not enforce its input budget: ${oversized.stderr}`);
  }
  await assertNoTransactionalResidue();

  const activeLockPath = path.join(home, ".artifact-template-write-lock");
  await fs.writeFile(activeLockPath, `${process.pid}\n`);
  const activeLock = await runCreator([
    "--reference-path", docxPath,
    "--preview-path", previewPath,
    "--display-name", "Blocked document fixture",
    "--description", "Attempt to create while another writer owns the lock.",
  ]);
  if (activeLock.code === 0) {
    throw new Error("Template creator wrote through an active write lock.");
  }
  await fs.readFile(activeLockPath, "utf8");
  await fs.rm(activeLockPath, { force: true, recursive: true });

  if (await fs.access(path.join(home, ".artifact-template-write-lock")).then(() => true).catch(() => false)) {
    throw new Error("Template creator did not release the write lock.");
  }

  console.log("template creator smoke ok");
} finally {
  await fs.rm(tempRoot, { force: true, recursive: true });
}
