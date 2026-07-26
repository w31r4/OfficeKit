import crypto from "node:crypto";
import { constants as FS_CONSTANTS } from "node:fs";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { FileBlob, PresentationFile } from "office-kit";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const MAX_ADVANCE_AFTER_MS = 86_400_000;
const EFFECTS = new Set(["fade", "push"]);
const SPEEDS = new Set(["slow", "medium", "fast"]);
const DIRECTIONS = new Set(["left", "up", "right", "down"]);
const TRANSITION_KEYS = new Set(["effect", "direction", "speed", "advanceOnClick", "advanceAfterMs"]);
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function packageVersion() {
  const entry = require.resolve("office-kit");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(label + " must be a non-empty string.");
  return value.trim();
}

function own(object, key) {
  return Object.prototype.hasOwnProperty.call(object, key);
}

// Keep the workflow input on the same deliberately small public contract as
// SlideTransition. The explicit expected value is a source precondition, not
// a request to interpret an arbitrary native p:transition graph.
function canonicalTransition(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new TypeError(label + " must be one transition object.");
  }
  const unsupported = Object.keys(value).filter((key) => !TRANSITION_KEYS.has(key));
  if (unsupported.length) throw new TypeError(label + " has unsupported fields: " + unsupported.join(", ") + ".");
  const effect = String(value.effect || "").trim().toLowerCase();
  if (!EFFECTS.has(effect)) throw new TypeError(label + ".effect must be fade or push.");
  const speed = String(value.speed ?? "medium").trim().toLowerCase();
  if (!SPEEDS.has(speed)) throw new TypeError(label + ".speed must be slow, medium, or fast.");
  const transition = { effect, speed };
  if (effect === "push") {
    const direction = String(value.direction ?? "left").trim().toLowerCase();
    if (!DIRECTIONS.has(direction)) throw new TypeError(label + ".direction must be left, up, right, or down for push.");
    transition.direction = direction;
  } else if (own(value, "direction") && value.direction != null) {
    throw new TypeError(label + ".direction is not valid for fade.");
  }
  if (own(value, "advanceOnClick") && typeof value.advanceOnClick !== "boolean") {
    throw new TypeError(label + ".advanceOnClick must be a boolean.");
  }
  transition.advanceOnClick = value.advanceOnClick ?? true;
  if (own(value, "advanceAfterMs") && value.advanceAfterMs != null) {
    const advanceAfterMs = Number(value.advanceAfterMs);
    if (!Number.isSafeInteger(advanceAfterMs) || advanceAfterMs < 0 || advanceAfterMs > MAX_ADVANCE_AFTER_MS) {
      throw new RangeError(label + ".advanceAfterMs must be an integer from 0 through " + MAX_ADVANCE_AFTER_MS + ".");
    }
    transition.advanceAfterMs = advanceAfterMs;
  }
  return transition;
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function withoutIdsAndTransition(value) {
  if (Array.isArray(value)) return value.map(withoutIdsAndTransition);
  if (!value || typeof value !== "object") return value;
  return Object.fromEntries(Object.entries(value)
    .filter(([key]) => key !== "id" && key !== "transition")
    .map(([key, item]) => [key, withoutIdsAndTransition(item)]));
}

function nonTransitionSnapshot(slide) {
  return withoutIdsAndTransition(slide.toProto());
}

async function slideRenderHashes(presentation) {
  return Promise.all(presentation.slides.items.map(async (slide) => {
    const rendered = await slide.export({ format: "svg" });
    const svg = await rendered.text();
    if (!/<svg\b/i.test(svg)) throw new Error("Presentation model render did not produce SVG.");
    return sha256(Buffer.from(svg));
  }));
}

function xmlAttributes(tag) {
  const attributes = Object.create(null);
  for (const match of String(tag).matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])([\s\S]*?)\2/g)) attributes[match[1]] = match[3];
  return attributes;
}

function resolveRelationshipTarget(target) {
  const resolved = new URL(target, "https://officekit.invalid/ppt/presentation.xml");
  if (resolved.origin !== "https://officekit.invalid") throw new Error("Unexpected PPTX relationship target origin.");
  const partPath = resolved.pathname.replace(/^\/+/, "");
  if (!partPath.startsWith("ppt/") || partPath.split("/").includes("..")) {
    throw new Error("Unsafe PPTX slide relationship target: " + JSON.stringify(target));
  }
  return partPath;
}

async function orderedSlidePartPaths(zip) {
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
  const relationshipsXml = await zip.file("ppt/_rels/presentation.xml.rels")?.async("text");
  if (!presentationXml || !relationshipsXml) throw new Error("PPTX is missing presentation.xml or its relationship part.");
  const relationships = new Map();
  for (const match of relationshipsXml.matchAll(/<Relationship\b[^>]*>/gi)) {
    const attributes = xmlAttributes(match[0]);
    if (!attributes.Id || !attributes.Type?.endsWith("/slide")) continue;
    if (attributes.TargetMode?.toLowerCase() === "external" || !attributes.Target) {
      throw new Error("Presentation slide relationship " + JSON.stringify(attributes.Id) + " is not an internal SlidePart.");
    }
    relationships.set(attributes.Id, resolveRelationshipTarget(attributes.Target));
  }
  const paths = [];
  for (const match of presentationXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sldId\b[^>]*>/gi)) {
    const relationshipId = xmlAttributes(match[0])["r:id"];
    const target = relationships.get(relationshipId);
    if (!target) throw new Error("Presentation slide list references an unresolved relationship " + JSON.stringify(relationshipId) + ".");
    if (!zip.file(target)) throw new Error("Presentation slide relationship points at missing part " + target + ".");
    paths.push(target);
  }
  if (!paths.length || new Set(paths).size !== paths.length) {
    throw new Error("Presentation slide list must contain distinct, resolvable SlideParts.");
  }
  return paths;
}

async function assertPackageScope(sourceBytes, outputBytes, targetIndex) {
  const sourceZip = await JSZip.loadAsync(sourceBytes);
  const outputZip = await JSZip.loadAsync(outputBytes);
  const sourcePaths = Object.keys(sourceZip.files).sort();
  const outputPaths = Object.keys(outputZip.files).sort();
  if (!sameJson(sourcePaths, outputPaths)) throw new Error("Transition edit changed PPTX package topology.");
  const sourceSlidePaths = await orderedSlidePartPaths(sourceZip);
  const outputSlidePaths = await orderedSlidePartPaths(outputZip);
  if (!sameJson(sourceSlidePaths, outputSlidePaths)) throw new Error("Transition edit changed presentation slide-part routing.");
  if (!Number.isInteger(targetIndex) || targetIndex < 0 || targetIndex >= sourceSlidePaths.length) {
    throw new Error("Resolved target slide index is outside the source PPTX slide list.");
  }
  const targetPart = sourceSlidePaths[targetIndex];
  const sourceTarget = await sourceZip.file(targetPart)?.async("uint8array");
  const outputTarget = await outputZip.file(targetPart)?.async("uint8array");
  if (!sourceTarget || !outputTarget) throw new Error("Transition target SlidePart is missing from source or output.");
  if (Buffer.from(sourceTarget).equals(Buffer.from(outputTarget))) {
    throw new Error("Transition edit produced no target SlidePart change; no output was published.");
  }
  for (const partPath of sourcePaths) {
    if (sourceZip.files[partPath].dir || partPath === targetPart) continue;
    const before = await sourceZip.file(partPath).async("uint8array");
    const after = await outputZip.file(partPath).async("uint8array");
    if (!Buffer.from(before).equals(Buffer.from(after))) {
      throw new Error("Transition edit changed non-target package part " + partPath + ".");
    }
  }
  return {
    targetPart,
    partCount: sourcePaths.length,
    nonTargetPartsByteIdentical: true,
    sourceTargetSha256: sha256(sourceTarget),
    outputTargetSha256: sha256(outputTarget),
  };
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

async function publishNoReplace(temporaryPath, finalPath, label) {
  try {
    await fs.link(temporaryPath, finalPath);
  } catch (error) {
    if (error?.code === "EEXIST") throw new Error(label + " already exists; refusing to overwrite it.");
    if (!["EPERM", "EXDEV", "ENOTSUP", "EOPNOTSUPP"].includes(error?.code)) throw error;
    try {
      await fs.copyFile(temporaryPath, finalPath, FS_CONSTANTS.COPYFILE_EXCL);
    } catch (copyError) {
      if (copyError?.code === "EEXIST") throw new Error(label + " already exists; refusing to overwrite it.");
      throw copyError;
    }
  }
  // The publication step has already succeeded. A best-effort cleanup failure
  // must not report a failed transaction after the no-replace final path exists.
  await fs.rm(temporaryPath, { force: true }).catch(() => {});
}

export async function editPptxTransition({ inputPath, outputPath, auditPath, slideName, expectedTransition, replacementTransition }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  const expectedSlideName = requiredText(slideName, "slideName");
  const expected = canonicalTransition(expectedTransition, "expectedTransition");
  const replacement = canonicalTransition(replacementTransition, "replacementTransition");
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original presentation remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) {
    throw new Error("auditPath must be distinct from source and PPTX output paths.");
  }
  if (sameJson(expected, replacement)) throw new Error("replacementTransition must differ from expectedTransition.");
  await assertAbsent(finalPath, "outputPath");
  await assertAbsent(finalAuditPath, "auditPath");

  const source = await fs.readFile(sourcePath);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, {
    type: PPTX_MIME,
    name: path.basename(sourcePath),
  }));
  const targets = presentation.slides.items.filter((slide) => slide.name === expectedSlideName);
  if (targets.length !== 1) {
    throw new Error("Expected exactly one imported slide named " + JSON.stringify(expectedSlideName) + "; found " + targets.length + ".");
  }
  const target = targets[0];
  const targetIndex = presentation.slides.items.indexOf(target);
  const capability = target.transition.capability;
  if (!capability.sourceBound || !capability.partPresent || !capability.editable) {
    throw new Error("Selected imported slide transition does not satisfy the editable canonical direct fade/push profile.");
  }
  const sourceTransition = canonicalTransition(target.transition.toJSON(), "imported transition");
  if (!sameJson(sourceTransition, expected)) {
    throw new Error("Selected imported slide transition does not match expectedTransition; no output was written.");
  }
  const sourceNonTransitionSnapshots = presentation.slides.items.map(nonTransitionSnapshot);
  const sourceRenderHashes = await slideRenderHashes(presentation);
  target.setTransition(replacement);

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  let publishedOutput = false;
  let publishedAudit = false;
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const packageScope = await assertPackageScope(source, output, targetIndex);
    const reimported = await PresentationFile.importPptx(new FileBlob(output, {
      type: PPTX_MIME,
      name: path.basename(finalPath),
    }));
    const reimportedTargets = reimported.slides.items.filter((slide) => slide.name === expectedSlideName);
    if (reimportedTargets.length !== 1) throw new Error("PPTX second import did not retain the unique target slide name.");
    const reimportedTarget = reimportedTargets[0];
    if (!sameJson(canonicalTransition(reimportedTarget.transition.toJSON(), "reimported transition"), replacement)) {
      throw new Error("PPTX second import did not retain replacementTransition semantics.");
    }
    if (!reimportedTarget.transition.capability.editable) {
      throw new Error("PPTX second import did not retain an editable canonical transition profile.");
    }
    const outputNonTransitionSnapshots = reimported.slides.items.map(nonTransitionSnapshot);
    if (!sameJson(sourceNonTransitionSnapshots, outputNonTransitionSnapshots)) {
      throw new Error("Transition edit changed a non-transition slide semantic.");
    }
    const outputRenderHashes = await slideRenderHashes(reimported);
    if (!sameJson(sourceRenderHashes, outputRenderHashes)) {
      throw new Error("Transition edit changed a static slide model render; slideshow playback is not being substituted for visual QA.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error("Presentation verification failed: " + verification.ndjson);
    if (!Buffer.from(await fs.readFile(sourcePath)).equals(source)) {
      throw new Error("Source PPTX changed during transition transaction; refusing to publish output.");
    }
    const audit = {
      schema: "office-kit.pptx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "office-kit", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "source-bound-transition-edit",
        slideName: expectedSlideName,
        slideNumber: targetIndex + 1,
        partPath: packageScope.targetPart,
        expectedTransition: expected,
        replacementTransition: replacement,
        capability,
      },
      warnings: ["Static render verification proves visible-slide stability, not native slideshow playback timing or effect behavior."],
      validation: {
        package: { ok: true, ...packageScope, onlyTargetSlidePartChanged: true },
        reimport: { ok: true, editable: true, replacementSemanticsRetained: true },
        nonTransitionSemantics: { ok: true, stable: true },
        modelRender: { ok: true, sourceSha256: sourceRenderHashes, outputSha256: outputRenderHashes, byteIdentical: true },
        verify: { ok: verification.ok },
      },
    };
    await fs.writeFile(temporaryAuditPath, JSON.stringify(audit, null, 2));
    await publishNoReplace(temporaryPath, finalPath, "outputPath");
    publishedOutput = true;
    await publishNoReplace(temporaryAuditPath, finalAuditPath, "auditPath");
    publishedAudit = true;
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([
      fs.rm(temporaryPath, { force: true }),
      fs.rm(temporaryAuditPath, { force: true }),
      ...(publishedOutput ? [fs.rm(finalPath, { force: true })] : []),
      ...(publishedAudit ? [fs.rm(finalAuditPath, { force: true })] : []),
    ]);
    throw error;
  }
}

function parseCli(argv) {
  const [inputPath, outputPath, auditPath, slideName, expectedJson, replacementJson] = argv;
  try {
    return {
      inputPath,
      outputPath,
      auditPath,
      slideName,
      expectedTransition: JSON.parse(requiredText(expectedJson, "expectedTransitionJson")),
      replacementTransition: JSON.parse(requiredText(replacementJson, "replacementTransitionJson")),
    };
  } catch (error) {
    throw new TypeError("expectedTransitionJson and replacementTransitionJson must each be one valid JSON object: " + error.message);
  }
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editPptxTransition(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    targetPart: result.audit.operation.partPath,
  }));
}
