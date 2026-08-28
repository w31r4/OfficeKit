#!/usr/bin/env node

import { createHash } from "node:crypto";
import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, PresentationFile } from "../src/index.mjs";
import { directPresentationChildren } from "../src/presentation/group-shapes.mjs";

const SCHEMA = "office-kit/pptx-six-sample-import-evidence/v1";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const MAX_SOURCE_BYTES = 128 * 1024 * 1024;
const DEFAULT_ASSETS_DIR = path.resolve("tmp/reference-pptx-downloads");
const DEFAULT_OUTPUT = path.resolve("tmp/presentation-six-sample-import/evidence.v1.json");

// These are reference inputs only.  The files stay under ignored tmp/ because
// the SlidesCarnival terms do not allow unmodified redistribution.
export const SOURCES = Object.freeze([
  {
    id: "nasa-froste-thermal-design",
    fileName: "nasa-froste-thermal-design.pptx",
    sha256: "41568341c527866c9c8917229d190bca8dc99f0a137e97343b2c30d03f1db7b0",
    source: "NASA NTRS 20250007120",
  },
  {
    id: "nasa-mms-machine-learning",
    fileName: "nasa-mms-machine-learning.pptx",
    sha256: "531c82797fde09b1ebe1e868ca9cd44c3e2f675dc8f09f58b54bab6a62629723",
    source: "NASA NTRS 20250000748",
  },
  {
    id: "slidescarnival-business-infographic",
    fileName: "slidescarnival-business-infographic.pptx",
    sha256: "8db900eb9fbc5375d6b69eccffebd5ebb002f2f6641a89f19364a74e1d7e1e26",
    source: "SlidesCarnival Business Infographic",
  },
  {
    id: "slidescarnival-data-particles",
    fileName: "slidescarnival-data-particles.pptx",
    sha256: "07cd6c7e3c12335716fbfddb1ccde353c9d21959427e2639dea29eca1573464f",
    source: "SlidesCarnival Data Particles",
  },
  {
    id: "slidescarnival-minimal-business",
    fileName: "slidescarnival-minimal-business.pptx",
    sha256: "5076299de79a0c8ede47bb2f3c286a8e021ef0f7a55c991279ad03d4717ca334",
    source: "SlidesCarnival Minimal Business",
  },
  {
    id: "slidescarnival-professional-minimalist",
    fileName: "slidescarnival-professional-minimalist.pptx",
    sha256: "ed26f074e25361e4baf76d1cc2508596655b9d4f3fc2c659ecd962f76e0547ab",
    source: "SlidesCarnival Professional Minimalist",
  },
]);

export async function collectSixSampleEvidence({ assetsDir = DEFAULT_ASSETS_DIR } = {}) {
  const root = path.resolve(assetsDir);
  const results = [];
  for (const source of SOURCES) {
    const sourcePath = path.join(root, source.fileName);
    const bytes = await readPptx(sourcePath);
    const digest = sha256(bytes);
    if (digest !== source.sha256) throw new Error(`${source.id} source SHA-256 mismatch: ${digest}`);
    const packageInfo = await packageEvidence(bytes);
    const presentation = await importPresentation(bytes);
    const records = parseNdjson(presentation.inspect({ kind: "importObject", maxChars: Infinity }).ndjson);
    const rawObjectCount = packageInfo.rawObjectCount;
    if (presentation.slides.count !== packageInfo.slideCount) {
      throw new Error(`${source.id} imported ${presentation.slides.count} slides, expected ${packageInfo.slideCount}.`);
    }
    if (records.length !== rawObjectCount) {
      throw new Error(`${source.id} classified ${records.length} of ${rawObjectCount} visible top-level objects.`);
    }
    if (new Set(records.map((record) => record.targetId)).size !== records.length) {
      throw new Error(`${source.id} has duplicate imported object IDs.`);
    }
    const noOp = await PresentationFile.exportPptx(presentation);
    if (!Buffer.from(noOp.bytes).equals(bytes)) throw new Error(`${source.id} no-op export is not byte-identical.`);

    const profile = presentation.designProfile({ maxItems: 64, includeComponentCandidates: true });
    const placement = await verifyPlacementEdit(bytes);
    const text = await verifyTextEdit(bytes);
    const nativeText = await verifyNativeTextEdit(bytes);
    const reuse = await verifyOneSlideReuse(bytes);
    results.push({
      id: source.id,
      fileName: source.fileName,
      source: source.source,
      sourceSha256: digest,
      bytes: bytes.byteLength,
      slides: packageInfo.slideCount,
      visibleTopLevelObjects: rawObjectCount,
      classifiedTopLevelObjects: records.length,
      rawRootKinds: packageInfo.rawRootKinds,
      objectKinds: counts(records.map((record) => record.objectKind)),
      classifications: counts(records.map((record) => record.classification)),
      nativeLeafKinds: counts(records.flatMap((record) => record.nativeLeafKinds || [])),
      noOpByteIdentical: true,
      placement,
      text,
      nativeText,
      sourceSlideReuse: reuse,
      nativeLeafCount: records.reduce((sum, record) => sum + Number(record.nativeLeafCount || 0), 0),
      nativeTextLeafCount: records
        .filter((record) => (record.nativeLeafKinds || []).includes("nativeText"))
        .reduce((sum, record) => sum + Number(record.nativeLeafCount || 0), 0),
      designProfile: profileSummary(profile),
    });
  }
  return {
    schema: SCHEMA,
    sourcePolicy: "ignored reference inputs; do not redistribute unmodified SlidesCarnival files",
    totals: {
      sources: results.length,
      slides: results.reduce((sum, result) => sum + result.slides, 0),
      visibleTopLevelObjects: results.reduce((sum, result) => sum + result.visibleTopLevelObjects, 0),
      noOpByteIdentical: results.every((result) => result.noOpByteIdentical),
      placementEdits: results.filter((result) => result.placement.status === "passed").length,
      textEdits: results.filter((result) => result.text.status === "passed").length,
      nativeTextEdits: results.filter((result) => result.nativeText.status === "passed").length,
      sourceSlideReuse: results.filter((result) => result.sourceSlideReuse.status === "passed").length,
    },
    sources: results,
  };
}

async function verifyTextEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const target = firstTextRun(presentation);
  if (!target) return { status: "blocked", reason: "no safe text run was discovered" };
  const needle = target.run.text.trim().split(/\s+/u)[0];
  target.shape.text.replace(needle, `${needle} OfficeKit`);
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const value = reopened.resolve(target.shape.id)?.text?.value || "";
  if (!value.includes(`${needle} OfficeKit`)) throw new Error(`Text edit did not survive re-import for ${target.shape.id}.`);
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Text edit changed unexpected parts for ${target.shape.id}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.shape.id, changedParts };
}

async function verifyNativeTextEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "nativeText");
  if (!target) return { status: "blocked", reason: "no bounded opaque-table text leaf was discovered" };
  const value = `${target.value} OfficeKit`;
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const object = reopened.resolve(target.targetId);
  const leaf = object?.nativeTextLeaves?.find((candidate) => candidate.textLeafIndex === target.textLeafIndex);
  if (leaf?.text !== value) throw new Error(`Native text edit did not survive re-import for ${target.targetId}.`);
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native text edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, textLeafIndex: target.textLeafIndex, changedParts };
}

async function verifyPlacementEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const target = firstPlacementObject(presentation);
  if (!target) return { status: "blocked", reason: "no bounded placement capability was discovered" };
  const before = { ...target.object.position };
  target.object.setPosition({ left: before.left + 3, top: before.top + 3 });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = reopened.resolve(target.object.id);
  if (!rebound || Math.abs(rebound.position.left - before.left - 3) > 0.01 || Math.abs(rebound.position.top - before.top - 3) > 0.01) {
    throw new Error(`Placement edit did not survive re-import for ${target.object.id}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Placement edit changed unexpected parts for ${target.object.id}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.object.id, nativeKind: target.object.nativeKind, changedParts };
}

async function verifyOneSlideReuse(bytes) {
  const presentation = await importPresentation(bytes);
  const sourceSlide = presentation.slides.items[0];
  if (!sourceSlide.cloneCapability.supported) return { status: "blocked", reason: sourceSlide.cloneCapability.blockedReason || "source slide is not reusable" };
  const originalSlideCount = presentation.slides.count;
  const allSlidesCloneCapable = presentation.slides.items.every((slide) => slide.cloneCapability.supported);
  sourceSlide.duplicate();
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  if (reopened.slides.count !== originalSlideCount + 1) throw new Error("Source slide reuse produced an unexpected slide count.");
  const sourceZip = await JSZip.loadAsync(bytes);
  const outputZip = await JSZip.loadAsync(output.bytes);
  const sourceSlideXml = await sourceZip.file("ppt/slides/slide1.xml").async("uint8array");
  const outputSlideXml = await outputZip.file("ppt/slides/slide1.xml").async("uint8array");
  if (!Buffer.from(sourceSlideXml).equals(Buffer.from(outputSlideXml))) throw new Error("Source slide changed during reuse.");
  return {
    status: "passed",
    sourceSlideId: sourceSlide.id,
    sourceSlideUnchanged: true,
    outputSlideCount: reopened.slides.count,
    allSlidesCloneCapable,
  };
}

function firstTextRun(presentation) {
  for (const slide of presentation.slides.items) {
    for (const shape of slide.shapes?.items || []) {
      for (const paragraph of shape.text?.paragraphs || []) {
        const run = paragraph.runs?.find((candidate) => typeof candidate.text === "string" && candidate.text.trim().length >= 4);
        if (run) return { slide: slide.index + 1, shape, run };
      }
    }
  }
  return undefined;
}

function firstPlacementObject(presentation) {
  for (const slide of presentation.slides.items) {
    const object = (slide.nativeObjects?.items || []).find((candidate) => candidate.placementCapability?.supported === true);
    if (object) return { slide: slide.index + 1, object };
  }
  return undefined;
}

async function packageEvidence(bytes) {
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  const slides = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide[1-9][0-9]*[.]xml$/u.test(name))
    .sort((left, right) => slideOrdinal(left) - slideOrdinal(right));
  let rawObjectCount = 0;
  const rawRootKinds = {};
  for (const slide of slides) {
    const roots = directPresentationChildren(await zip.file(slide).async("text"), "spTree");
    for (const root of roots) {
      if (["nvGrpSpPr", "grpSpPr", "extLst"].includes(root.localName)) continue;
      rawObjectCount += 1;
      rawRootKinds[root.localName] = (rawRootKinds[root.localName] || 0) + 1;
    }
  }
  return { slideCount: slides.length, rawObjectCount, rawRootKinds: sortObject(rawRootKinds) };
}

async function changedPackageParts(sourceBytes, outputBytes) {
  const source = await JSZip.loadAsync(sourceBytes, { checkCRC32: true });
  const output = await JSZip.loadAsync(outputBytes, { checkCRC32: true });
  const names = [...new Set([
    ...Object.keys(source.files).filter((name) => !source.files[name].dir),
    ...Object.keys(output.files).filter((name) => !output.files[name].dir),
  ])].sort();
  const changed = [];
  for (const name of names) {
    const before = source.file(name) ? await source.file(name).async("uint8array") : undefined;
    const after = output.file(name) ? await output.file(name).async("uint8array") : undefined;
    if (!before || !after || !Buffer.from(before).equals(Buffer.from(after))) changed.push(name);
  }
  return changed;
}

function profileSummary(profile) {
  return {
    schema: profile.schema,
    sourceBound: profile.source?.sourceBound === true,
    revisionSha256: profile.source?.revisionSha256,
    canvas: profile.canvas,
    layoutFamilies: profile.layoutFamilies?.length || 0,
    slideArchetypes: profile.slideArchetypes?.length || 0,
    reusableComponents: profile.reusableComponents?.length || 0,
    componentCandidates: Number(profile.componentCandidates?.total || 0),
    svgAssets: profile.designLanguage?.vectorAssets?.assetCount || 0,
    nativeOpaque: profile.nativeOpaque?.count || profile.nativeOpaque?.length || 0,
  };
}

async function importPresentation(bytes) {
  return PresentationFile.importPptx(new FileBlob(bytes, { type: PPTX_MIME }));
}

async function readPptx(filePath) {
  const info = await stat(filePath);
  if (!info.isFile() || info.size < 1 || info.size > MAX_SOURCE_BYTES) throw new RangeError(`PPTX input is outside 1..${MAX_SOURCE_BYTES}: ${filePath}`);
  const bytes = await readFile(filePath);
  if (bytes.byteLength !== info.size) throw new Error(`PPTX input changed while reading: ${filePath}`);
  return bytes;
}

function parseNdjson(value) {
  return String(value || "").split("\n").filter(Boolean).map((line) => JSON.parse(line));
}

function counts(values) {
  return sortObject(values.reduce((result, value) => {
    result[value] = (result[value] || 0) + 1;
    return result;
  }, {}));
}

function sortObject(value) {
  return Object.fromEntries(Object.entries(value).sort(([left], [right]) => left.localeCompare(right)));
}

function slideOrdinal(name) {
  return Number(/slide([1-9][0-9]*)[.]xml$/u.exec(name)?.[1]);
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function parseArgs(argv) {
  let assetsDir = DEFAULT_ASSETS_DIR;
  let output = DEFAULT_OUTPUT;
  for (let index = 0; index < argv.length; index += 1) {
    if (argv[index] === "--assets-dir") assetsDir = argv[++index];
    else if (argv[index] === "--output") output = argv[++index];
    else throw new Error(`Unknown option ${argv[index]}.`);
  }
  return { assetsDir, output };
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const evidence = await collectSixSampleEvidence(options);
  const output = path.resolve(options.output);
  await mkdir(path.dirname(output), { recursive: true });
  await writeFile(output, `${JSON.stringify(evidence, null, 2)}\n`);
  process.stdout.write(`${JSON.stringify({ ok: true, output, ...evidence.totals })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
