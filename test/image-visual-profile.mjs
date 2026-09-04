import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import sharp from "sharp";

import { inspectImageBytes } from "../src/shared/image-bytes.mjs";
import {
  deriveImageVisualProfile,
  mergeImageVisualProfiles,
  normalizeImageVisualProfile,
} from "../src/shared/image-profile.mjs";
import { createTask } from "../src/cli/task-store.mjs";
import { addTaskImageAsset, imageTaskState } from "../src/images/task-assets.mjs";

const transparentPng = await sharp(Buffer.from([
  0, 0, 0, 0, 255, 255, 255, 255, 255, 255, 255, 255, 0, 0, 0, 0,
  0, 0, 0, 0, 255, 255, 255, 255, 128, 128, 128, 128, 0, 0, 0, 0,
]), { raw: { width: 4, height: 2, channels: 4 } }).png().toBuffer();
const opaqueJpeg = await sharp({
  create: { width: 2, height: 2, channels: 3, background: { r: 20, g: 30, b: 40 } },
}).jpeg().toBuffer();

const pngInspection = inspectImageBytes(transparentPng, {
  declaredMimeType: "image/png",
  label: "profile PNG",
  maxBytes: 20 * 1024 * 1024,
  maxPixels: 40_000_000,
  maxDimension: 16_384,
});
assert.deepEqual(pngInspection.visualProfile, {
  alphaPresent: true,
  subjectBounds: { x: 0.25, y: 0, width: 0.5, height: 1 },
  edgeQuality: "soft",
  shadowMode: "unknown",
});

const jpegInspection = inspectImageBytes(opaqueJpeg, {
  declaredMimeType: "image/jpeg",
  label: "profile JPEG",
  maxBytes: 20 * 1024 * 1024,
  maxPixels: 40_000_000,
  maxDimension: 16_384,
});
assert.deepEqual(jpegInspection.visualProfile, {
  alphaPresent: false,
  subjectBounds: null,
  edgeQuality: "unknown",
  shadowMode: "unknown",
});

assert.deepEqual(mergeImageVisualProfiles(pngInspection.visualProfile, {
  edgeQuality: "fringe",
  shadowMode: "baked",
}), {
  alphaPresent: true,
  subjectBounds: { x: 0.25, y: 0, width: 0.5, height: 1 },
  edgeQuality: "fringe",
  shadowMode: "baked",
});
assert.throws(
  () => mergeImageVisualProfiles(jpegInspection.visualProfile, { alphaPresent: true }),
  /contradicts the image bytes/u,
);
assert.throws(
  () => normalizeImageVisualProfile({ subjectBounds: { x: 0.8, y: 0, width: 0.4, height: 0.2 } }),
  /inside 0..1/u,
);

const workspace = await mkdtemp(path.join(os.tmpdir(), "office-kit-image-profile-"));
try {
  const pngPath = path.join(workspace, "transparent.png");
  await writeFile(pngPath, transparentPng);
  assert.ok((await readFile(pngPath)).length > 0);
  const task = await createTask({ workspaceRoot: workspace, goal: "image profile contract" });
  const asset = await addTaskImageAsset(task, {
    bytes: transparentPng,
    mimeType: "image/png",
    rights: "user-provided",
    source: { kind: "file", originalPath: pngPath },
    visualProfile: { shadowMode: "separate" },
    now: new Date("2026-09-03T00:00:00.000Z"),
  });
  assert.equal(asset.visualProfile.shadowMode, "separate");
  assert.deepEqual((await imageTaskState({ workspaceRoot: workspace, taskId: task.manifest.id })).assets[0].visualProfile, {
    alphaPresent: true,
    subjectBounds: { x: 0.25, y: 0, width: 0.5, height: 1 },
    edgeQuality: "soft",
    shadowMode: "separate",
  });
} finally {
  await rm(workspace, { recursive: true, force: true });
}

assert.deepEqual(deriveImageVisualProfile(transparentPng, { mimeType: "image/png" }), pngInspection.visualProfile);
console.log("image visual profile smoke ok");
