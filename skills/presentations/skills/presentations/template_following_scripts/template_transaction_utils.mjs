import crypto from "node:crypto";
import { constants as FS_CONSTANTS } from "node:fs";
import fs from "node:fs/promises";
import path from "node:path";

export const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

export function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

export function isWithin(child, parent) {
  const relative = path.relative(parent, child);
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

export function pathsOverlap(left, right) {
  return isWithin(left, right) || isWithin(right, left);
}

export function slidesFromPresentation(presentation) {
  if (Array.isArray(presentation.slides?.items)) return presentation.slides.items;
  if (Number.isInteger(presentation.slides?.count) && typeof presentation.slides.getItem === "function") {
    return Array.from({ length: presentation.slides.count }, (_, index) => presentation.slides.getItem(index));
  }
  throw new Error("Could not enumerate imported presentation slides.");
}

export async function assertRegularFile(filePath, label, maxBytes) {
  const stat = await fs.stat(filePath).catch(() => undefined);
  if (!stat?.isFile()) throw new Error(`Missing ${label}: ${filePath}`);
  if (maxBytes && stat.size > maxBytes) throw new Error(`${label} exceeds the ${maxBytes}-byte budget: ${filePath}`);
  return stat;
}

export async function assertAbsent(filePath, label) {
  const existing = await fs.lstat(filePath).catch((error) => {
    if (error?.code === "ENOENT") return undefined;
    throw error;
  });
  if (existing) throw new Error(`${label} already exists; refusing to overwrite it: ${filePath}`);
}

export async function publishFileNoReplace(source, destination, label) {
  await fs.mkdir(path.dirname(destination), { recursive: true });
  try {
    await fs.link(source, destination);
  } catch (error) {
    if (error?.code === "EEXIST") throw new Error(`${label} already exists; refusing to overwrite it: ${destination}`);
    if (!["EPERM", "EXDEV", "ENOTSUP", "EOPNOTSUPP"].includes(error?.code)) throw error;
    try {
      await fs.copyFile(source, destination, FS_CONSTANTS.COPYFILE_EXCL);
    } catch (copyError) {
      if (copyError?.code === "EEXIST") throw new Error(`${label} already exists; refusing to overwrite it: ${destination}`);
      await fs.rm(destination, { force: true }).catch(() => undefined);
      throw copyError;
    }
  }
}

export async function publishDirectoryNoReplace(source, destination, label) {
  await fs.mkdir(path.dirname(destination), { recursive: true });
  try {
    await fs.mkdir(destination);
  } catch (error) {
    if (error?.code === "EEXIST") {
      throw new Error(`${label} already exists; refusing to overwrite it: ${destination}`);
    }
    throw error;
  }
  try {
    const entries = await fs.readdir(source);
    for (const entry of entries) {
      await fs.cp(path.join(source, entry), path.join(destination, entry), {
        recursive: true,
        force: false,
        errorOnExist: true,
      });
    }
  } catch (error) {
    await fs.rm(destination, { recursive: true, force: true });
    throw error;
  }
}

export async function writeJson(filePath, value) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

export function relativeFromWorkspace(workspaceDir, filePath) {
  return path.relative(workspaceDir, filePath).split(path.sep).join("/");
}

export async function runContactSheet(previewPaths, outputPath) {
  if (!outputPath) return undefined;
  let sharp;
  try {
    const module = await import("sharp");
    sharp = module.default || module;
  } catch (error) {
    throw new Error(`Contact sheet generation requires OfficeKit's packaged sharp runtime: ${error.message}`);
  }
  const metadata = await Promise.all(previewPaths.map((previewPath) => sharp(previewPath).metadata()));
  const tileWidth = Math.max(...metadata.map((item) => Number(item.width || 0)));
  const tileHeight = Math.max(...metadata.map((item) => Number(item.height || 0)));
  if (tileWidth <= 0 || tileHeight <= 0) throw new Error("Contact sheet previews must have positive raster dimensions.");
  const columns = Math.min(3, previewPaths.length);
  const rows = Math.ceil(previewPaths.length / columns);
  const labelHeight = 48;
  const padding = 18;
  const composites = previewPaths.flatMap((previewPath, index) => {
    const row = Math.floor(index / columns);
    const column = index % columns;
    const left = padding + column * (tileWidth + padding);
    const top = padding + row * (tileHeight + labelHeight + padding);
    const label = Buffer.from(
      `<svg width="${tileWidth}" height="${labelHeight}" xmlns="http://www.w3.org/2000/svg"><text x="8" y="31" font-family="sans-serif" font-size="18" fill="#141e32">Slide ${String(index + 1).padStart(2, "0")}</text></svg>`,
    );
    return [
      { input: previewPath, left, top },
      { input: label, left, top: top + tileHeight },
    ];
  });
  await sharp({
    create: {
      width: columns * tileWidth + (columns + 1) * padding,
      height: rows * (tileHeight + labelHeight) + (rows + 1) * padding,
      channels: 3,
      background: "white",
    },
  }).composite(composites).png().toFile(outputPath);
  return outputPath;
}

export function fileBlob(FileBlob, bytes, name) {
  return new FileBlob(bytes, { type: PPTX_MIME, name });
}

export async function exportBytes(PresentationFile, presentation) {
  const exported = await PresentationFile.exportPptx(presentation);
  return Buffer.from(exported.bytes);
}

export async function modelVisualSha256(slide) {
  const svg = await slide.export({ format: "svg" });
  const text = await svg.text();
  if (!/<svg\b/i.test(text)) throw new Error("Presentation model render did not produce SVG.");
  return sha256(Buffer.from(text.replace(/\sdata-[\w-]*id="[^"]*"/gi, "")));
}
