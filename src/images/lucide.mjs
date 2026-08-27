import { inspectImageBytes } from "../shared/image-bytes.mjs";
import { imageError } from "./errors.mjs";
import { normalizeImageRights } from "./rights.mjs";

const ICON_NAME = /^[a-z0-9][a-z0-9-]{0,127}$/u;
const LICENSE_URL = "https://lucide.dev/license";

function tokens(value) {
  return String(value || "").toLowerCase().split(/[^a-z0-9]+/u).filter(Boolean);
}

function scoreIcon(name, query) {
  const normalizedQuery = tokens(query).join("-");
  const nameTokens = new Set(tokens(name));
  const queryTokens = tokens(query);
  if (name === normalizedQuery) return 10_000;
  let score = name.startsWith(`${normalizedQuery}-`) ? 2_000 : name.includes(normalizedQuery) ? 1_000 : 0;
  for (const token of queryTokens) {
    if (nameTokens.has(token)) score += 300;
    else if (name.includes(token)) score += 100;
  }
  score -= Math.abs(nameTokens.size - queryTokens.length) * 2;
  return score;
}

function iconSvg(name, icon, collection) {
  const width = Number(icon.width ?? collection.width ?? 24);
  const height = Number(icon.height ?? collection.height ?? 24);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    throw imageError("invalid-lucide-icon", `Lucide icon ${name} has invalid dimensions.`);
  }
  const body = String(icon.body || "").replaceAll("currentColor", "#000000");
  const source = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" aria-hidden="true">${body}</svg>`;
  const bytes = Buffer.from(source);
  inspectImageBytes(bytes, { declaredMimeType: "image/svg+xml", label: `Lucide icon ${name}`, maxBytes: 1_048_576, maxPixels: 1_000_000, maxDimension: 4_096 });
  return bytes;
}

async function collection() {
  const { icons } = await import("@iconify-json/lucide");
  return icons;
}

export async function searchLucideIcons(query, { maximum = 5 } = {}) {
  const icons = await collection();
  const ranked = Object.keys(icons.icons || {})
    .map((name) => ({ name, score: scoreIcon(name, query) }))
    .filter((candidate) => candidate.score > 0)
    .sort((left, right) => right.score - left.score || left.name.localeCompare(right.name))
    .slice(0, maximum);
  return ranked.map(({ name, score }) => {
    const bytes = iconSvg(name, icons.icons[name], icons);
    return {
      provider: "lucide",
      kind: "icon",
      title: name,
      iconName: name,
      acquisitionUrl: `lucide:${name}`,
      previewUrl: `data:image/svg+xml;base64,${bytes.toString("base64")}`,
      sourcePage: `https://lucide.dev/icons/${name}`,
      width: Number(icons.icons[name].width ?? icons.width ?? 24),
      height: Number(icons.icons[name].height ?? icons.height ?? 24),
      mimeType: "image/svg+xml",
      score,
      rights: normalizeImageRights("lucide-isc", {
        provider: "lucide",
        title: name,
        author: "Lucide Contributors",
        sourcePage: `https://lucide.dev/icons/${name}`,
        licenseUrl: LICENSE_URL,
        evidence: "package-license",
      }),
    };
  });
}

export async function materializeLucideIcon(reference) {
  const value = String(reference || "");
  const name = value.startsWith("lucide:") ? value.slice("lucide:".length) : value;
  if (!ICON_NAME.test(name)) throw imageError("invalid-lucide-icon", "Lucide icon reference is invalid.");
  const icons = await collection();
  const icon = icons.icons?.[name];
  if (!icon) throw imageError("lucide-icon-not-found", `Lucide icon ${name} is not present in the pinned collection.`);
  const bytes = iconSvg(name, icon, icons);
  return {
    bytes,
    mimeType: "image/svg+xml",
    source: { kind: "lucide", provider: "lucide", iconName: name, sourcePage: `https://lucide.dev/icons/${name}` },
    rights: "lucide-isc",
    rightsMetadata: {
      provider: "lucide",
      title: name,
      author: "Lucide Contributors",
      sourcePage: `https://lucide.dev/icons/${name}`,
      licenseUrl: LICENSE_URL,
      evidence: "package-license",
    },
  };
}
