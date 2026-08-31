import { imageError } from "./errors.mjs";

export const IMAGE_RIGHTS = Object.freeze([
  "user-provided",
  "generated",
  "permission",
  "public-domain",
  "cc0",
  "cc-by",
  "official-press-kit",
  "lucide-isc",
]);

const ALLOWED_RIGHTS = new Set(IMAGE_RIGHTS);
const MAX_TEXT = 2_000;

function text(value, label, { required = false, maximum = MAX_TEXT } = {}) {
  const normalized = value == null ? "" : String(value).trim();
  if (required && normalized === "") throw imageError("missing-rights-metadata", `${label} is required.`);
  if (normalized.length > maximum || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/u.test(normalized)) {
    throw imageError("invalid-rights-metadata", `${label} must contain at most ${maximum} safe characters.`);
  }
  return normalized || undefined;
}

function httpsUrl(value, label, { required = false } = {}) {
  const normalized = text(value, label, { required, maximum: 4_096 });
  if (!normalized) return undefined;
  let parsed;
  try { parsed = new URL(normalized); }
  catch { throw imageError("invalid-rights-metadata", `${label} must be a valid HTTPS URL.`); }
  if (parsed.protocol !== "https:" || parsed.username || parsed.password) {
    throw imageError("invalid-rights-metadata", `${label} must be an HTTPS URL without credentials.`);
  }
  return parsed.href;
}

function defaultCreditLine({ rights, title, author, sourcePage, licenseUrl }) {
  if (rights === "cc-by") {
    const work = title ? `“${title}”` : "Image";
    return `${work} by ${author}, CC BY (${licenseUrl})${sourcePage ? `, source: ${sourcePage}` : ""}`;
  }
  if (rights === "lucide-isc") return "Lucide icon — ISC License (https://lucide.dev/license)";
  return undefined;
}

export function normalizeImageRights(value, metadata = {}) {
  const rights = String(value || "").trim().toLowerCase();
  if (!ALLOWED_RIGHTS.has(rights)) {
    throw imageError("image-rights-blocked", `Image rights ${rights || "(missing)"} are not allowed.`);
  }
  const author = text(metadata.author, "Image author", { required: rights === "cc-by" });
  const licenseUrl = httpsUrl(metadata.licenseUrl, "Image license URL", { required: rights === "cc-by" });
  const sourcePage = httpsUrl(metadata.sourcePage, "Image source page");
  const title = text(metadata.title, "Image title");
  const provider = text(metadata.provider, "Image provider", { maximum: 80 });
  const evidence = String(metadata.evidence || "user-declared").trim().toLowerCase();
  if (!new Set(["user-declared", "provider-declared", "machine-readable", "package-license"]).has(evidence)) {
    throw imageError("invalid-rights-metadata", "Image rights evidence is invalid.");
  }
  const visibleAttributionRequired = rights === "cc-by";
  const creditLine = text(
    metadata.creditLine ?? defaultCreditLine({ rights, title, author, sourcePage, licenseUrl }),
    "Image credit line",
    { required: visibleAttributionRequired, maximum: 4_096 },
  );
  return Object.freeze({
    rights,
    evidence,
    legalVerification: false,
    visibleAttributionRequired,
    ...(provider ? { provider } : {}),
    ...(title ? { title } : {}),
    ...(author ? { author } : {}),
    ...(sourcePage ? { sourcePage } : {}),
    ...(licenseUrl ? { licenseUrl } : {}),
    ...(creditLine ? { creditLine } : {}),
  });
}

export function providerLicenseToRights(value) {
  const normalized = String(value || "").trim().toUpperCase();
  if (normalized === "PUBLIC_DOMAIN") return "public-domain";
  if (normalized === "CC0") return "cc0";
  if (normalized === "CC_BY") return "cc-by";
  throw imageError("image-rights-blocked", `Provider license ${normalized || "UNKNOWN"} is not allowed.`);
}

export function imageRightsCompatible(left, right) {
  for (const field of ["rights", "provider", "author", "sourcePage", "licenseUrl", "creditLine"]) {
    if ((left?.[field] || undefined) !== (right?.[field] || undefined)) return false;
  }
  return true;
}
