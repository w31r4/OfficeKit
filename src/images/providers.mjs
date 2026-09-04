import { createHash } from "node:crypto";

import { normalizeImageMimeType } from "../shared/image-bytes.mjs";
import { imageError, boundedImageError } from "./errors.mjs";
import { searchLucideIcons } from "./lucide.mjs";
import { normalizeImageRights, providerLicenseToRights } from "./rights.mjs";
import { normalizeImageVisualProfile } from "../shared/image-profile.mjs";

const KINDS = new Set(["photo", "illustration", "icon"]);
const PURPOSES = new Set(["hero", "evidence", "context", "decoration"]);
const ORIENTATIONS = new Set(["landscape", "portrait", "square"]);
const DEFAULT_MAXIMUM = 5;
const MAX_MAXIMUM = 20;
const MAX_QUERY_LENGTH = 240;

function text(value, label) {
  const normalized = String(value || "").trim();
  if (!normalized || normalized.length > MAX_QUERY_LENGTH || /[\u0000-\u001f\u007f]/u.test(normalized)) {
    throw imageError("invalid-image-search", `${label} must contain 1 through ${MAX_QUERY_LENGTH} safe characters.`);
  }
  return normalized;
}

function enumValue(value, values, label) {
  const normalized = String(value || "").trim().toLowerCase();
  if (!values.has(normalized)) throw imageError("invalid-image-search", `${label} is invalid.`);
  return normalized;
}

function maximum(value) {
  const normalized = value == null ? DEFAULT_MAXIMUM : Number(value);
  if (!Number.isInteger(normalized) || normalized < 1 || normalized > MAX_MAXIMUM) {
    throw imageError("invalid-image-search", `Image search max must be an integer from 1 through ${MAX_MAXIMUM}.`);
  }
  return normalized;
}

function safeHttps(value) {
  if (!value) return undefined;
  try {
    const parsed = new URL(String(value));
    if (parsed.protocol !== "https:" || parsed.username || parsed.password) return undefined;
    return parsed.href;
  } catch {
    return undefined;
  }
}

function orientationFor(width, height) {
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return undefined;
  const ratio = width / height;
  if (ratio > 1.1) return "landscape";
  if (ratio < 0.9) return "portrait";
  return "square";
}

function queryScore(candidate, query) {
  const queryTokens = new Set(String(query).toLowerCase().split(/[^a-z0-9]+/u).filter(Boolean));
  const haystack = `${candidate.title || ""} ${candidate.author || ""}`.toLowerCase();
  let score = candidate.license === "PUBLIC_DOMAIN" || candidate.license === "CC0" ? 500 : candidate.license === "CC_BY" ? 350 : 0;
  for (const token of queryTokens) if (haystack.includes(token)) score += 40;
  if (Number(candidate.width) >= 1_600 && Number(candidate.height) >= 900) score += 60;
  score += Math.round(Number(candidate.confidence || 0) * 100);
  return score;
}

function blockedLicenseHint(candidate) {
  const evidence = `${candidate.license || ""} ${candidate.licenseUrl || ""}`.toLowerCase();
  if (/by-sa|sharealike/u.test(evidence)) return "share-alike";
  if (/by-nc|noncommercial/u.test(evidence)) return "non-commercial";
  if (/by-nd|noderivatives/u.test(evidence)) return "no-derivatives";
  return undefined;
}

function rejection(candidate, provider, reason) {
  return {
    provider,
    title: String(candidate?.title || "").slice(0, 300) || undefined,
    license: String(candidate?.license || "UNKNOWN").slice(0, 80),
    reason,
  };
}

function normalizeProviderCandidate(candidate, { provider, kind, orientation, query }) {
  const blockedHint = blockedLicenseHint(candidate);
  if (blockedHint) throw imageError("image-rights-blocked", `Candidate declares blocked ${blockedHint} terms.`);
  const acquisitionUrl = safeHttps(candidate.url);
  const sourcePage = safeHttps(candidate.sourcePageUrl);
  if (!acquisitionUrl || !sourcePage) throw imageError("invalid-provider-candidate", "Provider candidate requires HTTPS image and source-page URLs.");
  const rightsName = providerLicenseToRights(candidate.license);
  const rights = normalizeImageRights(rightsName, {
    provider,
    title: candidate.title,
    author: candidate.author,
    sourcePage,
    licenseUrl: candidate.licenseUrl,
    creditLine: candidate.attributionLine,
    evidence: provider === "wikimedia" ? "machine-readable" : "provider-declared",
  });
  const width = Number(candidate.width);
  const height = Number(candidate.height);
  const mimeType = normalizeImageMimeType(candidate.mime);
  if (mimeType && !new Set(["image/png", "image/jpeg", "image/gif"]).has(mimeType)) {
    throw imageError("unsupported-remote-image", `Provider candidate content type ${mimeType} is not PNG, JPEG, or GIF.`);
  }
  const actualOrientation = orientationFor(width, height);
  if (orientation && actualOrientation !== orientation) {
    throw imageError("image-orientation-mismatch", actualOrientation ? `Candidate is ${actualOrientation}, not ${orientation}.` : "Candidate orientation is unknown.");
  }
  return {
    provider,
    kind,
    title: String(candidate.title || "").slice(0, 300) || undefined,
    author: String(candidate.author || "").slice(0, 300) || undefined,
    acquisitionUrl,
    previewUrl: safeHttps(candidate.thumbnailUrl),
    sourcePage,
    width: Number.isFinite(width) && width > 0 ? width : undefined,
    height: Number.isFinite(height) && height > 0 ? height : undefined,
    mimeType: mimeType || undefined,
    rights,
    visualProfile: normalizeImageVisualProfile(candidate.visualProfile || candidate),
    score: queryScore(candidate, query),
    providerConfidence: Number(candidate.confidence || 0),
  };
}

function stableCandidateKey(candidate) {
  return createHash("sha256").update(`${candidate.provider}\0${candidate.acquisitionUrl}`).digest("hex");
}

function providerLicense(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (/\bcc0\b|public[ _-]?domain|\bpdm\b/u.test(normalized)) return normalized.includes("cc0") ? "CC0" : "PUBLIC_DOMAIN";
  if (/\bcc[ _-]?by\b|^by(?:\s|$)/u.test(normalized) && !/\b(?:sa|nc|nd)\b|sharealike|noncommercial|noderivatives/u.test(normalized)) return "CC_BY";
  if (/\bsa\b|sharealike/u.test(normalized)) return "CC_BY_SA";
  return "UNKNOWN";
}

function stripProviderHtml(value) {
  const normalized = String(value || "").replace(/<[^>]+>/gu, " ").replace(/\s+/gu, " ").trim();
  return normalized || undefined;
}

const openverseProvider = Object.freeze({
  async search(query, options = {}) {
    const fetcher = options.fetcher || fetch;
    const url = new URL("https://api.openverse.org/v1/images/");
    url.search = new URLSearchParams({
      q: query,
      page_size: String(options.maxPerProvider || 10),
      license_type: "commercial,modification",
      mature: options.safeSearch === "off" ? "true" : "false",
    }).toString();
    const response = await fetcher(url, { signal: options.signal });
    if (!response.ok) throw new Error(`openverse http ${response.status}`);
    const payload = await response.json();
    return (payload.results || []).map((result) => {
      const license = providerLicense([result.license, result.license_version].filter(Boolean).join(" "));
      return {
        url: result.url,
        thumbnailUrl: result.thumbnail,
        width: result.width,
        height: result.height,
        sourcePageUrl: result.foreign_landing_url || result.detail_url,
        title: result.title,
        author: result.creator,
        license,
        licenseUrl: result.license_url,
        confidence: license === "UNKNOWN" ? 0.2 : 0.9,
      };
    });
  },
});

const wikimediaProvider = Object.freeze({
  async search(query, options = {}) {
    const fetcher = options.fetcher || fetch;
    const url = new URL("https://commons.wikimedia.org/w/api.php");
    url.search = new URLSearchParams({
      action: "query",
      format: "json",
      generator: "search",
      gsrsearch: `${query} filetype:bitmap|drawing`,
      gsrlimit: String(options.maxPerProvider || 10),
      gsrnamespace: "6",
      prop: "imageinfo",
      iiprop: "url|size|mime|extmetadata",
      iiurlwidth: "1600",
      origin: "*",
    }).toString();
    const response = await fetcher(url, {
      headers: { "User-Agent": options.auth?.userAgent || "OfficeKit/2.0 (presentation image sourcing; https://github.com/w31r4/OfficeKit)" },
      signal: options.signal,
    });
    if (!response.ok) throw new Error(`wikimedia http ${response.status}`);
    const payload = await response.json();
    return Object.values(payload?.query?.pages || {}).flatMap((page) => {
      const info = page?.imageinfo?.[0];
      if (!info?.url) return [];
      const metadata = info.extmetadata || {};
      const license = providerLicense(metadata.LicenseShortName?.value || metadata.License?.value);
      return [{
        url: info.url,
        thumbnailUrl: info.thumburl,
        width: info.width,
        height: info.height,
        mime: info.mime,
        sourcePageUrl: `https://commons.wikimedia.org/wiki/${encodeURIComponent(page.title)}`,
        title: String(page.title || "").replace(/^File:/u, ""),
        author: stripProviderHtml(metadata.Artist?.value),
        license,
        licenseUrl: metadata.LicenseUrl?.value,
        confidence: license === "UNKNOWN" ? 0.2 : 0.95,
      }];
    });
  },
});

async function defaultProviderImplementations() {
  return { openverse: openverseProvider, wikimedia: wikimediaProvider };
}

export async function searchImageCandidates(input = {}, options = {}) {
  const query = text(input.query, "Image search query");
  const kind = enumValue(input.kind, KINDS, "Image kind");
  const purpose = enumValue(input.purpose, PURPOSES, "Image purpose");
  const orientation = enumValue(input.orientation, ORIENTATIONS, "Image orientation");
  const max = maximum(input.max);
  if (kind === "icon") {
    const candidates = orientation === "square" ? await searchLucideIcons(query, { maximum: max }) : [];
    return {
      query,
      kind,
      purpose,
      orientation,
      candidates,
      rejected: orientation === "square" ? [] : [{ provider: "lucide", reason: "image-orientation-mismatch", requestedOrientation: orientation }],
      providerReports: [{ provider: "lucide", ok: true, count: candidates.length, network: false }],
    };
  }

  const implementations = options.providerImplementations || await defaultProviderImplementations();
  const requestedProviders = options.providers || ["openverse", "wikimedia"];
  const candidates = [];
  const rejected = [];
  const providerReports = [];
  for (const providerName of requestedProviders) {
    const provider = implementations[providerName];
    if (!provider || typeof provider.search !== "function") throw imageError("invalid-image-provider", `Image provider ${providerName} is not available.`);
    const started = performance.now();
    try {
      const raw = await provider.search(query, {
        maxPerProvider: Math.max(max * 3, 10),
        safeSearch: "strict",
        licensePolicy: "open-only",
        timeoutMs: 12_000,
        fetcher: options.fetcher,
        auth: providerName === "wikimedia" ? { userAgent: "OfficeKit/2.0 (image sourcing; https://github.com/w31r4/OfficeKit)" } : undefined,
        signal: options.signal,
      });
      let accepted = 0;
      for (const candidate of raw) {
        try {
          candidates.push(normalizeProviderCandidate(candidate, { provider: providerName, kind, orientation, query }));
          accepted += 1;
        } catch (error) {
          rejected.push(rejection(candidate, providerName, error?.code || "invalid-provider-candidate"));
        }
      }
      providerReports.push({ provider: providerName, ok: true, count: accepted, rejected: raw.length - accepted, timeMs: Math.round(performance.now() - started), network: true });
    } catch (error) {
      providerReports.push({ provider: providerName, ok: false, count: 0, timeMs: Math.round(performance.now() - started), network: true, error: boundedImageError(error) });
    }
  }

  const seen = new Set();
  const ranked = candidates
    .sort((left, right) => right.score - left.score || left.provider.localeCompare(right.provider) || stableCandidateKey(left).localeCompare(stableCandidateKey(right)))
    .filter((candidate) => {
      const key = stableCandidateKey(candidate);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .slice(0, max);
  return { query, kind, purpose, orientation, candidates: ranked, rejected, providerReports };
}
