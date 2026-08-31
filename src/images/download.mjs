import { lookup as dnsLookup } from "node:dns/promises";
import { request as httpsRequest } from "node:https";
import { BlockList, isIP } from "node:net";

import { inspectImageBytes, normalizeImageMimeType } from "../shared/image-bytes.mjs";
import { imageError } from "./errors.mjs";
import { MAX_IMAGE_BYTES, MAX_IMAGE_DIMENSION, MAX_IMAGE_PIXELS } from "./task-assets.mjs";

const MAX_REDIRECTS = 3;
const DEFAULT_TIMEOUT_MS = 15_000;
const REMOTE_MIME_TYPES = new Set(["image/png", "image/jpeg", "image/gif"]);
const BLOCKED_HOSTS = new Set([
  "localhost",
  "localhost.localdomain",
  "metadata",
  "metadata.google.internal",
]);

const BLOCKED_IPV4 = new BlockList();
for (const [address, prefix] of [
  ["0.0.0.0", 8],
  ["10.0.0.0", 8],
  ["100.64.0.0", 10],
  ["127.0.0.0", 8],
  ["169.254.0.0", 16],
  ["172.16.0.0", 12],
  ["192.0.0.0", 24],
  ["192.0.2.0", 24],
  ["192.168.0.0", 16],
  ["198.18.0.0", 15],
  ["198.51.100.0", 24],
  ["203.0.113.0", 24],
  ["224.0.0.0", 4],
  ["240.0.0.0", 4],
]) BLOCKED_IPV4.addSubnet(address, prefix, "ipv4");

const BLOCKED_IPV6 = new BlockList();
for (const [address, prefix] of [
  ["::", 128],
  ["::1", 128],
  ["::ffff:0:0", 96],
  ["fc00::", 7],
  ["fe80::", 10],
  ["2001:db8::", 32],
  ["ff00::", 8],
]) BLOCKED_IPV6.addSubnet(address, prefix, "ipv6");

function parsedHttpsUrl(value) {
  let url;
  try { url = new URL(String(value || "")); }
  catch { throw imageError("invalid-image-url", "Image URL must be valid HTTPS."); }
  const hostname = url.hostname.toLowerCase().replace(/\.$/u, "");
  if (url.protocol !== "https:" || url.username || url.password || !hostname) {
    throw imageError("invalid-image-url", "Image URL must use HTTPS without credentials.");
  }
  if (BLOCKED_HOSTS.has(hostname) || hostname.endsWith(".localhost") || hostname.endsWith(".local") || hostname.endsWith(".internal")) {
    throw imageError("unsafe-image-destination", `Image host ${hostname} is not public.`);
  }
  return url;
}

function addressFamily(value) {
  const family = isIP(value);
  if (family === 4) return "ipv4";
  if (family === 6) return "ipv6";
  throw imageError("unsafe-image-destination", `Resolved image address ${value} is invalid.`);
}

function assertPublicAddress(address) {
  const family = addressFamily(address);
  const blocked = family === "ipv4"
    ? BLOCKED_IPV4.check(address, "ipv4")
    : BLOCKED_IPV6.check(address, "ipv6");
  if (blocked) {
    throw imageError("unsafe-image-destination", `Resolved image address ${address} is not public.`);
  }
  return family === "ipv4" ? 4 : 6;
}

async function resolvePublicAddresses(hostname, resolver) {
  if (isIP(hostname)) return [{ address: hostname, family: assertPublicAddress(hostname) }];
  const resolved = await resolver(hostname);
  const addresses = Array.isArray(resolved) ? resolved : [resolved];
  if (!addresses.length) throw imageError("unsafe-image-destination", `Image host ${hostname} did not resolve.`);
  const normalized = addresses.map((entry) => {
    const address = typeof entry === "string" ? entry : entry?.address;
    return { address, family: assertPublicAddress(address) };
  });
  return normalized;
}

function requestOnce(url, address, options) {
  const requestFactory = options.requestFactory || httpsRequest;
  return new Promise((resolve, reject) => {
    const request = requestFactory(url, {
      method: "GET",
      headers: {
        Accept: "image/png,image/jpeg,image/gif",
        "User-Agent": "OfficeKit/2.0 (image acquisition; https://github.com/w31r4/OfficeKit)",
      },
      lookup(_hostname, _lookupOptions, callback) {
        if (_lookupOptions?.all) callback(null, [{ address: address.address, family: address.family }]);
        else callback(null, address.address, address.family);
      },
      servername: isIP(url.hostname) ? undefined : url.hostname,
    }, resolve);
    request.setTimeout(options.timeoutMs, () => request.destroy(imageError("image-download-timeout", "Image download timed out.")));
    request.once("error", reject);
    request.end();
  });
}

async function responseBytes(response, maximum) {
  const length = Number(response.headers?.["content-length"]);
  if (Number.isFinite(length) && length > maximum) {
    response.destroy();
    throw imageError("image-download-too-large", `Remote image exceeds ${maximum} bytes.`);
  }
  const chunks = [];
  let total = 0;
  for await (const chunk of response) {
    total += chunk.length;
    if (total > maximum) {
      response.destroy();
      throw imageError("image-download-too-large", `Remote image exceeds ${maximum} bytes.`);
    }
    chunks.push(Buffer.from(chunk));
  }
  return Buffer.concat(chunks, total);
}

export async function downloadRemoteImage(value, options = {}) {
  const resolver = options.resolver || (async (hostname) => dnsLookup(hostname, { all: true, verbatim: true }));
  const timeoutMs = Number(options.timeoutMs ?? DEFAULT_TIMEOUT_MS);
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0 || timeoutMs > 120_000) throw imageError("invalid-image-download-limit", "Image download timeout is invalid.");
  let url = parsedHttpsUrl(value);
  const redirects = [];
  for (let redirectCount = 0; ; redirectCount += 1) {
    const addresses = await resolvePublicAddresses(url.hostname, resolver);
    const response = await requestOnce(url, addresses[0], { ...options, timeoutMs });
    const status = Number(response.statusCode || 0);
    if ([301, 302, 303, 307, 308].includes(status)) {
      response.resume();
      if (redirectCount >= MAX_REDIRECTS) throw imageError("image-redirect-limit", `Remote image exceeds ${MAX_REDIRECTS} redirects.`);
      const location = response.headers?.location;
      if (!location) throw imageError("invalid-image-response", "Image redirect has no Location header.");
      const next = parsedHttpsUrl(new URL(location, url).href);
      redirects.push({ from: url.href, to: next.href, status });
      url = next;
      continue;
    }
    if (status !== 200) {
      response.resume();
      throw imageError("image-download-http", `Image download returned HTTP ${status || "unknown"}.`);
    }
    const declaredMimeType = normalizeImageMimeType(response.headers?.["content-type"]);
    if (!REMOTE_MIME_TYPES.has(declaredMimeType)) {
      response.resume();
      throw imageError("unsupported-remote-image", `Remote image content type ${declaredMimeType || "(missing)"} is not PNG, JPEG, or GIF.`);
    }
    const bytes = await responseBytes(response, MAX_IMAGE_BYTES);
    const inspected = inspectImageBytes(bytes, {
      declaredMimeType,
      allowSvg: false,
      label: "Remote image",
      maxBytes: MAX_IMAGE_BYTES,
      maxPixels: MAX_IMAGE_PIXELS,
      maxDimension: MAX_IMAGE_DIMENSION,
    });
    return Object.freeze({ bytes, ...inspected, finalUrl: url.href, redirects });
  }
}

export async function assertPublicImageUrl(value, { resolver } = {}) {
  const url = parsedHttpsUrl(value);
  await resolvePublicAddresses(url.hostname, resolver || (async (hostname) => dnsLookup(hostname, { all: true, verbatim: true })));
  return url.href;
}
