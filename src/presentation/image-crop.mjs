import { inspectImageBytes } from "../shared/image-bytes.mjs";

const IMAGE_FITS = new Set(["contain", "cover", "stretch"]);
const MAX_CROP = 1;
const CROP_SCALE = 100_000;

function finite(value, label) {
  const number = Number(value ?? 0);
  if (!Number.isFinite(number)) throw new TypeError(`${label} must be finite.`);
  return number;
}

export function presentationImageDataUrlDimensions(value) {
  const match = /^data:(image\/(?:png|jpe?g|gif|svg\+xml));base64,([A-Za-z0-9+/=\s]+)$/i.exec(String(value || ""));
  if (!match) throw new TypeError("Presentation image contain/cover fitting requires an embedded base64 PNG, JPEG, GIF, or SVG dataUrl.");
  const bytes = Buffer.from(match[2].replace(/\s/g, ""), "base64");
  const type = match[1].toLowerCase().replace("image/jpg", "image/jpeg");
  let parsed;
  try {
    parsed = inspectImageBytes(bytes, { declaredMimeType: type, label: `Presentation ${type} image` });
  } catch (error) {
    if (type !== "image/svg+xml") throw new TypeError(`Presentation ${type} image does not expose bounded positive intrinsic dimensions.`, { cause: error });
    throw error;
  }
  return { width: parsed.width, height: parsed.height };
}

export function normalizePresentationImageFit(value = "contain") {
  const fit = String(value || "contain");
  if (!IMAGE_FITS.has(fit)) throw new TypeError("Presentation image fit must be contain, cover, or stretch.");
  return fit;
}

export function normalizePresentationImageCrop(value) {
  if (value == null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError("Presentation image crop must be an object.");
  const crop = {
    left: finite(value.left, "Presentation image crop.left"),
    top: finite(value.top, "Presentation image crop.top"),
    right: finite(value.right, "Presentation image crop.right"),
    bottom: finite(value.bottom, "Presentation image crop.bottom"),
  };
  if (Object.values(crop).some((edge) => edge < -MAX_CROP || edge > MAX_CROP) || crop.left + crop.right >= 1 || crop.top + crop.bottom >= 1) {
    throw new RangeError("Presentation image crop edges must be between -1 and 1 and opposing sums must remain below 1.");
  }
  return crop;
}

function normalizedFrame(frame = {}) {
  const width = Number(frame.width);
  const height = Number(frame.height);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0 || width > 10_000_000 || height > 10_000_000) throw new RangeError("Presentation image contain/cover fitting requires a positive bounded frame.");
  return { width, height };
}

function roundedCrop(crop) {
  const rounded = Object.fromEntries(Object.entries(crop).map(([key, value]) => [key, Math.round(value * CROP_SCALE) / CROP_SCALE]));
  return normalizePresentationImageCrop(rounded);
}

export function effectivePresentationImageCrop({ crop, fit = "contain", dataUrl, frame } = {}) {
  const normalizedFit = normalizePresentationImageFit(fit);
  const manual = normalizePresentationImageCrop(crop);
  if (normalizedFit === "stretch") return manual;
  const image = presentationImageDataUrlDimensions(dataUrl);
  const target = normalizedFrame(frame);
  const result = { ...(manual || { left: 0, top: 0, right: 0, bottom: 0 }) };
  const sourceWidth = 1 - result.left - result.right;
  const sourceHeight = 1 - result.top - result.bottom;
  const sourceAspect = (image.width * sourceWidth) / (image.height * sourceHeight);
  const targetAspect = target.width / target.height;
  if (Math.abs(sourceAspect - targetAspect) > 1e-12) {
    if (normalizedFit === "cover" && sourceAspect > targetAspect) {
      const desired = sourceHeight * targetAspect * image.height / image.width;
      const delta = (sourceWidth - desired) / 2;
      result.left += delta;
      result.right += delta;
    } else if (normalizedFit === "cover") {
      const desired = sourceWidth * image.width / (targetAspect * image.height);
      const delta = (sourceHeight - desired) / 2;
      result.top += delta;
      result.bottom += delta;
    } else if (sourceAspect > targetAspect) {
      const desired = sourceWidth * image.width / (targetAspect * image.height);
      const delta = (desired - sourceHeight) / 2;
      result.top -= delta;
      result.bottom -= delta;
    } else {
      const desired = sourceHeight * targetAspect * image.height / image.width;
      const delta = (desired - sourceWidth) / 2;
      result.left -= delta;
      result.right -= delta;
    }
  }
  const normalized = roundedCrop(result);
  if (manual == null && Object.values(normalized).every((edge) => edge === 0)) return undefined;
  return normalized;
}

export function presentationImageCropToWire(crop) {
  const normalized = normalizePresentationImageCrop(crop);
  if (!normalized) return undefined;
  return {
    leftThousandthPercent: Math.round(normalized.left * CROP_SCALE),
    topThousandthPercent: Math.round(normalized.top * CROP_SCALE),
    rightThousandthPercent: Math.round(normalized.right * CROP_SCALE),
    bottomThousandthPercent: Math.round(normalized.bottom * CROP_SCALE),
  };
}

export function presentationImageCropFromWire(crop) {
  if (!crop) return undefined;
  return normalizePresentationImageCrop({
    left: Number(crop.leftThousandthPercent) / CROP_SCALE,
    top: Number(crop.topThousandthPercent) / CROP_SCALE,
    right: Number(crop.rightThousandthPercent) / CROP_SCALE,
    bottom: Number(crop.bottomThousandthPercent) / CROP_SCALE,
  });
}

export function presentationImageCropViewport({ crop, fit, dataUrl, frame } = {}) {
  const effective = effectivePresentationImageCrop({ crop, fit, dataUrl, frame });
  if (!effective) return undefined;
  const image = presentationImageDataUrlDimensions(dataUrl);
  return {
    x: effective.left * image.width,
    y: effective.top * image.height,
    width: (1 - effective.left - effective.right) * image.width,
    height: (1 - effective.top - effective.bottom) * image.height,
    imageWidth: image.width,
    imageHeight: image.height,
  };
}
