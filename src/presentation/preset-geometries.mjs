import presetGeometryManifest from "../ppj/preset-geometry-profiles.json" with { type: "json" };

// DrawingML preset geometry is a finite, public vocabulary.  Keep the
// catalog shared by the model and codec so imported preset shapes do not lose
// their native geometry merely because a caller edits an unrelated leaf.
const PRESET_GEOMETRY_PROFILES = Object.freeze(presetGeometryManifest.profiles || {});
const MIN_PRESET_ADJUSTMENT = Number(presetGeometryManifest.minimumValue ?? -21_600_000);
const MAX_PRESET_ADJUSTMENT = Number(presetGeometryManifest.maximumValue ?? 21_600_000);
const MAX_PRESET_ADJUSTMENTS = 256;

export function isPresentationPresetGeometry(value) {
  return typeof value === "string" && Object.hasOwn(PRESET_GEOMETRY_PROFILES, value);
}

export function normalizePresentationPresetAdjustments(value, geometry, label = "Presentation preset adjustments") {
  if (value == null) return [];
  if (!Array.isArray(value) || value.length > MAX_PRESET_ADJUSTMENTS) {
    throw new RangeError(`${label} must contain at most ${MAX_PRESET_ADJUSTMENTS} entries.`);
  }
  if (!isPresentationPresetGeometry(geometry)) {
    if (value.length) throw new TypeError(`${label} require a recognized preset geometry.`);
    return [];
  }
  const guideCount = PRESET_GEOMETRY_PROFILES[geometry].guides.length;
  if (value.length !== 0 && value.length !== guideCount) {
    throw new RangeError(`${label} for ${geometry} must contain either zero or ${guideCount} ordered values.`);
  }
  return value.map((entry, index) => {
    const number = Number(entry);
    if (!Number.isSafeInteger(number) || number < MIN_PRESET_ADJUSTMENT || number > MAX_PRESET_ADJUSTMENT) {
      throw new RangeError(`${label}[${index}] must be an integer from ${MIN_PRESET_ADJUSTMENT} through ${MAX_PRESET_ADJUSTMENT}.`);
    }
    return number;
  });
}

export function presentationPresetGeometryNames() {
  return Object.keys(PRESET_GEOMETRY_PROFILES);
}
