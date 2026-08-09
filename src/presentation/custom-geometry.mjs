import {
  evaluatePresentationCustomGeometryFormulaGraph,
  normalizePresentationCustomGeometryFormulaGraph,
  normalizePresentationCustomGeometryReference,
  presentationCustomGeometryReferenceNames,
  resolvePresentationCustomGeometryReference,
} from "./custom-geometry-formulas.mjs";

const MAX_PATHS = 64;
const MAX_COMMANDS = 16_384;
const MAX_CONNECTION_SITES = 1_024;
const MAX_ADJUSTMENT_HANDLES = 1_024;
const MAX_COORDINATE = 2_147_483_647;
const EMU_PER_PIXEL = 9_525;
const ANGLE_UNITS_PER_DEGREE = 60_000;
const HALF_TURN_ANGLE = 180 * ANGLE_UNITS_PER_DEGREE;
const FULL_TURN_ANGLE = 360 * ANGLE_UNITS_PER_DEGREE;
const ARC_FIELDS = Object.freeze(["widthRadius", "heightRadius", "startAngle", "sweepAngle"]);
const PATH_FIELDS = new Set(["width", "height", "commands", "fillMode", "stroke", "extrusionAllowed"]);
const PATH_FILL_MODES = new Set(["normal", "none"]);
const TEXT_RECTANGLE_FIELDS = Object.freeze(["left", "top", "right", "bottom"]);
const TEXT_RECTANGLE_FIELD_SET = new Set(TEXT_RECTANGLE_FIELDS);
const CURVE_FIELDS = Object.freeze({
  quadraticBezTo: Object.freeze(["x1", "y1", "x", "y"]),
  cubicBezTo: Object.freeze(["x1", "y1", "x2", "y2", "x", "y"]),
});
const CONNECTION_SITE_FIELDS = new Set(["angle", "x", "y"]);
const XY_ADJUSTMENT_HANDLE_FIELDS = new Set([
  "kind", "xAdjustment", "minX", "maxX", "yAdjustment", "minY", "maxY", "x", "y",
]);
const POLAR_ADJUSTMENT_HANDLE_FIELDS = new Set([
  "kind", "radialAdjustment", "minRadius", "maxRadius", "angleAdjustment", "minAngle", "maxAngle", "x", "y",
]);

function coordinate(value, label, references) {
  if (typeof value === "string") return normalizePresentationCustomGeometryReference(value, references, label);
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number < -MAX_COORDINATE || number > MAX_COORDINATE) {
    throw new RangeError(`${label} must be a safe integer within the DrawingML signed 32-bit coordinate range.`);
  }
  return number;
}

function textRectangleCoordinate(value, label, references) {
  if (typeof value === "string") return normalizePresentationCustomGeometryReference(value, references, label);
  const number = Number(value);
  const emu = Math.round(number * EMU_PER_PIXEL);
  if (!Number.isFinite(number) || !Number.isSafeInteger(emu) || emu < -MAX_COORDINATE || emu > MAX_COORDINATE) {
    throw new RangeError(`${label} must be a finite pixel coordinate representable in the DrawingML signed 32-bit EMU range.`);
  }
  return number;
}

export function normalizePresentationCustomTextRectangle(value, { adjustments, guides, widthEmu, heightEmu } = {}) {
  if (value == null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError("Presentation custom geometry textRectangle must be an object.");
  const unknown = Object.keys(value).filter((key) => !TEXT_RECTANGLE_FIELD_SET.has(key));
  if (unknown.length) throw new TypeError(`Presentation custom geometry textRectangle has unsupported fields: ${unknown.join(", ")}.`);
  const graph = normalizePresentationCustomGeometryFormulaGraph({ adjustments, guides });
  const references = presentationCustomGeometryReferenceNames(graph, { includeBuiltins: true });
  const rectangle = Object.fromEntries(TEXT_RECTANGLE_FIELDS.map((field) => [
    field,
    textRectangleCoordinate(value[field], `Presentation custom geometry textRectangle.${field}`, references),
  ]));
  const hasReferences = TEXT_RECTANGLE_FIELDS.some((field) => typeof rectangle[field] === "string");
  const values = hasReferences
    ? evaluatePresentationCustomGeometryFormulaGraph(graph, { widthEmu, heightEmu })
    : undefined;
  const resolved = Object.fromEntries(TEXT_RECTANGLE_FIELDS.map((field) => [
    field,
    typeof rectangle[field] === "string"
      ? resolvePresentationCustomGeometryReference(rectangle[field], values, `Presentation custom geometry textRectangle.${field}`)
      : Math.round(rectangle[field] * EMU_PER_PIXEL),
  ]));
  if (resolved.left >= resolved.right) {
    throw new RangeError("Presentation custom geometry textRectangle.right must be greater than left at native EMU precision.");
  }
  if (resolved.top >= resolved.bottom) {
    throw new RangeError("Presentation custom geometry textRectangle.bottom must be greater than top at native EMU precision.");
  }
  return rectangle;
}

export function presentationCustomTextRectangleFrame(value, frame, sourceFrame = frame, graph = {}) {
  const sourceWidth = Number(sourceFrame?.width);
  const sourceHeight = Number(sourceFrame?.height);
  const formulaContext = {
    ...graph,
    widthEmu: Math.round(sourceWidth * EMU_PER_PIXEL),
    heightEmu: Math.round(sourceHeight * EMU_PER_PIXEL),
  };
  const rectangle = normalizePresentationCustomTextRectangle(value, formulaContext);
  if (!rectangle) return { ...frame };
  const left = Number(frame?.left);
  const top = Number(frame?.top);
  const width = Number(frame?.width);
  const height = Number(frame?.height);
  if (![sourceWidth, sourceHeight, left, top, width, height].every(Number.isFinite) || sourceWidth <= 0 || sourceHeight <= 0 || width <= 0 || height <= 0) {
    throw new RangeError("Presentation custom geometry textRectangle requires positive source and rendered shape frames.");
  }
  const values = evaluatePresentationCustomGeometryFormulaGraph(graph, formulaContext);
  const resolved = Object.fromEntries(TEXT_RECTANGLE_FIELDS.map((field) => [
    field,
    typeof rectangle[field] === "string"
      ? resolvePresentationCustomGeometryReference(rectangle[field], values, `Presentation custom geometry textRectangle.${field}`) / EMU_PER_PIXEL
      : rectangle[field],
  ]));
  const scaleX = width / sourceWidth;
  const scaleY = height / sourceHeight;
  return {
    left: left + resolved.left * scaleX,
    top: top + resolved.top * scaleY,
    width: (resolved.right - resolved.left) * scaleX,
    height: (resolved.bottom - resolved.top) * scaleY,
  };
}

function angle(value, label, references) {
  if (typeof value === "string") return normalizePresentationCustomGeometryReference(value, references, label);
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number < -MAX_COORDINATE || number > MAX_COORDINATE) {
    throw new RangeError(`${label} must be a safe integer within the DrawingML signed 32-bit angle range.`);
  }
  return number;
}

function connectionSiteAngle(value, label, references, values) {
  if (typeof value === "string") {
    const reference = normalizePresentationCustomGeometryReference(value, references, label);
    const resolved = resolvePresentationCustomGeometryReference(reference, values, label);
    if (Math.abs(resolved) > FULL_TURN_ANGLE) throw new RangeError(`${label} must evaluate within one full DrawingML turn.`);
    return reference;
  }
  const degrees = Number(value);
  const native = Math.round(degrees * ANGLE_UNITS_PER_DEGREE);
  if (!Number.isFinite(degrees) || !Number.isSafeInteger(native) || Math.abs(native) > FULL_TURN_ANGLE) {
    throw new RangeError(`${label} must be a finite degree value from -360 through 360.`);
  }
  return native / ANGLE_UNITS_PER_DEGREE;
}

function connectionSiteCoordinate(value, label, references, values, maximumEmu) {
  if (typeof value === "string") {
    const reference = normalizePresentationCustomGeometryReference(value, references, label);
    const resolved = resolvePresentationCustomGeometryReference(reference, values, label);
    if (resolved < 0 || resolved > maximumEmu) throw new RangeError(`${label} must evaluate inside the custom shape frame.`);
    return reference;
  }
  const pixels = Number(value);
  const emu = Math.round(pixels * EMU_PER_PIXEL);
  if (!Number.isFinite(pixels) || !Number.isSafeInteger(emu) || emu < 0 || emu > maximumEmu) {
    throw new RangeError(`${label} must be a finite pixel coordinate inside the custom shape frame.`);
  }
  return emu / EMU_PER_PIXEL;
}

export function normalizePresentationCustomConnectionSites(value, { adjustments, guides, widthEmu, heightEmu } = {}) {
  if (value == null) return [];
  if (!Array.isArray(value) || value.length > MAX_CONNECTION_SITES) {
    throw new RangeError(`Presentation custom geometry customConnectionSites must contain at most ${MAX_CONNECTION_SITES} entries.`);
  }
  if (value.length === 0) return [];
  const graph = normalizePresentationCustomGeometryFormulaGraph({ adjustments, guides });
  const references = presentationCustomGeometryReferenceNames(graph);
  const values = evaluatePresentationCustomGeometryFormulaGraph(graph, { widthEmu, heightEmu });
  return value.map((site, index) => {
    const label = `Presentation custom geometry connection site ${index + 1}`;
    if (!site || typeof site !== "object" || Array.isArray(site)) throw new TypeError(`${label} must be an object.`);
    const unknown = Object.keys(site).filter((key) => !CONNECTION_SITE_FIELDS.has(key));
    if (unknown.length) throw new TypeError(`${label} has unsupported fields: ${unknown.join(", ")}.`);
    return {
      angle: connectionSiteAngle(site.angle, `${label}.angle`, references, values),
      x: connectionSiteCoordinate(site.x, `${label}.x`, references, values, widthEmu),
      y: connectionSiteCoordinate(site.y, `${label}.y`, references, values, heightEmu),
    };
  });
}

function adjustmentHandleGuide(value, label, adjustmentNames) {
  if (typeof value !== "string" || !adjustmentNames.has(value)) {
    throw new ReferenceError(`${label} must name one declared custom adjustment.`);
  }
  return value;
}

function handleCoordinateBound(value, label, references, values) {
  const normalized = coordinate(value, label, references);
  return { normalized, resolved: resolvePresentationCustomGeometryReference(normalized, values, label) };
}

function handleAngleBound(value, label, references, values) {
  const normalized = connectionSiteAngle(value, label, references, values);
  const resolved = typeof normalized === "string"
    ? resolvePresentationCustomGeometryReference(normalized, values, label)
    : normalized * ANGLE_UNITS_PER_DEGREE;
  return { normalized, resolved };
}

function adjustmentHandleRange(handle, output, {
  adjustmentField,
  minimumField,
  maximumField,
  label,
  adjustmentNames,
  references,
  values,
  normalizeBound,
  requireNonNegative = false,
  maximumAbsolute,
}) {
  const hasAdjustment = Object.hasOwn(handle, adjustmentField) && handle[adjustmentField] != null;
  const hasMinimum = Object.hasOwn(handle, minimumField) && handle[minimumField] != null;
  const hasMaximum = Object.hasOwn(handle, maximumField) && handle[maximumField] != null;
  if (!hasAdjustment) {
    if (hasMinimum || hasMaximum) throw new TypeError(`${label} bounds require ${adjustmentField}.`);
    return false;
  }
  const adjustment = adjustmentHandleGuide(handle[adjustmentField], `${label}.${adjustmentField}`, adjustmentNames);
  output[adjustmentField] = adjustment;
  const current = resolvePresentationCustomGeometryReference(adjustment, values, `${label}.${adjustmentField}`);
  if (requireNonNegative && current < 0) throw new RangeError(`${label}.${adjustmentField} must evaluate to a non-negative value.`);
  if (maximumAbsolute !== undefined && Math.abs(current) > maximumAbsolute) {
    throw new RangeError(`${label}.${adjustmentField} must evaluate within one full DrawingML turn.`);
  }
  if (hasMinimum !== hasMaximum) {
    throw new TypeError(`${label}.${minimumField} and ${maximumField} must be supplied together.`);
  }
  if (!hasMinimum) return true;
  const minimum = normalizeBound(handle[minimumField], `${label}.${minimumField}`, references, values);
  const maximum = normalizeBound(handle[maximumField], `${label}.${maximumField}`, references, values);
  if (requireNonNegative && (minimum.resolved < 0 || maximum.resolved < 0)) {
    throw new RangeError(`${label}.${minimumField} and ${maximumField} must evaluate to non-negative values.`);
  }
  if (minimum.resolved > maximum.resolved) {
    throw new RangeError(`${label}.${maximumField} must evaluate greater than or equal to ${minimumField}.`);
  }
  if (current < minimum.resolved || current > maximum.resolved) {
    throw new RangeError(`${label}.${adjustmentField} must evaluate inside its ${minimumField}/${maximumField} range.`);
  }
  output[minimumField] = minimum.normalized;
  output[maximumField] = maximum.normalized;
  return true;
}

export function normalizePresentationCustomAdjustmentHandles(value, { adjustments, guides, widthEmu, heightEmu } = {}) {
  if (value == null) return [];
  if (!Array.isArray(value) || value.length > MAX_ADJUSTMENT_HANDLES) {
    throw new RangeError(`Presentation custom geometry customAdjustmentHandles must contain at most ${MAX_ADJUSTMENT_HANDLES} entries.`);
  }
  if (value.length === 0) return [];
  const graph = normalizePresentationCustomGeometryFormulaGraph({ adjustments, guides });
  const references = presentationCustomGeometryReferenceNames(graph);
  const adjustmentNames = new Set(graph.adjustments.map((adjustment) => adjustment.name));
  const values = evaluatePresentationCustomGeometryFormulaGraph(graph, { widthEmu, heightEmu });
  return value.map((handle, index) => {
    const label = `Presentation custom geometry adjustment handle ${index + 1}`;
    if (!handle || typeof handle !== "object" || Array.isArray(handle)) throw new TypeError(`${label} must be an object.`);
    if (handle.kind !== "xy" && handle.kind !== "polar") throw new TypeError(`${label}.kind must be xy or polar.`);
    const allowed = handle.kind === "xy" ? XY_ADJUSTMENT_HANDLE_FIELDS : POLAR_ADJUSTMENT_HANDLE_FIELDS;
    const unknown = Object.keys(handle).filter((key) => !allowed.has(key));
    if (unknown.length) throw new TypeError(`${label} has unsupported fields: ${unknown.join(", ")}.`);
    const normalized = { kind: handle.kind };
    let controlled = false;
    if (handle.kind === "xy") {
      controlled = adjustmentHandleRange(handle, normalized, {
        adjustmentField: "xAdjustment", minimumField: "minX", maximumField: "maxX", label,
        adjustmentNames, references, values, normalizeBound: handleCoordinateBound,
      }) || controlled;
      controlled = adjustmentHandleRange(handle, normalized, {
        adjustmentField: "yAdjustment", minimumField: "minY", maximumField: "maxY", label,
        adjustmentNames, references, values, normalizeBound: handleCoordinateBound,
      }) || controlled;
    } else {
      controlled = adjustmentHandleRange(handle, normalized, {
        adjustmentField: "radialAdjustment", minimumField: "minRadius", maximumField: "maxRadius", label,
        adjustmentNames, references, values, normalizeBound: handleCoordinateBound, requireNonNegative: true,
      }) || controlled;
      controlled = adjustmentHandleRange(handle, normalized, {
        adjustmentField: "angleAdjustment", minimumField: "minAngle", maximumField: "maxAngle", label,
        adjustmentNames, references, values, normalizeBound: handleAngleBound, maximumAbsolute: FULL_TURN_ANGLE,
      }) || controlled;
    }
    if (!controlled) throw new TypeError(`${label} must control at least one declared custom adjustment.`);
    normalized.x = connectionSiteCoordinate(handle.x, `${label}.x`, references, values, widthEmu);
    normalized.y = connectionSiteCoordinate(handle.y, `${label}.y`, references, values, heightEmu);
    return normalized;
  });
}

export function presentationCustomConnectionSitePoint(connectionSites, index, frame, { adjustments, guides } = {}) {
  const left = Number(frame?.left);
  const top = Number(frame?.top);
  const width = Number(frame?.width);
  const height = Number(frame?.height);
  if (![left, top, width, height].every(Number.isFinite) || width <= 0 || height <= 0) {
    throw new RangeError("Presentation custom geometry connection-site routing requires a positive finite shape frame.");
  }
  const widthEmu = Math.round(width * EMU_PER_PIXEL);
  const heightEmu = Math.round(height * EMU_PER_PIXEL);
  const sites = normalizePresentationCustomConnectionSites(connectionSites, { adjustments, guides, widthEmu, heightEmu });
  if (!Number.isInteger(index) || index < 0 || index >= sites.length) {
    throw new RangeError(`Presentation custom geometry connection-site index ${index} is outside the modeled range 0..${Math.max(0, sites.length - 1)}.`);
  }
  const values = evaluatePresentationCustomGeometryFormulaGraph({ adjustments, guides }, { widthEmu, heightEmu });
  const site = sites[index];
  const x = typeof site.x === "string" ? resolvePresentationCustomGeometryReference(site.x, values, `Presentation custom geometry connection site ${index + 1}.x`) / EMU_PER_PIXEL : site.x;
  const y = typeof site.y === "string" ? resolvePresentationCustomGeometryReference(site.y, values, `Presentation custom geometry connection site ${index + 1}.y`) / EMU_PER_PIXEL : site.y;
  return { x: left + x, y: top + y };
}

function point(value, label, references) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const unknown = Object.keys(value).filter((key) => key !== "x" && key !== "y");
  if (unknown.length) throw new TypeError(`${label} has unsupported fields: ${unknown.join(", ")}.`);
  return { x: coordinate(value.x, `${label}.x`, references), y: coordinate(value.y, `${label}.y`, references) };
}

function curve(value, label, fields, references) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const allowed = new Set(fields);
  const unknown = Object.keys(value).filter((key) => !allowed.has(key));
  if (unknown.length) throw new TypeError(`${label} has unsupported fields: ${unknown.join(", ")}.`);
  return Object.fromEntries(fields.map((field) => [field, coordinate(value[field], `${label}.${field}`, references)]));
}

function arc(value, label, references, values) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const allowed = new Set(ARC_FIELDS);
  const unknown = Object.keys(value).filter((key) => !allowed.has(key));
  if (unknown.length) throw new TypeError(`${label} has unsupported fields: ${unknown.join(", ")}.`);
  const widthRadius = coordinate(value.widthRadius, `${label}.widthRadius`, references);
  const heightRadius = coordinate(value.heightRadius, `${label}.heightRadius`, references);
  const startAngle = angle(value.startAngle, `${label}.startAngle`, references);
  const sweepAngle = angle(value.sweepAngle, `${label}.sweepAngle`, references);
  const resolvedWidth = resolvePresentationCustomGeometryReference(widthRadius, values, `${label}.widthRadius`);
  const resolvedHeight = resolvePresentationCustomGeometryReference(heightRadius, values, `${label}.heightRadius`);
  const resolvedSweep = resolvePresentationCustomGeometryReference(sweepAngle, values, `${label}.sweepAngle`);
  if (resolvedWidth <= 0 || resolvedHeight <= 0) throw new RangeError(`${label} radii must evaluate to positive values.`);
  if (resolvedSweep === 0 || Math.abs(resolvedSweep) > FULL_TURN_ANGLE) {
    throw new RangeError(`${label}.sweepAngle must be non-zero and no greater than one full DrawingML turn (${FULL_TURN_ANGLE}).`);
  }
  return { widthRadius, heightRadius, startAngle, sweepAngle };
}

function command(value, pathIndex, commandIndex, references, values) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`Presentation custom path ${pathIndex + 1} command ${commandIndex + 1} must be an object.`);
  const keys = Object.keys(value);
  if (keys.length !== 1) throw new TypeError(`Presentation custom path ${pathIndex + 1} command ${commandIndex + 1} must contain exactly one command.`);
  const label = `Presentation custom path ${pathIndex + 1} command ${commandIndex + 1}`;
  if (keys[0] === "moveTo" || keys[0] === "lineTo") return { [keys[0]]: point(value[keys[0]], `${label}.${keys[0]}`, references) };
  if (keys[0] === "arcTo") return { arcTo: arc(value.arcTo, `${label}.arcTo`, references, values) };
  const curveFields = Object.hasOwn(CURVE_FIELDS, keys[0]) ? CURVE_FIELDS[keys[0]] : undefined;
  if (curveFields) return { [keys[0]]: curve(value[keys[0]], `${label}.${keys[0]}`, curveFields, references) };
  if (keys[0] === "close") {
    if (value.close !== true && (typeof value.close !== "object" || value.close == null || Array.isArray(value.close) || Object.keys(value.close).length)) {
      throw new TypeError(`${label}.close must be true or an empty object.`);
    }
    return { close: {} };
  }
  throw new TypeError(`${label} uses unsupported command ${keys[0]}.`);
}

export function normalizePresentationCustomPaths(value, { adjustments, guides, widthEmu = 1, heightEmu = 1 } = {}) {
  if (value == null) return [];
  if (!Array.isArray(value) || value.length === 0 || value.length > MAX_PATHS) throw new RangeError(`Presentation custom geometry must contain 1 through ${MAX_PATHS} paths.`);
  const graph = normalizePresentationCustomGeometryFormulaGraph({ adjustments, guides });
  const references = presentationCustomGeometryReferenceNames(graph);
  const values = evaluatePresentationCustomGeometryFormulaGraph(graph, { widthEmu, heightEmu });
  let commandCount = 0;
  return value.map((path, pathIndex) => {
    if (!path || typeof path !== "object" || Array.isArray(path)) throw new TypeError(`Presentation custom path ${pathIndex + 1} must be an object.`);
    const unknown = Object.keys(path).filter((key) => !PATH_FIELDS.has(key));
    if (unknown.length) throw new TypeError(`Presentation custom path ${pathIndex + 1} has unsupported fields: ${unknown.join(", ")}.`);
    const width = coordinate(path.width, `Presentation custom path ${pathIndex + 1}.width`, references);
    const height = coordinate(path.height, `Presentation custom path ${pathIndex + 1}.height`, references);
    if (typeof width !== "number" || typeof height !== "number") throw new TypeError(`Presentation custom path ${pathIndex + 1} width and height must be literal coordinates.`);
    if (width <= 0 || height <= 0) throw new RangeError(`Presentation custom path ${pathIndex + 1} width and height must be positive.`);
    if (!Array.isArray(path.commands) || path.commands.length === 0) throw new TypeError(`Presentation custom path ${pathIndex + 1} requires commands.`);
    commandCount += path.commands.length;
    if (commandCount > MAX_COMMANDS) throw new RangeError(`Presentation custom geometry exceeds the ${MAX_COMMANDS}-command budget.`);
    let hasCurrentPoint = false;
    let hasSubpathStart = false;
    const commands = path.commands.map((item, commandIndex) => {
      const normalized = command(item, pathIndex, commandIndex, references, values);
      const label = `Presentation custom path ${pathIndex + 1} command ${commandIndex + 1}`;
      if (normalized.arcTo && !hasCurrentPoint) throw new RangeError(`${label}.arcTo requires an established current point.`);
      if (normalized.moveTo) {
        hasCurrentPoint = true;
        hasSubpathStart = true;
      } else if (normalized.lineTo || normalized.quadraticBezTo || normalized.cubicBezTo) {
        hasCurrentPoint = true;
      } else if (normalized.close) {
        hasCurrentPoint = hasSubpathStart;
      }
      return normalized;
    });
    const normalized = { width, height, commands };
    if (Object.hasOwn(path, "fillMode")) {
      if (!PATH_FILL_MODES.has(path.fillMode)) throw new TypeError(`Presentation custom path ${pathIndex + 1}.fillMode must be normal or none.`);
      normalized.fillMode = path.fillMode;
    }
    for (const field of ["stroke", "extrusionAllowed"]) {
      if (!Object.hasOwn(path, field)) continue;
      if (typeof path[field] !== "boolean") throw new TypeError(`Presentation custom path ${pathIndex + 1}.${field} must be a boolean.`);
      normalized[field] = path[field];
    }
    return normalized;
  });
}

function svgNumber(value) {
  const rounded = Number(value.toFixed(9));
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

function angleRadians(value) {
  const normalized = ((value % FULL_TURN_ANGLE) + FULL_TURN_ANGLE) % FULL_TURN_ANGLE;
  return normalized / ANGLE_UNITS_PER_DEGREE * Math.PI / 180;
}

function svgArcCommands(arcTo, currentPoint) {
  const start = angleRadians(arcTo.startAngle);
  const center = {
    x: currentPoint.x - arcTo.widthRadius * Math.cos(start),
    y: currentPoint.y - arcTo.heightRadius * Math.sin(start),
  };
  const commands = [];
  let angleValue = arcTo.startAngle;
  let remaining = arcTo.sweepAngle;
  let end = currentPoint;
  while (remaining !== 0) {
    const magnitude = Math.min(Math.abs(remaining), HALF_TURN_ANGLE);
    const segment = Math.sign(remaining) * magnitude;
    angleValue += segment;
    const radians = angleRadians(angleValue);
    end = {
      x: center.x + arcTo.widthRadius * Math.cos(radians),
      y: center.y + arcTo.heightRadius * Math.sin(radians),
    };
    commands.push(`A ${arcTo.widthRadius} ${arcTo.heightRadius} 0 0 ${segment > 0 ? 1 : 0} ${svgNumber(end.x)} ${svgNumber(end.y)}`);
    remaining -= segment;
  }
  return { commands, end };
}

export function presentationCustomPathsSvg(paths, frame, { escape = String, adjustments, guides, sourceFrame = frame } = {}) {
  const widthEmu = Math.round(Number(sourceFrame?.width) * EMU_PER_PIXEL);
  const heightEmu = Math.round(Number(sourceFrame?.height) * EMU_PER_PIXEL);
  const graph = normalizePresentationCustomGeometryFormulaGraph({ adjustments, guides });
  const values = evaluatePresentationCustomGeometryFormulaGraph(graph, { widthEmu, heightEmu });
  const resolvedPoint = (value, label) => ({
    x: resolvePresentationCustomGeometryReference(value.x, values, `${label}.x`),
    y: resolvePresentationCustomGeometryReference(value.y, values, `${label}.y`),
  });
  return normalizePresentationCustomPaths(paths, { ...graph, widthEmu, heightEmu }).map((path, pathIndex) => {
    const chunks = [];
    let currentPoint;
    let subpathStart;
    for (const item of path.commands) {
      if (item.moveTo) {
        const point = resolvedPoint(item.moveTo, `Presentation custom path ${pathIndex + 1} moveTo`);
        chunks.push(`M ${svgNumber(point.x)} ${svgNumber(point.y)}`);
        currentPoint = point;
        subpathStart = point;
      } else if (item.lineTo) {
        const point = resolvedPoint(item.lineTo, `Presentation custom path ${pathIndex + 1} lineTo`);
        chunks.push(`L ${svgNumber(point.x)} ${svgNumber(point.y)}`);
        currentPoint = point;
      } else if (item.quadraticBezTo) {
        const control = resolvedPoint({ x: item.quadraticBezTo.x1, y: item.quadraticBezTo.y1 }, `Presentation custom path ${pathIndex + 1} quadratic control`);
        const end = resolvedPoint(item.quadraticBezTo, `Presentation custom path ${pathIndex + 1} quadratic end`);
        chunks.push(`Q ${svgNumber(control.x)} ${svgNumber(control.y)} ${svgNumber(end.x)} ${svgNumber(end.y)}`);
        currentPoint = end;
      } else if (item.cubicBezTo) {
        const control1 = resolvedPoint({ x: item.cubicBezTo.x1, y: item.cubicBezTo.y1 }, `Presentation custom path ${pathIndex + 1} cubic control 1`);
        const control2 = resolvedPoint({ x: item.cubicBezTo.x2, y: item.cubicBezTo.y2 }, `Presentation custom path ${pathIndex + 1} cubic control 2`);
        const end = resolvedPoint(item.cubicBezTo, `Presentation custom path ${pathIndex + 1} cubic end`);
        chunks.push(`C ${svgNumber(control1.x)} ${svgNumber(control1.y)} ${svgNumber(control2.x)} ${svgNumber(control2.y)} ${svgNumber(end.x)} ${svgNumber(end.y)}`);
        currentPoint = end;
      } else if (item.arcTo) {
        const arc = svgArcCommands({
          widthRadius: resolvePresentationCustomGeometryReference(item.arcTo.widthRadius, values, `Presentation custom path ${pathIndex + 1} arcTo.widthRadius`),
          heightRadius: resolvePresentationCustomGeometryReference(item.arcTo.heightRadius, values, `Presentation custom path ${pathIndex + 1} arcTo.heightRadius`),
          startAngle: resolvePresentationCustomGeometryReference(item.arcTo.startAngle, values, `Presentation custom path ${pathIndex + 1} arcTo.startAngle`),
          sweepAngle: resolvePresentationCustomGeometryReference(item.arcTo.sweepAngle, values, `Presentation custom path ${pathIndex + 1} arcTo.sweepAngle`),
        }, currentPoint);
        chunks.push(...arc.commands);
        currentPoint = arc.end;
      } else {
        chunks.push("Z");
        currentPoint = subpathStart ? { ...subpathStart } : undefined;
      }
    }
    const commands = chunks.join(" ");
    const paint = [
      path.fillMode === "none" ? ' fill="none"' : "",
      path.stroke === false ? ' stroke="none"' : "",
    ].join("");
    return `<path d="${escape(commands)}" transform="translate(${frame.left} ${frame.top}) scale(${frame.width / path.width} ${frame.height / path.height})"${paint}/>`;
  }).join("");
}
