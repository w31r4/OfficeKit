const MAX_PATHS = 64;
const MAX_COMMANDS = 16_384;
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

function coordinate(value, label) {
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number < -MAX_COORDINATE || number > MAX_COORDINATE) {
    throw new RangeError(`${label} must be a safe integer within the DrawingML signed 32-bit coordinate range.`);
  }
  return number;
}

function textRectangleCoordinate(value, label) {
  const number = Number(value);
  const emu = Math.round(number * EMU_PER_PIXEL);
  if (!Number.isFinite(number) || !Number.isSafeInteger(emu) || emu < -MAX_COORDINATE || emu > MAX_COORDINATE) {
    throw new RangeError(`${label} must be a finite pixel coordinate representable in the DrawingML signed 32-bit EMU range.`);
  }
  return number;
}

export function normalizePresentationCustomTextRectangle(value) {
  if (value == null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError("Presentation custom geometry textRectangle must be an object.");
  const unknown = Object.keys(value).filter((key) => !TEXT_RECTANGLE_FIELD_SET.has(key));
  if (unknown.length) throw new TypeError(`Presentation custom geometry textRectangle has unsupported fields: ${unknown.join(", ")}.`);
  const rectangle = Object.fromEntries(TEXT_RECTANGLE_FIELDS.map((field) => [field, textRectangleCoordinate(value[field], `Presentation custom geometry textRectangle.${field}`)]));
  if (Math.round(rectangle.left * EMU_PER_PIXEL) >= Math.round(rectangle.right * EMU_PER_PIXEL)) {
    throw new RangeError("Presentation custom geometry textRectangle.right must be greater than left at native EMU precision.");
  }
  if (Math.round(rectangle.top * EMU_PER_PIXEL) >= Math.round(rectangle.bottom * EMU_PER_PIXEL)) {
    throw new RangeError("Presentation custom geometry textRectangle.bottom must be greater than top at native EMU precision.");
  }
  return rectangle;
}

export function presentationCustomTextRectangleFrame(value, frame, sourceFrame = frame) {
  const rectangle = normalizePresentationCustomTextRectangle(value);
  if (!rectangle) return { ...frame };
  const sourceWidth = Number(sourceFrame?.width);
  const sourceHeight = Number(sourceFrame?.height);
  const left = Number(frame?.left);
  const top = Number(frame?.top);
  const width = Number(frame?.width);
  const height = Number(frame?.height);
  if (![sourceWidth, sourceHeight, left, top, width, height].every(Number.isFinite) || sourceWidth <= 0 || sourceHeight <= 0 || width <= 0 || height <= 0) {
    throw new RangeError("Presentation custom geometry textRectangle requires positive source and rendered shape frames.");
  }
  const scaleX = width / sourceWidth;
  const scaleY = height / sourceHeight;
  return {
    left: left + rectangle.left * scaleX,
    top: top + rectangle.top * scaleY,
    width: (rectangle.right - rectangle.left) * scaleX,
    height: (rectangle.bottom - rectangle.top) * scaleY,
  };
}

function angle(value, label) {
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number < -MAX_COORDINATE || number > MAX_COORDINATE) {
    throw new RangeError(`${label} must be a safe integer within the DrawingML signed 32-bit angle range.`);
  }
  return number;
}

function point(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const unknown = Object.keys(value).filter((key) => key !== "x" && key !== "y");
  if (unknown.length) throw new TypeError(`${label} has unsupported fields: ${unknown.join(", ")}.`);
  return { x: coordinate(value.x, `${label}.x`), y: coordinate(value.y, `${label}.y`) };
}

function curve(value, label, fields) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const allowed = new Set(fields);
  const unknown = Object.keys(value).filter((key) => !allowed.has(key));
  if (unknown.length) throw new TypeError(`${label} has unsupported fields: ${unknown.join(", ")}.`);
  return Object.fromEntries(fields.map((field) => [field, coordinate(value[field], `${label}.${field}`)]));
}

function arc(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const allowed = new Set(ARC_FIELDS);
  const unknown = Object.keys(value).filter((key) => !allowed.has(key));
  if (unknown.length) throw new TypeError(`${label} has unsupported fields: ${unknown.join(", ")}.`);
  const widthRadius = coordinate(value.widthRadius, `${label}.widthRadius`);
  const heightRadius = coordinate(value.heightRadius, `${label}.heightRadius`);
  if (widthRadius <= 0 || heightRadius <= 0) throw new RangeError(`${label} radii must be positive.`);
  const startAngle = angle(value.startAngle, `${label}.startAngle`);
  const sweepAngle = angle(value.sweepAngle, `${label}.sweepAngle`);
  if (sweepAngle === 0 || Math.abs(sweepAngle) > FULL_TURN_ANGLE) {
    throw new RangeError(`${label}.sweepAngle must be non-zero and no greater than one full DrawingML turn (${FULL_TURN_ANGLE}).`);
  }
  return { widthRadius, heightRadius, startAngle, sweepAngle };
}

function command(value, pathIndex, commandIndex) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`Presentation custom path ${pathIndex + 1} command ${commandIndex + 1} must be an object.`);
  const keys = Object.keys(value);
  if (keys.length !== 1) throw new TypeError(`Presentation custom path ${pathIndex + 1} command ${commandIndex + 1} must contain exactly one command.`);
  const label = `Presentation custom path ${pathIndex + 1} command ${commandIndex + 1}`;
  if (keys[0] === "moveTo" || keys[0] === "lineTo") return { [keys[0]]: point(value[keys[0]], `${label}.${keys[0]}`) };
  if (keys[0] === "arcTo") return { arcTo: arc(value.arcTo, `${label}.arcTo`) };
  const curveFields = Object.hasOwn(CURVE_FIELDS, keys[0]) ? CURVE_FIELDS[keys[0]] : undefined;
  if (curveFields) return { [keys[0]]: curve(value[keys[0]], `${label}.${keys[0]}`, curveFields) };
  if (keys[0] === "close") {
    if (value.close !== true && (typeof value.close !== "object" || value.close == null || Array.isArray(value.close) || Object.keys(value.close).length)) {
      throw new TypeError(`${label}.close must be true or an empty object.`);
    }
    return { close: {} };
  }
  throw new TypeError(`${label} uses unsupported command ${keys[0]}.`);
}

export function normalizePresentationCustomPaths(value) {
  if (value == null) return [];
  if (!Array.isArray(value) || value.length === 0 || value.length > MAX_PATHS) throw new RangeError(`Presentation custom geometry must contain 1 through ${MAX_PATHS} paths.`);
  let commandCount = 0;
  return value.map((path, pathIndex) => {
    if (!path || typeof path !== "object" || Array.isArray(path)) throw new TypeError(`Presentation custom path ${pathIndex + 1} must be an object.`);
    const unknown = Object.keys(path).filter((key) => !PATH_FIELDS.has(key));
    if (unknown.length) throw new TypeError(`Presentation custom path ${pathIndex + 1} has unsupported fields: ${unknown.join(", ")}.`);
    const width = coordinate(path.width, `Presentation custom path ${pathIndex + 1}.width`);
    const height = coordinate(path.height, `Presentation custom path ${pathIndex + 1}.height`);
    if (width <= 0 || height <= 0) throw new RangeError(`Presentation custom path ${pathIndex + 1} width and height must be positive.`);
    if (!Array.isArray(path.commands) || path.commands.length === 0) throw new TypeError(`Presentation custom path ${pathIndex + 1} requires commands.`);
    commandCount += path.commands.length;
    if (commandCount > MAX_COMMANDS) throw new RangeError(`Presentation custom geometry exceeds the ${MAX_COMMANDS}-command budget.`);
    let hasCurrentPoint = false;
    let hasSubpathStart = false;
    const commands = path.commands.map((item, commandIndex) => {
      const normalized = command(item, pathIndex, commandIndex);
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

export function presentationCustomPathsSvg(paths, frame, { escape = String } = {}) {
  return normalizePresentationCustomPaths(paths).map((path) => {
    const chunks = [];
    let currentPoint;
    let subpathStart;
    for (const item of path.commands) {
      if (item.moveTo) {
        chunks.push(`M ${item.moveTo.x} ${item.moveTo.y}`);
        currentPoint = { ...item.moveTo };
        subpathStart = { ...item.moveTo };
      } else if (item.lineTo) {
        chunks.push(`L ${item.lineTo.x} ${item.lineTo.y}`);
        currentPoint = { ...item.lineTo };
      } else if (item.quadraticBezTo) {
        chunks.push(`Q ${item.quadraticBezTo.x1} ${item.quadraticBezTo.y1} ${item.quadraticBezTo.x} ${item.quadraticBezTo.y}`);
        currentPoint = { x: item.quadraticBezTo.x, y: item.quadraticBezTo.y };
      } else if (item.cubicBezTo) {
        chunks.push(`C ${item.cubicBezTo.x1} ${item.cubicBezTo.y1} ${item.cubicBezTo.x2} ${item.cubicBezTo.y2} ${item.cubicBezTo.x} ${item.cubicBezTo.y}`);
        currentPoint = { x: item.cubicBezTo.x, y: item.cubicBezTo.y };
      } else if (item.arcTo) {
        const arc = svgArcCommands(item.arcTo, currentPoint);
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
