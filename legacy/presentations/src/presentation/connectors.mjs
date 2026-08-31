import { aid } from "../shared/ids.mjs";
import {
  normalizePresentationLineEnd,
  normalizePresentationLineStyle,
  presentationLineSvgStyle,
} from "./line-styles.mjs";
import {
  normalizePresentationCustomConnectionSites,
  presentationCustomConnectionSitePoint,
} from "./custom-geometry.mjs";
import {
  initializePresentationAccessibility,
  presentationAccessibilityCapability,
  setPresentationAccessibilityMetadata,
} from "./accessibility.mjs";
import { deletePresentationElement, presentationElementDeletionCapability } from "./element-deletion.mjs";

const CONNECTOR_TYPE_ALIASES = new Map([
  ["straight", "straight"],
  ["elbow", "elbow"],
  ["elbow2", "elbow"],
  ["elbow3", "elbow"],
  ["elbow4", "elbow"],
  ["elbow5", "elbow"],
  ["curved", "curved"],
]);
const EMU_PER_PIXEL = 9_525;

const CARDINAL_SITE_INDEXES = new Map([
  ["rect", Object.freeze({ top: 0, left: 1, bottom: 2, right: 3 })],
  ["roundRect", Object.freeze({ top: 0, left: 1, bottom: 2, right: 3 })],
  ["textbox", Object.freeze({ top: 0, left: 1, bottom: 2, right: 3 })],
  ["ellipse", Object.freeze({ top: 0, left: 2, bottom: 4, right: 6 })],
]);

function finitePoint(value, name) {
  if (!value || !Number.isFinite(Number(value.x)) || !Number.isFinite(Number(value.y))) {
    throw new TypeError(`${name} must define finite x and y coordinates.`);
  }
  return { x: Number(value.x), y: Number(value.y) };
}

function shapeOwner(shape) {
  return shape?.parentGroup;
}

function presentationShapeTarget(slide, owner, target, name) {
  const id = typeof target === "string" ? target : target?.id;
  if (!id) throw new TypeError(`${name} must be a presentation shape or non-empty shape id.`);
  const resolved = slide.resolve(id);
  if (!resolved || resolved.kind === "connector" || !resolved.position || typeof resolved.geometry !== "string" || resolved.text == null) {
    throw new Error(`${name} ${id} is not a presentation shape on this slide.`);
  }
  if (shapeOwner(resolved) !== owner) {
    throw new Error(`${name} ${id} must belong to the same slide or group shape tree as the connector.`);
  }
  return resolved;
}

function normalizedSide(value, name) {
  const side = String(value || "");
  if (!new Set(["top", "left", "bottom", "right"]).has(side)) {
    throw new RangeError(`${name} must be top, left, bottom, or right.`);
  }
  return side;
}

export function normalizePresentationConnectorType(value, name = "Presentation connector type") {
  const requested = String(value || "straight");
  const normalized = CONNECTOR_TYPE_ALIASES.get(requested);
  if (!normalized) throw new RangeError(`${name} ${requested} is unsupported.`);
  return normalized;
}

export function normalizePresentationConnectionSiteIndex(value, name = "Presentation connection-site index") {
  const index = Number(value);
  if (!Number.isInteger(index) || index < 0 || index > 0xffff_ffff) {
    throw new RangeError(`${name} must be an unsigned 32-bit integer.`);
  }
  return index;
}

function siteIndexes(shape) {
  return CARDINAL_SITE_INDEXES.get(shape.geometry);
}

function customSiteCount(shape) {
  if (shape.geometry !== "custom") return undefined;
  return normalizePresentationCustomConnectionSites(shape.customConnectionSites, {
    adjustments: shape.customAdjustments,
    guides: shape.customGuides,
    widthEmu: Math.round(Number(shape.position?.width) * EMU_PER_PIXEL),
    heightEmu: Math.round(Number(shape.position?.height) * EMU_PER_PIXEL),
  }).length;
}

export function presentationConnectionSiteIndex(slide, owner, target, side) {
  const shape = presentationShapeTarget(slide, owner, target, "Presentation connector target");
  const indexes = siteIndexes(shape);
  if (!indexes) {
    if (shape.geometry === "custom" && customSiteCount(shape) > 0) {
      throw new RangeError(`Presentation custom shape ${shape.id} requires an explicit connection-site index; cardinal side aliases are unavailable.`);
    }
    throw new RangeError(`Presentation shape ${shape.id} geometry ${shape.geometry} has no modeled connection-site map.`);
  }
  return indexes[normalizedSide(side, "Presentation connector side")];
}

function validateSiteIndex(shape, index, name) {
  const normalized = normalizePresentationConnectionSiteIndex(index, name);
  const customCount = customSiteCount(shape);
  if (customCount !== undefined) {
    if (customCount === 0) throw new RangeError(`Presentation custom shape ${shape.id} has no modeled connection sites.`);
    if (normalized >= customCount) throw new RangeError(`${name} ${normalized} is outside the modeled custom connection-site range 0..${customCount - 1}.`);
    return normalized;
  }
  const indexes = siteIndexes(shape);
  if (!indexes) {
    throw new RangeError(`Presentation shape ${shape.id} geometry ${shape.geometry} has no modeled connection-site map.`);
  }
  const maximum = shape.geometry === "ellipse" ? 7 : 3;
  if (normalized > maximum) {
    throw new RangeError(`${name} ${normalized} is outside the modeled ${shape.geometry} connection-site range 0..${maximum}.`);
  }
  return normalized;
}

function rotateShapePoint(shape, point) {
  const frame = shape.position;
  const centerX = Number(frame.left) + Number(frame.width) / 2;
  const centerY = Number(frame.top) + Number(frame.height) / 2;
  let x = point.x;
  let y = point.y;
  if (shape.transform?.flipHorizontal === true) x = centerX * 2 - x;
  if (shape.transform?.flipVertical === true) y = centerY * 2 - y;
  const rotation = Number(shape.transform?.rotationDegrees || 0);
  if (!rotation) return { x, y };
  const radians = rotation * Math.PI / 180;
  const dx = x - centerX;
  const dy = y - centerY;
  return {
    x: centerX + dx * Math.cos(radians) - dy * Math.sin(radians),
    y: centerY + dx * Math.sin(radians) + dy * Math.cos(radians),
  };
}

export function presentationConnectionSitePoint(shape, index, name = "Presentation connection-site index") {
  const site = validateSiteIndex(shape, index, name);
  const frame = shape.position;
  const left = Number(frame.left);
  const top = Number(frame.top);
  const width = Number(frame.width);
  const height = Number(frame.height);
  if (![left, top, width, height].every(Number.isFinite) || width < 0 || height < 0) {
    throw new RangeError(`Presentation shape ${shape.id} has an invalid frame for connector routing.`);
  }
  let point;
  if (shape.geometry === "custom") {
    point = presentationCustomConnectionSitePoint(shape.customConnectionSites, site, frame, {
      adjustments: shape.customAdjustments,
      guides: shape.customGuides,
    });
  } else if (shape.geometry === "ellipse") {
    const angles = [-90, -135, 180, 135, 90, 45, 0, -45];
    const radians = angles[site] * Math.PI / 180;
    point = {
      x: left + width / 2 + Math.cos(radians) * width / 2,
      y: top + height / 2 + Math.sin(radians) * height / 2,
    };
  } else {
    point = [
      { x: left + width / 2, y: top },
      { x: left, y: top + height / 2 },
      { x: left + width / 2, y: top + height },
      { x: left + width, y: top + height / 2 },
    ][site];
  }
  return rotateShapePoint(shape, point);
}

function shapeFingerprint(shape) {
  if (!shape) return undefined;
  return JSON.stringify({
    id: shape.id,
    geometry: shape.geometry,
    position: shape.position,
    transform: shape.transform,
    ...(shape.geometry === "custom" ? {
      customAdjustments: shape.customAdjustments,
      customGuides: shape.customGuides,
      customConnectionSites: shape.customConnectionSites,
    } : {}),
  });
}

function endpointFingerprint(shape, index, explicitSite) {
  const fingerprint = shapeFingerprint(shape);
  return fingerprint == null ? undefined : `${fingerprint}|${index}|${explicitSite ? "explicit" : "implicit"}`;
}

function legacyTargetPoint(target, index, name, explicitSite = true) {
  const frame = target?.position || target?.frame;
  if (!explicitSite && frame && [frame.left, frame.top, frame.width, frame.height].every((value) => Number.isFinite(Number(value)))) {
    return { x: Number(frame.left) + Number(frame.width) / 2, y: Number(frame.top) + Number(frame.height) / 2 };
  }
  if (target?.geometry && target?.text != null) return presentationConnectionSitePoint(target, index, name);
  throw new RangeError(`${name} cannot route against an unmodeled target connection-site table.`);
}

function automaticSides(from, to) {
  const fromX = Number(from.position.left) + Number(from.position.width) / 2;
  const fromY = Number(from.position.top) + Number(from.position.height) / 2;
  const toX = Number(to.position.left) + Number(to.position.width) / 2;
  const toY = Number(to.position.top) + Number(to.position.height) / 2;
  if (Math.abs(toX - fromX) >= Math.abs(toY - fromY)) {
    return toX >= fromX ? { from: "right", to: "left" } : { from: "left", to: "right" };
  }
  return toY >= fromY ? { from: "bottom", to: "top" } : { from: "top", to: "bottom" };
}

function requestedSite(shape, options, indexKey, sideKey, automaticSide, name) {
  if (Object.hasOwn(options, indexKey) && options[indexKey] != null) {
    if (Object.hasOwn(options, sideKey) && options[sideKey] != null) {
      throw new TypeError(`${name} cannot define both ${indexKey} and ${sideKey}.`);
    }
    return validateSiteIndex(shape, options[indexKey], `${name}.${indexKey}`);
  }
  const side = options[sideKey] == null ? automaticSide : normalizedSide(options[sideKey], `${name}.${sideKey}`);
  return siteIndexes(shape)?.[side] ?? presentationConnectionSiteIndex(shape.slide, shapeOwner(shape), shape, side);
}

export function connectedPresentationShapeConfig(slide, owner, from, to, options = {}, { requireExplicitSites = false } = {}) {
  if (!options || typeof options !== "object" || Array.isArray(options)) throw new TypeError("Presentation connector options must be an object.");
  const fromShape = presentationShapeTarget(slide, owner, from, "Presentation connector from endpoint");
  const toShape = presentationShapeTarget(slide, owner, to, "Presentation connector to endpoint");
  if (requireExplicitSites && (!Object.hasOwn(options, "fromIdx") || !Object.hasOwn(options, "toIdx"))) {
    throw new TypeError("Direct presentation connector creation requires fromIdx and toIdx.");
  }
  const automatic = automaticSides(fromShape, toShape);
  const fromIdx = requestedSite(fromShape, options, "fromIdx", "fromSide", automatic.from, "Presentation connector");
  const toIdx = requestedSite(toShape, options, "toIdx", "toSide", automatic.to, "Presentation connector");
  return {
    ...options,
    connectorType: normalizePresentationConnectorType(options.kind || options.connectorType || (requireExplicitSites ? "straight" : "elbow")),
    from: fromShape,
    to: toShape,
    fromIdx,
    toIdx,
    start: presentationConnectionSitePoint(fromShape, fromIdx, "Presentation connector.fromIdx"),
    end: presentationConnectionSitePoint(toShape, toIdx, "Presentation connector.toIdx"),
    line: options.line || { style: "solid", fill: "#334155", width: 2 },
    ...(Object.hasOwn(options, "head") ? { head: options.head } : {}),
    ...(Object.hasOwn(options, "tail") ? { tail: options.tail } : {}),
    cap: options.cap,
    join: options.join,
    _zPlacement: "back",
  };
}

function normalizeConnectorStyle(config) {
  const source = config.line == null ? {} : config.line;
  if (typeof source !== "object" || Array.isArray(source)) throw new TypeError("Presentation connector line must be an object.");
  const requested = {
    ...source,
    ...(config.head == null ? (source.head == null && source.startArrow == null && config.startArrow != null ? { head: config.startArrow } : {}) : { head: config.head }),
    ...(config.tail == null ? (source.tail == null && source.endArrow == null && config.endArrow != null ? { tail: config.endArrow } : {}) : { tail: config.tail }),
    ...(config.cap == null ? {} : { cap: config.cap }),
    ...(config.join == null ? {} : { join: config.join }),
  };
  const normalized = normalizePresentationLineStyle(requested, {
    name: "Presentation connector line",
    defaultWidth: 2,
  });
  const { head, tail, cap, join, ...paint } = normalized;
  return {
    line: {
      ...paint,
      startArrow: head?.type,
      startArrowWidth: head?.width,
      startArrowLength: head?.length,
      endArrow: tail?.type,
      endArrowWidth: tail?.width,
      endArrowLength: tail?.length,
    },
    head,
    tail,
    cap,
    join,
  };
}

export class PresentationConnectorElement {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.kind = "connector";
    this.id = config.id || aid("cx");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
    this.name = config.name || "";
    this.accessibility = initializePresentationAccessibility(this, config, `Presentation connector ${this.id}`);
    this.connectorType = normalizePresentationConnectorType(config.connectorType || config.kind || config.type || "straight");
    this.startTargetId = typeof config.from === "string" ? config.from : config.from?.id || config.startTargetId;
    this.endTargetId = typeof config.to === "string" ? config.to : config.to?.id || config.endTargetId;
    this._startSiteExplicit = Object.hasOwn(config, "fromIdx") || Object.hasOwn(config, "startSiteIndex");
    this._endSiteExplicit = Object.hasOwn(config, "toIdx") || Object.hasOwn(config, "endSiteIndex");
    this.startSiteIndex = normalizePresentationConnectionSiteIndex(config.fromIdx ?? config.startSiteIndex ?? 0, "Presentation connector start site index");
    this.endSiteIndex = normalizePresentationConnectionSiteIndex(config.toIdx ?? config.endSiteIndex ?? 0, "Presentation connector end site index");
    const startTarget = this.startTargetId ? slide.resolve(this.startTargetId) : undefined;
    const endTarget = this.endTargetId ? slide.resolve(this.endTargetId) : undefined;
    this._start = finitePoint(config.start || (startTarget ? legacyTargetPoint(startTarget, this.startSiteIndex, "Presentation connector start", this._startSiteExplicit) : { x: 0, y: 0 }), "Presentation connector start");
    this._end = finitePoint(config.end || (endTarget ? legacyTargetPoint(endTarget, this.endSiteIndex, "Presentation connector end", this._endSiteExplicit) : { x: 160, y: 0 }), "Presentation connector end");
    const style = normalizeConnectorStyle(config.line == null && config.head == null && config.tail == null
      ? { ...config, line: { fill: "#334155", width: 2, endArrow: config.endArrow || "triangle" } }
      : config);
    this.line = style.line;
    this.cap = style.cap;
    this.join = style.join;
    this._sourceBound = config._officeKitSourceBound === true;
    this._zPlacement = config._zPlacement === "front" ? "front" : "back";
    this._startFingerprint = this.startTargetId ? endpointFingerprint(slide.resolve(this.startTargetId), this.startSiteIndex, this._startSiteExplicit) : undefined;
    this._endFingerprint = this.endTargetId ? endpointFingerprint(slide.resolve(this.endTargetId), this.endSiteIndex, this._endSiteExplicit) : undefined;
  }

  #attachedPoint(side, strict) {
    const targetId = side === "start" ? this.startTargetId : this.endTargetId;
    const stored = side === "start" ? this._start : this._end;
    if (!targetId) return stored;
    const target = this.slide.resolve(targetId);
    if (!target) {
      if (strict) throw new Error(`Presentation connector ${this.id} has an unresolved ${side} target ${targetId}.`);
      return stored;
    }
    const fingerprintKey = side === "start" ? "_startFingerprint" : "_endFingerprint";
    const index = side === "start" ? this.startSiteIndex : this.endSiteIndex;
    const explicit = side === "start" ? this._startSiteExplicit : this._endSiteExplicit;
    const fingerprint = endpointFingerprint(target, index, explicit);
    if (this[fingerprintKey] === fingerprint) return stored;
    try {
      if (shapeOwner(target) !== this.parentGroup) {
        throw new Error(`Presentation connector ${this.id} ${side} target ${targetId} moved outside its slide or group shape tree.`);
      }
      const point = legacyTargetPoint(target, index, `Presentation connector ${this.id} ${side} site index`, explicit);
      if (side === "start") this._start = point;
      else this._end = point;
      this[fingerprintKey] = fingerprint;
      return point;
    } catch (error) {
      if (strict) throw error;
      return stored;
    }
  }

  get start() { return { ...this.#attachedPoint("start", false) }; }
  set start(value) {
    if (this.startTargetId && this._startSiteExplicit) throw new Error("A site-bound presentation connector start must be changed with setConnectorFrom().");
    this._start = finitePoint(value, "Presentation connector start");
  }
  get end() { return { ...this.#attachedPoint("end", false) }; }
  set end(value) {
    if (this.endTargetId && this._endSiteExplicit) throw new Error("A site-bound presentation connector end must be changed with setConnectorTo().");
    this._end = finitePoint(value, "Presentation connector end");
  }

  resolvedEndpoints(options = {}) {
    return {
      start: { ...this.#attachedPoint("start", options.strict === true) },
      end: { ...this.#attachedPoint("end", options.strict === true) },
    };
  }

  captureAttachedEndpointState() {
    this._startFingerprint = this.startTargetId ? endpointFingerprint(this.slide.resolve(this.startTargetId), this.startSiteIndex, this._startSiteExplicit) : undefined;
    this._endFingerprint = this.endTargetId ? endpointFingerprint(this.slide.resolve(this.endTargetId), this.endSiteIndex, this._endSiteExplicit) : undefined;
    return this;
  }

  get connector() {
    return {
      fromElementId: this.startTargetId,
      fromIdx: this.startSiteIndex,
      toElementId: this.endTargetId,
      toIdx: this.endSiteIndex,
    };
  }
  get connectorLineStyle() { return { head: this.connectorHead, tail: this.connectorTail, cap: this.cap, join: this.join }; }
  get head() {
    if (!this.line?.startArrow) return undefined;
    return { type: this.line.startArrow, ...(this.line.startArrowWidth ? { width: this.line.startArrowWidth } : {}), ...(this.line.startArrowLength ? { length: this.line.startArrowLength } : {}) };
  }
  set head(value) {
    const end = normalizePresentationLineEnd(value, "Presentation connector head");
    for (const key of ["startArrow", "startArrowWidth", "startArrowLength"]) delete this.line[key];
    if (end) Object.assign(this.line, { startArrow: end.type, ...(end.width ? { startArrowWidth: end.width } : {}), ...(end.length ? { startArrowLength: end.length } : {}) });
  }
  get tail() {
    if (!this.line?.endArrow) return undefined;
    return { type: this.line.endArrow, ...(this.line.endArrowWidth ? { width: this.line.endArrowWidth } : {}), ...(this.line.endArrowLength ? { length: this.line.endArrowLength } : {}) };
  }
  set tail(value) {
    const end = normalizePresentationLineEnd(value, "Presentation connector tail");
    for (const key of ["endArrow", "endArrowWidth", "endArrowLength"]) delete this.line[key];
    if (end) Object.assign(this.line, { endArrow: end.type, ...(end.width ? { endArrowWidth: end.width } : {}), ...(end.length ? { endArrowLength: end.length } : {}) });
  }
  get connectorHead() { return this.head ? { ...this.head } : undefined; }
  get connectorTail() { return this.tail ? { ...this.tail } : undefined; }
  get isForeground() { return this._zPlacement === "front"; }
  get accessibilityCapability() { return presentationAccessibilityCapability(this); }
  get deletionCapability() { return presentationElementDeletionCapability(this, "connector"); }

  delete() {
    const owner = this.parentGroup;
    const collection = owner?.connectors || this.slide?.connectors;
    return deletePresentationElement(this, collection, "connector");
  }

  setAccessibilityMetadata(update) {
    this.accessibility = setPresentationAccessibilityMetadata(this, this.accessibility, update, `Presentation connector ${this.id}`);
    return this;
  }

  #setEndpoint(side, target, index) {
    const shape = presentationShapeTarget(this.slide, this.parentGroup, target, `Presentation connector ${side} endpoint`);
    const siteIndex = validateSiteIndex(shape, index, `Presentation connector ${side} site index`);
    const point = presentationConnectionSitePoint(shape, siteIndex);
    if (side === "from") {
      this.startTargetId = shape.id;
      this.startSiteIndex = siteIndex;
      this._startSiteExplicit = true;
      this._start = point;
      this._startFingerprint = endpointFingerprint(shape, siteIndex, true);
    } else {
      this.endTargetId = shape.id;
      this.endSiteIndex = siteIndex;
      this._endSiteExplicit = true;
      this._end = point;
      this._endFingerprint = endpointFingerprint(shape, siteIndex, true);
    }
    return this;
  }

  setConnectorFrom(target, index) { return this.#setEndpoint("from", target, index); }
  setConnectorTo(target, index) { return this.#setEndpoint("to", target, index); }

  bringToFront() {
    if (this._sourceBound) throw new Error("Imported presentation connector z-order is source-bound and cannot be changed by this profile.");
    if (this.parentGroup) {
      const index = this.parentGroup.children.indexOf(this);
      if (index >= 0) this.parentGroup.children.splice(index, 1);
      this.parentGroup.children.push(this);
    } else {
      const index = this.slide.connectors.items.indexOf(this);
      if (index >= 0) this.slide.connectors.items.splice(index, 1);
      this.slide.connectors.items.push(this);
    }
    this._zPlacement = "front";
    return this;
  }

  sendToBack() {
    if (this._sourceBound) throw new Error("Imported presentation connector z-order is source-bound and cannot be changed by this profile.");
    if (this.parentGroup) {
      const index = this.parentGroup.children.indexOf(this);
      if (index >= 0) this.parentGroup.children.splice(index, 1);
      this.parentGroup.children.unshift(this);
    } else {
      const index = this.slide.connectors.items.indexOf(this);
      if (index >= 0) this.slide.connectors.items.splice(index, 1);
      this.slide.connectors.items.unshift(this);
    }
    this._zPlacement = "back";
    return this;
  }

  get position() {
    const { start, end } = this.resolvedEndpoints();
    const left = Math.min(start.x, end.x);
    const top = Math.min(start.y, end.y);
    return { left, top, width: Math.abs(end.x - start.x), height: Math.abs(end.y - start.y) };
  }

  inspectRecord() {
    const { start, end } = this.resolvedEndpoints();
    return { kind: "connector", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, connectorType: this.connectorType, start, end, startTargetId: this.startTargetId, endTargetId: this.endTargetId, startSiteIndex: this.startSiteIndex, endSiteIndex: this.endSiteIndex, line: this.line, head: this.connectorHead, tail: this.connectorTail, cap: this.cap, join: this.join, accessibility: this.accessibility ? { ...this.accessibility } : undefined, accessibilityCapability: this.accessibilityCapability, deletionCapability: this.deletionCapability };
  }

  layoutJson() {
    const { start, end } = this.resolvedEndpoints();
    return { kind: "connector", id: this.id, name: this.name, connectorType: this.connectorType, start, end, startTargetId: this.startTargetId, endTargetId: this.endTargetId, startSiteIndex: this.startSiteIndex, endSiteIndex: this.endSiteIndex, line: this.line, head: this.connectorHead, tail: this.connectorTail, cap: this.cap, join: this.join, accessibility: this.accessibility ? { ...this.accessibility } : undefined, frame: this.position };
  }

  toSvg() {
    const { start, end } = this.resolvedEndpoints();
    const style = presentationLineSvgStyle({
      ...this.line,
      ...(this.head ? { head: this.head } : {}),
      ...(this.tail ? { tail: this.tail } : {}),
      ...(this.cap ? { cap: this.cap } : {}),
      ...(this.join ? { join: this.join } : {}),
    }, {
      name: `Presentation connector ${this.name || this.id} line`,
      defaultWidth: 2,
      markerBase: this.id,
    });
    const paint = `fill="none" ${style.attributes}`;
    let line;
    if (this.connectorType === "elbow") {
      const middleX = (start.x + end.x) / 2;
      line = `<path d="M ${start.x} ${start.y} H ${middleX} V ${end.y} H ${end.x}" ${paint}/>`;
    } else if (this.connectorType === "curved") {
      const middleX = (start.x + end.x) / 2;
      line = `<path d="M ${start.x} ${start.y} C ${middleX} ${start.y}, ${middleX} ${end.y}, ${end.x} ${end.y}" ${paint}/>`;
    } else {
      line = `<line x1="${start.x}" y1="${start.y}" x2="${end.x}" y2="${end.y}" ${paint}/>`;
    }
    return `${style.definitions ? `<defs>${style.definitions}</defs>` : ""}${line}`;
  }
}
