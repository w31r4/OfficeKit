import { createHash } from "node:crypto";

const MAX_DURATION_MS = 60_000;
const MAX_DELAY_MS = 60_000;
const MAX_STAGGER_MS = 10_000;
const EFFECTS = new Set(["fade", "wipe", "fly", "zoom", "pulse"]);
const PHASES = new Set(["entrance", "emphasis", "exit"]);
const STARTS = new Set(["withPrevious", "afterPrevious", "onClick"]);
const DIRECTIONS = new Set(["left", "right", "up", "down"]);
const TEXT_BUILDS = new Set(["whole", "paragraph"]);
const CHART_BUILDS = new Set(["allAtOnce", "series", "category", "seriesElement", "categoryElement"]);

export const PRESENTATION_ANIMATIONS_CAPABILITY = Symbol.for("office-kit.slide-animations-capability");
export const PRESENTATION_MORPH_CAPABILITY = Symbol.for("office-kit.slide-morph-capability");

function boundedInteger(value, name, maximum) {
  const number = value === undefined ? 0 : Number(value);
  if (!Number.isInteger(number) || number < 0 || number > maximum) {
    throw new RangeError(`Presentation animation ${name} must be an integer from 0 through ${maximum}.`);
  }
  return number;
}

function clone(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

function idHash(value) {
  return createHash("sha256").update(String(value), "utf8").digest("hex").slice(0, 16);
}

export function normalizePresentationAnimation(config = {}, owner = "slide") {
  if (!config || typeof config !== "object" || Array.isArray(config)) throw new TypeError("Presentation animation must be an object.");
  const unsupported = Object.keys(config).filter((key) => !new Set([
    "id", "target", "targetId", "effect", "phase", "start", "direction", "durationMs", "delayMs", "textBuild", "chartBuild", "staggerMs",
  ]).has(key));
  if (unsupported.length) throw new TypeError(`Presentation animation has unsupported fields: ${unsupported.join(", ")}.`);
  const target = config.targetId ?? (typeof config.target === "string" ? config.target : config.target?.id);
  if (!target) throw new TypeError(`Presentation ${owner} animation requires a target object or targetId.`);
  const effect = String(config.effect || "fade");
  if (!EFFECTS.has(effect)) throw new TypeError(`Presentation animation effect must be one of: ${[...EFFECTS].join(", ")}.`);
  const phase = String(config.phase || "entrance");
  if (!PHASES.has(phase)) throw new TypeError("Presentation animation phase must be entrance, emphasis, or exit.");
  const start = String(config.start || "afterPrevious");
  if (!STARTS.has(start)) throw new TypeError("Presentation animation start must be withPrevious, afterPrevious, or onClick.");
  const direction = config.direction === undefined ? undefined : String(config.direction);
  if (direction !== undefined && !DIRECTIONS.has(direction)) throw new TypeError("Presentation animation direction must be left, right, up, or down.");
  const textBuild = config.textBuild === undefined ? undefined : String(config.textBuild);
  if (textBuild !== undefined && !TEXT_BUILDS.has(textBuild)) throw new TypeError("Presentation animation textBuild must be whole or paragraph.");
  const chartBuild = config.chartBuild === undefined ? undefined : String(config.chartBuild);
  if (chartBuild !== undefined && !CHART_BUILDS.has(chartBuild)) throw new TypeError("Presentation animation chartBuild must be allAtOnce, series, category, seriesElement, or categoryElement.");
  if (textBuild && chartBuild) throw new TypeError("Presentation animation cannot specify both textBuild and chartBuild.");
  if ((effect === "wipe" || effect === "fly") && !direction) throw new TypeError(`Presentation ${effect} animation requires direction.`);
  if (effect !== "wipe" && effect !== "fly" && direction) throw new TypeError(`Presentation ${effect} animation does not accept direction.`);
  if (chartBuild && effect !== "wipe" && effect !== "fade") throw new TypeError("Chart build animations support fade or wipe effects only.");
  const durationMs = boundedInteger(config.durationMs ?? 500, "durationMs", MAX_DURATION_MS);
  if (durationMs === 0) throw new RangeError("Presentation animation durationMs must be greater than zero.");
  const delayMs = boundedInteger(config.delayMs, "delayMs", MAX_DELAY_MS);
  const staggerMs = boundedInteger(config.staggerMs, "staggerMs", MAX_STAGGER_MS);
  return {
    id: String(config.id || `anim-${idHash(`${target}:${effect}:${phase}:${start}:${durationMs}:${delayMs}:${staggerMs}`)}`),
    targetId: String(target),
    effect,
    phase,
    start,
    ...(direction ? { direction } : {}),
    durationMs,
    ...(delayMs ? { delayMs } : {}),
    ...(textBuild ? { textBuild } : {}),
    ...(chartBuild ? { chartBuild } : {}),
    ...(staggerMs ? { staggerMs } : {}),
  };
}

export function normalizePresentationMorph(config = {}) {
  if (!config || typeof config !== "object" || Array.isArray(config)) throw new TypeError("Presentation Morph must be an object.");
  const durationMs = boundedInteger(config.durationMs ?? 600, "Morph durationMs", MAX_DURATION_MS);
  if (durationMs === 0) throw new RangeError("Presentation Morph durationMs must be greater than zero.");
  if (!Array.isArray(config.pairs) || config.pairs.length === 0 || config.pairs.length > 256) throw new TypeError("Presentation Morph requires one through 256 pairs.");
  const pairs = config.pairs.map((pair, index) => {
    if (!pair || typeof pair !== "object") throw new TypeError(`Presentation Morph pair ${index + 1} must be an object.`);
    const key = String(pair.key || "");
    const fromId = String(pair.fromId || pair.from || "");
    const toId = String(pair.toId || pair.to || "");
    if (!key || !fromId || !toId) throw new TypeError(`Presentation Morph pair ${index + 1} requires key, fromId, and toId.`);
    return { key, fromId, toId };
  });
  const keys = new Set(pairs.map((pair) => pair.key));
  if (keys.size !== pairs.length) throw new TypeError("Presentation Morph pair keys must be unique.");
  return { durationMs, pairs };
}

export class SlideAnimations {
  constructor(slide, values = []) {
    this.slide = slide;
    this._items = values.map((value) => normalizePresentationAnimation(value));
  }

  get items() { return this._items.map(clone); }
  get count() { return this._items.length; }
  get capability() {
    return this[PRESENTATION_ANIMATIONS_CAPABILITY]
      ? { ...this[PRESENTATION_ANIMATIONS_CAPABILITY] }
      : { sourceBound: false, present: this.count > 0, editable: true, addable: true };
  }

  add(target, options = {}) {
    const resolved = typeof target === "string" ? this.slide.resolve(target) : target;
    if (!resolved || resolved.slide !== this.slide) throw new Error("Presentation animation target must belong to this slide.");
    const value = normalizePresentationAnimation({ ...options, target: resolved });
    if (this._items.some((item) => item.id === value.id)) throw new Error(`Presentation animation id ${value.id} already exists on this slide.`);
    if (this[PRESENTATION_ANIMATIONS_CAPABILITY]?.sourceBound && !this.capability.addable && this._items.length === 0) {
      throw new Error("Imported presentation timing is not safely extensible by this codec profile.");
    }
    this._items.push(value);
    return clone(value);
  }

  remove(animationOrId) {
    const id = typeof animationOrId === "string" ? animationOrId : animationOrId?.id;
    const index = this._items.findIndex((item) => item.id === id);
    if (index < 0) return false;
    if (this[PRESENTATION_ANIMATIONS_CAPABILITY]?.sourceBound && !this.capability.editable) throw new Error("Imported presentation timing is opaque and cannot be edited.");
    this._items.splice(index, 1);
    return true;
  }

  clear() {
    if (this[PRESENTATION_ANIMATIONS_CAPABILITY]?.sourceBound && !this.capability.editable && this._items.length) throw new Error("Imported presentation timing is opaque and cannot be edited.");
    this._items = [];
    return this.slide;
  }

  getItem(id) { const found = this._items.find((item) => item.id === id); return clone(found); }
  toJSON() { return this.items; }
  inspectRecord() { return { kind: "animations", id: `${this.slide.id}/animations`, slide: this.slide.index + 1, items: this.items, capability: this.capability }; }
}

export class SlideMorph {
  constructor(slide, value) {
    this.slide = slide;
    this._value = value == null ? undefined : normalizePresentationMorph(value);
  }
  get id() { return `${this.slide.id}/morph`; }
  get configured() { return Boolean(this._value); }
  get capability() { return this[PRESENTATION_MORPH_CAPABILITY] ? { ...this[PRESENTATION_MORPH_CAPABILITY] } : { sourceBound: false, editable: true, addable: true }; }
  get value() { return clone(this._value); }
  set(value) {
    if (this[PRESENTATION_MORPH_CAPABILITY]?.sourceBound && !this.capability.editable && this._value) throw new Error("Imported Morph extension is opaque and cannot be edited.");
    this._value = value == null ? undefined : normalizePresentationMorph(value);
    return this.slide;
  }
  clear() { return this.set(undefined); }
  toJSON() { return this.value; }
  inspectRecord() { return { kind: "morph", id: this.id, slide: this.slide.index + 1, value: this.value, capability: this.capability }; }
}
