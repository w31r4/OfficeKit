const CARDINAL_DIRECTIONS = Object.freeze(["left", "up", "right", "down"]);
const CORNER_DIRECTIONS = Object.freeze(["leftUp", "rightUp", "leftDown", "rightDown"]);
const EIGHT_DIRECTIONS = Object.freeze([...CARDINAL_DIRECTIONS, ...CORNER_DIRECTIONS]);
const IN_OUT_DIRECTIONS = Object.freeze(["in", "out"]);
const ORIENTATIONS = Object.freeze(["horizontal", "vertical"]);

// Complete ECMA-376 base-namespace p:transition effect vocabulary. Office
// version extensions remain opaque. Keeping each native shape in data makes
// validation and defaults local instead of growing one branch per effect.
const TRANSITION_EFFECTS = Object.freeze({
  blinds: Object.freeze({ orientations: ORIENTATIONS, defaultOrientation: "horizontal" }),
  checker: Object.freeze({ orientations: ORIENTATIONS, defaultOrientation: "horizontal" }),
  circle: Object.freeze({}),
  comb: Object.freeze({ orientations: ORIENTATIONS, defaultOrientation: "horizontal" }),
  cover: Object.freeze({ directions: EIGHT_DIRECTIONS, defaultDirection: "left" }),
  cut: Object.freeze({ throughBlack: true }),
  diamond: Object.freeze({}),
  dissolve: Object.freeze({}),
  fade: Object.freeze({ throughBlack: true }),
  newsflash: Object.freeze({}),
  plus: Object.freeze({}),
  pull: Object.freeze({ directions: EIGHT_DIRECTIONS, defaultDirection: "left" }),
  push: Object.freeze({ directions: CARDINAL_DIRECTIONS, defaultDirection: "left" }),
  random: Object.freeze({}),
  randomBar: Object.freeze({ orientations: ORIENTATIONS, defaultOrientation: "horizontal" }),
  split: Object.freeze({ orientations: ORIENTATIONS, defaultOrientation: "vertical", directions: IN_OUT_DIRECTIONS, defaultDirection: "out" }),
  strips: Object.freeze({ directions: CORNER_DIRECTIONS, defaultDirection: "rightDown" }),
  wedge: Object.freeze({}),
  wheel: Object.freeze({ spokes: true, defaultSpokes: 1 }),
  wipe: Object.freeze({ directions: CARDINAL_DIRECTIONS, defaultDirection: "left" }),
  zoom: Object.freeze({ directions: IN_OUT_DIRECTIONS, defaultDirection: "in" }),
});
const TRANSITION_EFFECT_NAMES = new Map(Object.keys(TRANSITION_EFFECTS).map((name) => [name.toLowerCase(), name]));
const TRANSITION_DIRECTION_NAMES = new Map(EIGHT_DIRECTIONS.concat(IN_OUT_DIRECTIONS).map((name) => [name.toLowerCase(), name]));
const TRANSITION_ORIENTATION_NAMES = new Map(ORIENTATIONS.map((name) => [name, name]));
const TRANSITION_SPEEDS = new Set(["slow", "medium", "fast"]);
const TRANSITION_KEYS = new Set(["effect", "direction", "orientation", "throughBlack", "spokes", "speed", "advanceOnClick", "advanceAfterMs"]);
const MAX_ADVANCE_AFTER_MS = 86_400_000;

export const PRESENTATION_TRANSITION_CAPABILITY = Symbol.for("office-kit.slide-transition-capability");

function own(object, key) {
  return Object.prototype.hasOwnProperty.call(object, key);
}

function normalizeEffect(value) {
  const token = String(value || "").trim().toLowerCase();
  const effect = TRANSITION_EFFECT_NAMES.get(token);
  if (!effect) throw new TypeError(`Presentation transition effect must be one of: ${Object.keys(TRANSITION_EFFECTS).join(", ")}.`);
  return effect;
}

function normalizeSpeed(value) {
  const speed = String(value ?? "medium").trim().toLowerCase();
  if (!TRANSITION_SPEEDS.has(speed)) {
    throw new TypeError("Presentation transition speed must be slow, medium, or fast.");
  }
  return speed;
}

function normalizeAdvanceAfter(value) {
  const milliseconds = Number(value);
  if (!Number.isSafeInteger(milliseconds) || milliseconds < 0 || milliseconds > MAX_ADVANCE_AFTER_MS) {
    throw new RangeError(`Presentation transition advanceAfterMs must be an integer from 0 through ${MAX_ADVANCE_AFTER_MS}.`);
  }
  return milliseconds;
}

function choiceList(values) {
  if (values.length === 1) return values[0];
  if (values.length === 2) return `${values[0]} or ${values[1]}`;
  return `${values.slice(0, -1).join(", ")}, or ${values.at(-1)}`;
}

function normalizeProfileToken(value, allowed, names, label) {
  const token = String(value ?? "").trim().replace(/[-_\s]/g, "").toLowerCase();
  const canonical = names.get(token);
  if (!canonical || !allowed.includes(canonical)) throw new TypeError(`${label} must be ${choiceList(allowed)}.`);
  return canonical;
}

function rejectUnused(config, key) {
  if (own(config, key) && config[key] != null) {
    throw new TypeError(`Presentation ${config.effect} transition does not accept ${key}.`);
  }
}

export function normalizePresentationTransition(config) {
  if (!config || typeof config !== "object" || Array.isArray(config)) {
    throw new TypeError("Presentation transition must be an object.");
  }
  const unsupported = Object.keys(config).filter((key) => !TRANSITION_KEYS.has(key));
  if (unsupported.length) {
    throw new TypeError(`Presentation transition has unsupported fields: ${unsupported.join(", ")}.`);
  }
  const effect = normalizeEffect(config.effect);
  const profile = TRANSITION_EFFECTS[effect];
  const speed = normalizeSpeed(config.speed);
  const transition = { effect, speed };
  if (profile.directions) {
    const direction = String(config.direction ?? profile.defaultDirection).trim().toLowerCase();
    transition.direction = normalizeProfileToken(direction, profile.directions, TRANSITION_DIRECTION_NAMES, `Presentation ${effect} transition direction`);
  } else if (own(config, "direction") && config.direction != null) {
    throw new TypeError(`Presentation ${effect} transition does not accept direction.`);
  }
  if (profile.orientations) {
    transition.orientation = normalizeProfileToken(
      config.orientation ?? profile.defaultOrientation,
      profile.orientations,
      TRANSITION_ORIENTATION_NAMES,
      `Presentation ${effect} transition orientation`,
    );
  } else rejectUnused(config, "orientation");
  if (profile.throughBlack) {
    if (own(config, "throughBlack") && config.throughBlack != null) {
      if (typeof config.throughBlack !== "boolean") throw new TypeError(`Presentation ${effect} transition throughBlack must be a boolean.`);
      transition.throughBlack = config.throughBlack;
    }
  } else rejectUnused(config, "throughBlack");
  if (profile.spokes) {
    const spokes = Number(config.spokes ?? profile.defaultSpokes);
    if (!Number.isSafeInteger(spokes) || spokes < 1 || spokes > 8) {
      throw new RangeError("Presentation wheel transition spokes must be an integer from 1 through 8.");
    }
    transition.spokes = spokes;
  } else rejectUnused(config, "spokes");
  if (own(config, "advanceOnClick") && typeof config.advanceOnClick !== "boolean") {
    throw new TypeError("Presentation transition advanceOnClick must be a boolean.");
  }
  transition.advanceOnClick = config.advanceOnClick ?? true;
  if (own(config, "advanceAfterMs") && config.advanceAfterMs != null) {
    transition.advanceAfterMs = normalizeAdvanceAfter(config.advanceAfterMs);
  }
  return transition;
}

function cloneTransition(value) {
  return value ? { ...value } : undefined;
}

export class SlideTransition {
  constructor(slide, config) {
    this.slide = slide;
    this._value = config == null ? undefined : normalizePresentationTransition(config);
  }

  get id() { return `${this.slide.id}/transition`; }
  get configured() { return Boolean(this._value); }
  get effect() { return this._value?.effect; }
  get direction() { return this._value?.direction; }
  get orientation() { return this._value?.orientation; }
  get throughBlack() { return this._value?.throughBlack; }
  get spokes() { return this._value?.spokes; }
  get speed() { return this._value?.speed; }
  get advanceOnClick() { return this._value?.advanceOnClick; }
  get advanceAfterMs() { return this._value?.advanceAfterMs; }
  get capability() {
    const imported = this[PRESENTATION_TRANSITION_CAPABILITY];
    return imported
      ? { ...imported }
      : { sourceBound: false, partPresent: this.configured, editable: true, addable: true };
  }

  set(config) {
    const capability = this.capability;
    if (capability.sourceBound && !capability.editable && !capability.addable) {
      throw new Error("Presentation slide transition is source-bound and cannot be semantically set by this codec profile.");
    }
    this._value = normalizePresentationTransition(config);
    return this;
  }

  clear() {
    const capability = this.capability;
    if (capability.sourceBound && capability.partPresent && !capability.editable) {
      throw new Error("Presentation slide transition is source-bound and cannot be removed by this codec profile.");
    }
    this._value = undefined;
    return this;
  }

  inspectRecord() {
    return {
      kind: "transition",
      id: this.id,
      slide: this.slide.index + 1,
      configured: this.configured,
      ...(this._value || {}),
      capability: this.capability,
    };
  }

  toJSON() { return cloneTransition(this._value); }

  // The OfficeKit adapter uses this after it has decoded a validated
  // protobuf payload. It deliberately bypasses public source-bound mutation
  // checks; callers use set()/clear(), which remain capability-aware.
  _setImported(config) {
    this._value = config == null ? undefined : normalizePresentationTransition(config);
    return this;
  }
}
