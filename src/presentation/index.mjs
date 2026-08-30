import { inspectOoxmlPackage, ooxmlResolveRelationshipTarget, ooxmlSafePartPath, patchOoxmlPackage } from "../ooxml/package.mjs";
import { validatePptxPackageSemantics } from "../ooxml/pptx-package-semantics.mjs";
import { queryHelpRecords } from "../help/index.mjs";
import { Buffer } from "node:buffer";
import { FileBlob } from "../shared/file-blob.mjs";
import { toUint8Array } from "../shared/binary.mjs";
import { officeFontFamilies } from "../shared/font-design-metrics.mjs";
import { resolveColorToken } from "../shared/colors.mjs";
import { aid } from "../shared/ids.mjs";
import { imageDataFromDataUrl } from "../shared/images.mjs";
import { filterInspectRecords, inspectRecordMatchesTarget, inspectTargetTokens, ndjson, normalizeKinds, verificationIssue, verificationResult } from "../shared/inspection.mjs";
import { LAYOUT_MIME } from "../shared/render-output.mjs";
import { attrEscape, xmlEscape } from "../shared/xml.mjs";
import { createTextRange, textRangeRecord } from "../shared/text-range.mjs";
import { materializeComposeNode } from "./compose.mjs";
import { normalizePresentationThemeConfig } from "./ooxml-theme.mjs";
import { mergePresentationPlaceholders, normalizePresentationBackground, resolvePresentationBackgroundColor } from "./ooxml-masters.mjs";
import { isPresentationGradientFill, normalizePresentationGradientFill, presentationGradientFillSvg } from "./gradient-fills.mjs";
import { createPresentationGroupShapeClass } from "./group-shapes.mjs";
import { connectedPresentationShapeConfig, presentationConnectionSiteIndex, PresentationConnectorElement as ConnectorElement } from "./connectors.mjs";
import { createNativePresentationObjectClass } from "./native-objects.mjs";
import { normalizePresentationChartAxisGroup, normalizePresentationChartDataLabels, normalizePresentationChartErrorBars, normalizePresentationChartSeriesStyle, normalizePresentationChartStyle, normalizePresentationChartTrendlines } from "./ooxml-charts.mjs";
import { normalizePresentationChartExternalData, presentationChartUsesFormulaReferences } from "./ooxml-chart-data.mjs";
import { presentationChartLineSvgAttributes, presentationChartTrendlinesSvg } from "./chart-trendline-svg.mjs";
import { chartErrorBarMagnitudes } from "../shared/chart-error-bars.mjs";
import { planPresentationCustomShows, PresentationCustomShowCollection } from "./ooxml-custom-shows.mjs";
import { planPresentationSections, PresentationSectionCollection } from "./ooxml-sections.mjs";
import { SlideTransition } from "./ooxml-transitions.mjs";
import { SlideAnimations, SlideMorph } from "./ooxml-animations.mjs";
import { inheritPresentationParagraphs, normalizePresentationParagraphs, normalizePresentationParagraphStyles, presentationParagraphsNeedSerialization, presentationParagraphsSvg, presentationParagraphsText, replacePresentationParagraphText } from "./text-paragraphs.mjs";
import { normalizePresentationTextBodyProperties } from "./text-body-properties.mjs";
import { normalizePresentationCustomAdjustmentHandles, normalizePresentationCustomConnectionSites, normalizePresentationCustomPaths, normalizePresentationCustomTextRectangle, presentationCustomPathsSvg, presentationCustomTextRectangleFrame } from "./custom-geometry.mjs";
import { normalizePresentationCustomGeometryFormulaGraph } from "./custom-geometry-formulas.mjs";
import { normalizePresentationImageCrop, normalizePresentationImageFit, presentationImageCropViewport } from "./image-crop.mjs";
import { planPresentationModernComments } from "./ooxml-modern-comments.mjs";
import { presentationFreeLineSvg, presentationShapeLineSvgAttributes } from "./line-styles.mjs";
import { initializePresentationAccessibility, presentationAccessibilityCapability, setPresentationAccessibilityMetadata } from "./accessibility.mjs";
import { auditPresentationAccessibility } from "./accessibility-audit.mjs";
import { deletePresentationElement, PRESENTATION_ELEMENT_DELETED, presentationElementDeletionCapability } from "./element-deletion.mjs";
import { assertPresentationElementIndexes, installPresentationElementOrdering } from "./element-order.mjs";
import { editSvgText as replaceSvgTextNode, inspectSvgText } from "./svg-text.mjs";
import { editSvgLeaf as replaceSvgLeaf, inspectSvgLeaves } from "./svg-leaves.mjs";
import { buildPresentationDesignProfile } from "./design-profile.mjs";
import { buildTemplateGenerationPlan } from "./template-plan.mjs";
import { classifyImportedPresentationObjects } from "./import-object-classification.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const EMU_PER_PIXEL = 9_525;
const importedShapeBackgroundFill = new WeakMap();
const PRESENTATION_SLIDE_DUPLICATOR = Symbol.for("office-kit.presentation-duplicate");
const PRESENTATION_SPEAKER_NOTES_CAPABILITY = Symbol.for("office-kit.speaker-notes-capability");
const PRESENTATION_LEGACY_COMMENTS_CAPABILITY = Symbol.for("office-kit.legacy-comments-capability");
const PRESENTATION_SLIDE_VISIBILITY_CAPABILITY = Symbol.for("office-kit.slide-visibility-capability");
const PRESENTATION_SLIDE_DELETION_CAPABILITY = Symbol.for("office-kit.slide-deletion-capability");
const PRESENTATION_SLIDE_CLONE_CAPABILITY = Symbol.for("office-kit.slide-clone-capability");
const PRESENTATION_SLIDE_CONTINUATION_CAPABILITY = Symbol.for("office-kit.slide-continuation-capability");
const PRESENTATION_STATE = Symbol.for("office-kit.presentation-state");
const PRESENTATION_NATIVE_LEAF_CAPABILITY = Symbol.for("office-kit.presentation-native-leaf-capability");
const PRESENTATION_COMPONENT_CAPABILITY = Symbol.for("office-kit.presentation-component-capability");
const PRESENTATION_IMAGE_DATA_URL_SOURCE = Symbol.for("office-kit.presentation-image-data-url-source");
const PRESENTATION_IMAGE_SVG_DATA_URL_SOURCE = Symbol.for("office-kit.presentation-image-svg-data-url-source");
export const PRESENTATION_IMPORTED_THEME_PROFILE = Symbol.for("office-kit.presentation-imported-theme-profile");

// The imported theme is descriptive source evidence, not the mutable
// authoring theme. Keeping it out of the wire model lets source-bound exports
// preserve the original theme bytes while designProfile still reflects the
// visual language of a third-party deck.
export function setPresentationImportedThemeProfile(presentation, profile) {
  if (!presentation || typeof presentation !== "object") throw new TypeError("Imported presentation theme requires a Presentation instance.");
  if (profile !== undefined && (typeof profile !== "object" || profile.kind !== "theme")) {
    throw new TypeError("Imported presentation theme profile must be a theme record.");
  }
  Object.defineProperty(presentation, PRESENTATION_IMPORTED_THEME_PROFILE, {
    configurable: true,
    value: profile,
  });
  return presentation;
}

export { SlideTransition, SlideAnimations, SlideMorph };

const PPTX_PACKAGE_CONFIG = {
  family: "PPTX",
  packageKind: "pptxPackage",
  partKind: "pptxPart",
  officeDocument: {
    contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    partPath: "ppt/presentation.xml",
  },
  counts: { slides: /^ppt\/slides\/slide\d+\.xml$/ },
  semanticIssues: validatePptxPackageSemantics,
};

class SlideCollection {
  constructor(presentation) {
    this.presentation = presentation;
    this.items = [];
  }

  add(options = {}) {
    return this.#insertAt(options, this.items.length);
  }

  insert(options = {}) {
    if (!options || typeof options !== "object" || Array.isArray(options)) throw new TypeError("Presentation slide options must be an object.");
    const { after, ...slideOptions } = options;
    let index = this.items.length;
    if (after === null) {
      index = 0;
    } else if (after !== undefined) {
      if (after instanceof Slide) {
        index = this.items.indexOf(after);
        if (index < 0 || after.presentation !== this.presentation) throw new Error("Presentation slide insertion target must belong to this presentation.");
      } else if (Number.isInteger(after) && after >= 0 && after < this.items.length) {
        index = after;
      } else {
        throw new RangeError("Presentation slide insertion after must be an existing Slide, a 0-based slide index, or null.");
      }
      index += 1;
    }
    return this.#insertAt(slideOptions, index);
  }

  #insertAt(options, index) {
    if (!options || typeof options !== "object" || Array.isArray(options)) throw new TypeError("Presentation slide options must be an object.");
    const slide = new Slide(this.presentation, options);
    const requestedLayout = options.layout ?? options.layoutId;
    if (requestedLayout == null) {
      this.items.splice(index, 0, slide);
      return slide;
    }
    const layout = typeof requestedLayout === "string"
      ? this.presentation.layouts.getItem(requestedLayout)
      : requestedLayout;
    if (!(layout instanceof SlideLayoutTemplate) || layout.presentation !== this.presentation) {
      throw new Error(`Unknown presentation layout: ${typeof requestedLayout === "string" ? requestedLayout : "provided layout"}`);
    }
    this.items.splice(index, 0, slide);
    try {
      layout.apply(slide);
    } catch (error) {
      this.items.splice(index, 1);
      throw error;
    }
    return slide;
  }

  getItem(index) { return this.items[index]; }
  get count() { return this.items.length; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

class PresentationTheme {
  constructor(presentation, config = {}, base = {}) {
    const normalized = normalizePresentationThemeConfig(config, base);
    this.presentation = presentation;
    this.id = config.id || "theme/default";
    this.name = normalized.name;
    this.colors = normalized.colors;
    this.fonts = normalized.fonts;
    this.textStyles = normalized.textStyles;
    this.colorMap = normalized.colorMap;
  }

  update(config = {}) {
    const normalized = normalizePresentationThemeConfig(config, this);
    Object.assign(this, normalized);
    return this;
  }

  setColors(colors = {}) { return this.update({ colors }); }
  setFonts(fonts = {}) { return this.update({ fonts }); }
  setTextStyles(textStyles = {}) { return this.update({ textStyles }); }
  setColorMap(colorMap = {}) { return this.update({ colorMap }); }
  inspectRecord() { return { kind: "theme", id: this.id, name: this.name, colors: this.colors, fonts: this.fonts, textStyles: this.textStyles, colorMap: this.colorMap }; }
  toJSON() { return { id: this.id, name: this.name, colors: this.colors, fonts: this.fonts, textStyles: this.textStyles, colorMap: this.colorMap }; }
}

function presentationThemeSemantics(theme) {
  const normalized = normalizePresentationThemeConfig(theme);
  return JSON.stringify({ name: normalized.name, colors: normalized.colors, fonts: normalized.fonts, textStyles: normalized.textStyles, colorMap: normalized.colorMap });
}

function normalizePresentationPlaceholderTransform(value, name = "Presentation placeholder transform") {
  if (value == null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${name} must be an object.`);
  const output = {};
  if (Object.hasOwn(value, "rotationDegrees") && value.rotationDegrees != null) {
    const degrees = Number(value.rotationDegrees);
    if (!Number.isFinite(degrees) || degrees < -360 || degrees > 360) throw new RangeError(`${name}.rotationDegrees must be between -360 and 360 degrees.`);
    output.rotationDegrees = degrees;
  }
  for (const key of ["flipHorizontal", "flipVertical"]) {
    if (!Object.hasOwn(value, key) || value[key] == null) continue;
    if (typeof value[key] !== "boolean") throw new TypeError(`${name}.${key} must be a boolean.`);
    output[key] = value[key];
  }
  if (Object.keys(output).length === 0) throw new TypeError(`${name} must define rotationDegrees, flipHorizontal, or flipVertical.`);
  return output;
}

const PRESENTATION_PLACEHOLDER_TYPE_ALIASES = new Map([
  ["title", "title"],
  ["body", "body"],
  ["ctrTitle", "ctrTitle"],
  ["centeredTitle", "ctrTitle"],
  ["subTitle", "subTitle"],
  ["subtitle", "subTitle"],
  ["dt", "dt"],
  ["dateAndTime", "dt"],
  ["sldNum", "sldNum"],
  ["slideNumber", "sldNum"],
  ["ftr", "ftr"],
  ["footer", "ftr"],
  ["hdr", "hdr"],
  ["header", "hdr"],
  ["obj", "obj"],
  ["object", "obj"],
  ["chart", "chart"],
  ["tbl", "tbl"],
  ["table", "tbl"],
  ["clipArt", "clipArt"],
  ["dgm", "dgm"],
  ["diagram", "dgm"],
  ["media", "media"],
  ["sldImg", "sldImg"],
  ["slideImage", "sldImg"],
  ["pic", "pic"],
  ["picture", "pic"],
]);

function normalizePresentationPlaceholderType(value) {
  const requested = String(value || "body").trim();
  const type = PRESENTATION_PLACEHOLDER_TYPE_ALIASES.get(requested);
  if (!type) throw new TypeError(`Unsupported Presentation placeholder type: ${requested || "(empty)"}.`);
  return type;
}

function presentationPlaceholderLookup(placeholders, idOrName) {
  const key = String(idOrName ?? "");
  const index = Number(idOrName);
  return placeholders.find((placeholder) =>
    placeholder.id === idOrName || placeholder.name === idOrName || placeholder.type === idOrName ||
    (Number.isInteger(index) && placeholder.idx === index));
}

function presentationPlaceholderSummary(owner, placeholders) {
  const items = placeholders.map((placeholder) => ({
    id: placeholder.id,
    name: placeholder.name,
    type: placeholder.type,
    idx: placeholder.idx,
    index: placeholder.idx,
    required: Boolean(placeholder.required),
    hasDirectPosition: Boolean(placeholder.position),
    ...(placeholder.position ? { position: { ...placeholder.position } } : {}),
  }));
  return {
    ownerId: owner.id,
    count: items.length,
    requiredCount: items.filter((placeholder) => placeholder.required).length,
    types: [...new Set(items.map((placeholder) => placeholder.type))].sort(),
    items,
  };
}

function attachPresentationPlaceholderCollectionApi(owner, placeholders, { allowMissingPosition = false } = {}) {
  Object.defineProperties(placeholders, {
    add: {
      enumerable: false,
      value(config = {}) {
        if (!config || typeof config !== "object" || Array.isArray(config)) throw new TypeError("Presentation placeholder config must be an object.");
        const placeholder = normalizePresentationPlaceholders([{
          ...config,
          id: config.id || `${owner.id}/ph/${placeholders.length + 1}`,
          idx: config.idx ?? config.index ?? placeholders.length,
        }], `${owner.id}/ph`, { allowMissingPosition })[0];
        if (placeholders.some((current) => current.id === placeholder.id || (current.type === placeholder.type && current.idx === placeholder.idx))) {
          throw new Error(`Presentation placeholder ${placeholder.name || placeholder.id} duplicates an existing id or type/index pair.`);
        }
        placeholders.push(placeholder);
        return placeholder;
      },
    },
    getItem: { enumerable: false, value(idOrName) { return presentationPlaceholderLookup(placeholders, idOrName); } },
    summary: { enumerable: false, value() { return presentationPlaceholderSummary(owner, placeholders); } },
    count: { enumerable: false, get() { return placeholders.length; } },
  });
  return placeholders;
}

function normalizePresentationPlaceholders(value = [], idPrefix = "placeholder", options = {}) {
  if (!Array.isArray(value)) throw new TypeError("Presentation placeholders must be an array.");
  if (value.length > 128) throw new RangeError("Presentation placeholders exceed 128 entries.");
  const placeholders = value.map((placeholder, index) => {
    if (!placeholder || typeof placeholder !== "object" || Array.isArray(placeholder)) throw new TypeError("Presentation placeholder entries must be objects.");
    const position = options.allowMissingPosition && !placeholder.position && !placeholder.frame && !["left", "top", "width", "height"].some((key) => placeholder[key] != null)
      ? undefined
      : normalizeFrame(placeholder, { left: 80, top: 80 + index * 80, width: 640, height: 64 });
    const transform = normalizePresentationPlaceholderTransform(placeholder.transform, `Presentation placeholder ${placeholder.name || index + 1} transform`);
    if (transform && !position) throw new TypeError(`Presentation placeholder ${placeholder.name || index + 1} cannot define a transform without a direct position.`);
    return {
      id: placeholder.id || `${idPrefix}/${index + 1}`,
      type: normalizePresentationPlaceholderType(placeholder.type),
      idx: Number(placeholder.idx ?? placeholder.index ?? index + 1),
      name: placeholder.name || `${normalizePresentationPlaceholderType(placeholder.type)} placeholder`,
      position,
      transform,
      text: placeholder.text ?? "",
      required: Boolean(placeholder.required),
      style: { ...(placeholder.style || {}) },
      paragraphStyles: normalizePresentationParagraphStyles(placeholder.paragraphStyles || placeholder.listStyles || {}),
      textBodyProperties: normalizePresentationTextBodyProperties(placeholder.textBodyProperties || placeholder.bodyProperties || {}),
    };
  });
  if (placeholders.some((placeholder) => !Number.isInteger(placeholder.idx) || placeholder.idx < 0 || placeholder.idx > 4_294_967_295)) throw new RangeError("Presentation placeholder idx must be an unsigned 32-bit integer.");
  if (new Set(placeholders.map((placeholder) => `${placeholder.type}:${placeholder.idx}`)).size !== placeholders.length) throw new Error("Presentation placeholder type/idx pairs must be unique.");
  return placeholders;
}

function clonePresentationParagraphStyles(styles = {}) {
  return Object.fromEntries(Object.entries(styles).map(([level, style]) => [Number(level), { ...style, style: { ...(style.style || {}) } }]));
}

function mergePresentationParagraphStyles(base = {}, overrides = {}) {
  const result = clonePresentationParagraphStyles(base);
  for (const [level, style] of Object.entries(overrides || {})) {
    const inherited = { ...(result[Number(level)] || {}) };
    if (["bulletCharacter", "bulletImage", "autoNumber", "bulletNone"].some((field) => Object.hasOwn(style, field))) {
      delete inherited.bulletCharacter;
      delete inherited.bulletImage;
      delete inherited.autoNumber;
      delete inherited.bulletNone;
    }
    for (const fields of [["bulletFont", "bulletFontFollowText"], ["bulletColor", "bulletColorFollowText"], ["bulletSize", "bulletSizePercent", "bulletSizeFollowText"]]) {
      if (!fields.some((field) => Object.hasOwn(style, field))) continue;
      for (const field of fields) delete inherited[field];
    }
    result[Number(level)] = { ...inherited, ...style, style: { ...(inherited.style || {}), ...(style.style || {}) } };
  }
  return result;
}

function normalizePresentationMasterParagraphStyles(value = {}) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError("Presentation master textParagraphStyles must be an object.");
  return Object.fromEntries(["title", "body", "other"].map((kind) => [kind, normalizePresentationParagraphStyles(value[kind] || {})]));
}

function presentationPlaceholderTextStyleKind(type = "body") {
  if (["title", "ctrTitle"].includes(type)) return "title";
  if (["body", "subTitle", "obj", "chart", "tbl", "clipArt", "dgm", "media", "pic"].includes(type)) return "body";
  return "other";
}

function normalizePresentationSlideGuides(value = []) {
  if (!Array.isArray(value)) throw new TypeError("Presentation slideGuides must be an array.");
  if (value.length > 1024) throw new RangeError("Presentation slideGuides exceed 1,024 entries.");
  return Object.freeze(value.map((guide) => {
    if (!guide || !["horizontal", "vertical"].includes(guide.orientation)) {
      throw new TypeError("Presentation guide orientation must be horizontal or vertical.");
    }
    const position = Number(guide.position);
    if (!Number.isInteger(position) || position < -2_147_483_648 || position > 2_147_483_647) {
      throw new RangeError("Presentation guide position must be a signed 32-bit integer.");
    }
    return Object.freeze({ orientation: guide.orientation, position });
  }));
}

const EMPTY_PRESENTATION_SLIDE_GUIDES = Object.freeze([]);

function normalizePresentationViewSourceBinding(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  return Object.freeze({
    partPath: String(value.partPath || ""),
    relationshipId: String(value.relationshipId || ""),
    viewXmlSha256: String(value.viewXmlSha256 || ""),
    semanticSha256: String(value.semanticSha256 || ""),
    residualSha256: String(value.residualSha256 || ""),
    editable: value.editable === true,
  });
}

function clonePresentationViewSourceBinding(value) {
  return value ? { ...value } : undefined;
}

function normalizePresentationViewGridSpacing(value, field) {
  const spacing = Number(value);
  if (!Number.isSafeInteger(spacing) || spacing <= 0 || spacing > 2_147_483_647) {
    throw new RangeError(`Presentation ${field} must be a positive signed 32-bit EMU integer.`);
  }
  return spacing;
}

class PresentationView {
  #presentation;
  #gridlinesVisible = false;
  #guidesVisible = false;
  #sourceBinding;

  constructor(presentation) { this.#presentation = presentation; }
  get gridlinesVisible() { return this.#gridlinesVisible; }
  get guidesVisible() { return this.#guidesVisible; }
  get gridSpacingCxEmu() { return this.#presentation._viewProperties?.gridSpacingCxEmu; }
  get gridSpacingCyEmu() { return this.#presentation._viewProperties?.gridSpacingCyEmu; }
  get slideViewSnapToGrid() { return this.#presentation._viewProperties?.slideViewSnapToGrid; }
  get slideViewSnapToObjects() { return this.#presentation._viewProperties?.slideViewSnapToObjects; }
  get slideGuides() { return this.#presentation._viewProperties?.slideGuides || EMPTY_PRESENTATION_SLIDE_GUIDES; }
  get capability() {
    const properties = this.#presentation._viewProperties || {};
    return {
      sourceBound: Boolean(this.#sourceBinding),
      partPresent: Boolean(this.#sourceBinding),
      editable: this.#sourceBinding?.editable === true,
      gridSpacingCxEmuPresent: Object.hasOwn(properties, "gridSpacingCxEmu"),
      gridSpacingCyEmuPresent: Object.hasOwn(properties, "gridSpacingCyEmu"),
      slideViewSnapToGridPresent: Object.hasOwn(properties, "slideViewSnapToGrid"),
      slideViewSnapToObjectsPresent: Object.hasOwn(properties, "slideViewSnapToObjects"),
      guideCount: properties.slideGuides?.length || 0,
    };
  }
  showGridlines() { this.#gridlinesVisible = true; }
  hideGridlines() { this.#gridlinesVisible = false; }
  toggleGridlines() { this.#gridlinesVisible = !this.#gridlinesVisible; return this.#gridlinesVisible; }
  showGuides() { this.#guidesVisible = true; this.#hideGuidesOnExport(); }
  hideGuides() { this.#guidesVisible = false; this.#hideGuidesOnExport(); }
  toggleGuides() { this.#guidesVisible = !this.#guidesVisible; this.#hideGuidesOnExport(); return this.#guidesVisible; }
  setSourceProperties(patch) {
    if (!patch || typeof patch !== "object" || Array.isArray(patch)) {
      throw new TypeError("Presentation view source properties must be an object.");
    }
    const supported = new Set(["gridSpacingCxEmu", "gridSpacingCyEmu", "slideViewSnapToGrid", "slideViewSnapToObjects", "slideGuides"]);
    const unsupported = Object.keys(patch).filter((key) => !supported.has(key));
    if (unsupported.length) throw new TypeError(`Presentation view source properties have unsupported fields: ${unsupported.join(", ")}.`);
    if (!Object.keys(patch).length) throw new TypeError("Presentation view source properties must include at least one editable field.");
    const capability = this.capability;
    if (!capability.sourceBound || !this.#presentation._viewProperties) {
      throw new Error("Presentation view properties can be changed only on an imported PPTX view-properties part.");
    }
    if (!capability.editable) {
      throw new Error("Presentation view properties are source-bound and do not match the fixed-topology editable profile.");
    }
    const current = this.#presentation._viewProperties;
    const next = { ...current, slideGuides: current.slideGuides || EMPTY_PRESENTATION_SLIDE_GUIDES };
    for (const field of ["gridSpacingCxEmu", "gridSpacingCyEmu"]) {
      if (!Object.hasOwn(patch, field)) continue;
      if (!Object.hasOwn(current, field)) throw new Error(`Imported presentation view does not contain ${field}; this method cannot add it.`);
      next[field] = normalizePresentationViewGridSpacing(patch[field], field);
    }
    for (const field of ["slideViewSnapToGrid", "slideViewSnapToObjects"]) {
      if (!Object.hasOwn(patch, field)) continue;
      if (!Object.hasOwn(current, field)) throw new Error(`Imported presentation view does not contain ${field}; this method cannot add it.`);
      if (typeof patch[field] !== "boolean") throw new TypeError(`Presentation ${field} must be a boolean.`);
      next[field] = patch[field];
    }
    if (Object.hasOwn(patch, "slideGuides")) {
      const guides = normalizePresentationSlideGuides(patch.slideGuides);
      if (guides.length !== current.slideGuides.length || guides.some((guide, index) => guide.orientation !== current.slideGuides[index].orientation)) {
        throw new Error("Imported presentation view guide count, order, and orientation are source-bound.");
      }
      next.slideGuides = guides;
    }
    this.#presentation._viewProperties = {
      ...next,
      // The local editor visibility switches are intentionally not a file edit.
      slideViewShowGuides: false,
      slideGuides: next.slideGuides,
    };
    return this;
  }
  toProto() {
    const source = this.#presentation._viewProperties;
    if (!source) return undefined;
    const { source: _source, ...properties } = source;
    return { ...properties, slideGuides: source.slideGuides?.map((guide) => ({ ...guide })) || [] };
  }
  _sourceBindingForExport() { return clonePresentationViewSourceBinding(this.#sourceBinding); }
  _setImportedProperties(properties) {
    this.#sourceBinding = normalizePresentationViewSourceBinding(properties?.source);
    const { source: _source, ...viewProperties } = properties || {};
    this.#presentation._viewProperties = properties ? {
      ...viewProperties,
      slideViewShowGuides: false,
      slideGuides: normalizePresentationSlideGuides(properties.slideGuides),
    } : undefined;
  }
  #hideGuidesOnExport() {
    this.#presentation._viewProperties = {
      ...(this.#presentation._viewProperties || {}),
      slideViewShowGuides: false,
      slideGuides: this.#presentation._viewProperties?.slideGuides || Object.freeze([]),
    };
  }
}

class PresentationSlideMaster {
  constructor(presentation, config = {}) {
    this.presentation = presentation;
    this.configured = Object.keys(config).length > 0;
    this.id = config.id || "master/default";
    this.name = config.name || "Default Master";
    this.theme = config.theme ? new PresentationTheme(presentation, { ...config.theme, id: config.theme.id || `${this.id}/theme` }, presentation.theme) : undefined;
    Object.defineProperty(this, "_backgroundClearRequested", { value: false, writable: true });
    this.background = Object.hasOwn(config, "background")
      ? normalizePresentationBackground(config.background)
      : normalizePresentationBackground(presentation.theme.colors.bg1);
    this.placeholders = attachPresentationPlaceholderCollectionApi(this, normalizePresentationPlaceholders(config.placeholders || [], `${this.id}/ph`));
    this.textParagraphStyles = normalizePresentationMasterParagraphStyles(config.textParagraphStyles || {});
    Object.defineProperty(this, "slideGuides", { value: normalizePresentationSlideGuides(config.slideGuides), enumerable: true });
  }

  update(config = {}) {
    if (Object.keys(config).length > 0) this.configured = true;
    const previousId = this.id;
    if (config.id) this.id = String(config.id);
    if (this.theme?.id === `${previousId}/theme`) this.theme.id = `${this.id}/theme`;
    if (config.name) this.name = String(config.name);
    if (Object.hasOwn(config, "theme")) this.theme = config.theme ? new PresentationTheme(this.presentation, { ...config.theme, id: config.theme.id || `${this.id}/theme` }, this.presentation.theme) : undefined;
    if (Object.hasOwn(config, "background")) {
      this.background = config.background == null ? undefined : normalizePresentationBackground(config.background, this.background);
      this._backgroundClearRequested = false;
    }
    if (config.placeholders) this.placeholders = attachPresentationPlaceholderCollectionApi(this, normalizePresentationPlaceholders(config.placeholders, `${this.id}/ph`));
    if (config.textParagraphStyles) this.textParagraphStyles = normalizePresentationMasterParagraphStyles(config.textParagraphStyles);
    return this;
  }

  setBackground(background) { this.configured = true; this.background = normalizePresentationBackground(background, this.background); this._backgroundClearRequested = false; return this; }
  setNativeBackgroundImage(config = {}) { this.configured = true; this.background = normalizeNativePresentationBackgroundImage(config, `Presentation master ${this.id}`, this.background); this._backgroundClearRequested = false; return this; }
  clearBackground() { this.configured = true; this.background = undefined; this._backgroundClearRequested = true; return this; }
  clearNativeBackgroundImage() { this.configured = true; if (presentationBackgroundHasImage(this.background)) this.clearBackground(); return this; }
  setTheme(theme) { this.configured = true; this.theme = theme ? new PresentationTheme(this.presentation, { ...theme, id: theme.id || `${this.id}/theme` }, this.presentation.theme) : undefined; return this; }
  effectiveTheme() { return this.theme || this.presentation.theme; }
  effectiveBackground() { return this.background?.fill || this.background?.gradient ? this.background : normalizePresentationBackground(this.effectiveTheme().colors.bg1, "#ffffff"); }
  effectiveBackgroundImage() { return presentationBackgroundHasImage(this.background) ? this.background : undefined; }
  paragraphStylesForPlaceholder(type) { return this.textParagraphStyles[presentationPlaceholderTextStyleKind(type)] || {}; }
  inspectRecord() { const theme = this.effectiveTheme(); return { kind: "slideMaster", id: this.id, name: this.name, background: this.background, nativeBackgroundImage: presentationBackgroundHasImage(this.background) ? { fit: "stretch", editable: true, inherited: false } : undefined, effectiveBackground: this.effectiveBackground(), placeholders: this.placeholders.length, placeholderTypes: this.placeholders.map((placeholder) => placeholder.type), slideGuides: this.slideGuides.length, textParagraphStyleLevels: Object.fromEntries(Object.entries(this.textParagraphStyles).map(([kind, styles]) => [kind, Object.keys(styles).length])), hasThemeOverride: Boolean(this.theme), themeId: theme.id, themeName: theme.name }; }
  toJSON() { return { id: this.id, name: this.name, background: this.background, theme: this.theme?.toJSON(), placeholders: this.placeholders.map((placeholder) => ({ ...placeholder })), slideGuides: this.slideGuides.map((guide) => ({ ...guide })), textParagraphStyles: normalizePresentationMasterParagraphStyles(this.textParagraphStyles) }; }
}

class PresentationSlideMasterCollection {
  constructor(presentation) { this.presentation = presentation; this.items = []; }
  add(config = {}) {
    if (this.items.length >= 64) throw new RangeError("Presentation masters exceed 64 entries.");
    const master = config instanceof PresentationSlideMaster ? config : new PresentationSlideMaster(this.presentation, config);
    if (this.items.some((item) => item.id === master.id)) throw new Error(`Duplicate presentation master ID ${master.id}.`);
    master.presentation = this.presentation;
    if (master.theme) master.theme.presentation = this.presentation;
    this.items.push(master);
    return master;
  }
  getItem(idOrName) { return this.items.find((master) => master.id === idOrName || master.name === idOrName); }
  get count() { return this.items.length; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

class SlideLayoutTemplate {
  constructor(presentation, config = {}) {
    this.presentation = presentation;
    this.id = config.id || aid("lo");
    this.name = config.name || "Blank";
    this.type = config.type || "blank";
    this.masterId = config.masterId || presentation.master.id;
    Object.defineProperty(this, "_backgroundClearRequested", { value: false, writable: true });
    this.background = config.background ? normalizePresentationBackground(config.background) : undefined;
    this.placeholders = attachPresentationPlaceholderCollectionApi(this, normalizePresentationPlaceholders(config.placeholders || [], `${this.id}/ph`, { allowMissingPosition: true }), { allowMissingPosition: true });
    Object.defineProperty(this, "slideGuides", { value: normalizePresentationSlideGuides(config.slideGuides), enumerable: true });
  }

  effectiveMaster() { return this.presentation.masters.getItem(this.masterId); }
  effectiveTheme() { return this.effectiveMaster()?.effectiveTheme() || this.presentation.theme; }
  setBackground(background) { this.background = normalizePresentationBackground(background, this.background); this._backgroundClearRequested = false; return this; }
  setNativeBackgroundImage(config = {}) { this.background = normalizeNativePresentationBackgroundImage(config, `Presentation layout ${this.id}`, this.background); this._backgroundClearRequested = false; return this; }
  clearBackground() { this.background = undefined; this._backgroundClearRequested = true; return this; }
  clearNativeBackgroundImage() { if (presentationBackgroundHasImage(this.background)) this.clearBackground(); return this; }
  effectivePlaceholders() {
    const master = this.effectiveMaster();
    return mergePresentationPlaceholders(master?.placeholders || [], this.placeholders).map((placeholder) => ({
      ...placeholder,
      paragraphStyles: mergePresentationParagraphStyles(master?.paragraphStylesForPlaceholder(placeholder.type), placeholder.paragraphStyles),
    }));
  }
  effectiveBackground() { return this.background?.fill || this.background?.gradient ? this.background : this.effectiveMaster()?.effectiveBackground() || normalizePresentationBackground(this.presentation.theme.colors.bg1, "#ffffff"); }
  effectiveBackgroundImage() {
    if (presentationBackgroundHasImage(this.background)) return this.background;
    if (this.background?.fill || this.background?.gradient) return undefined;
    return this.effectiveMaster()?.effectiveBackgroundImage();
  }

  apply(slide) {
    if (!(slide instanceof Slide) || slide.presentation !== this.presentation) {
      throw new TypeError("Presentation layouts can only be applied to a slide from the same presentation.");
    }
    const materializedOtherLayout = slide.shapes.items.find((shape) =>
      shape.placeholder?.layoutId && shape.placeholder.layoutId !== this.id);
    if (materializedOtherLayout) {
      throw new Error(`Slide ${slide.id} already has materialized placeholders from layout ${materializedOtherLayout.placeholder.layoutId}; changing layouts would leave an ambiguous placeholder topology.`);
    }
    slide.layoutId = this.id;
    const placeholders = this.effectivePlaceholders();
    return placeholders.map((placeholder) => {
      const existing = slide.shapes.items.find((shape) =>
        shape.placeholder?.layoutId === this.id &&
        shape.placeholder?.type === placeholder.type &&
        Number(shape.placeholder?.idx) === placeholder.idx);
      if (existing) return existing;
      const shape = slide.shapes.add({
        id: placeholder.id,
        name: placeholder.name,
        geometry: "rect",
        position: placeholder.position,
        transform: placeholder.transform,
        fill: "transparent",
        line: { fill: "transparent", width: 0 },
        text: placeholder.text,
        textBodyProperties: placeholder.textBodyProperties,
        placeholder: { layoutId: this.id, type: placeholder.type, name: placeholder.name, required: placeholder.required, idx: placeholder.idx },
      });
      shape.text.style = { ...placeholder.style };
      shape.text.inheritedParagraphStyles = Object.fromEntries(Object.entries(placeholder.paragraphStyles || {}).map(([level, style]) => [level, { ...style, style: { ...(style.style || {}) } }]));
      return shape;
    });
  }

  inspectRecord() { const directImage = presentationBackgroundHasImage(this.background); const inheritedImage = !directImage && !this.background?.fill && !this.background?.gradient && presentationBackgroundHasImage(this.effectiveMaster()?.effectiveBackgroundImage()); return { kind: "layoutTemplate", id: this.id, name: this.name, type: this.type, masterId: this.masterId, themeId: this.effectiveTheme().id, background: this.background, nativeBackgroundImage: directImage || inheritedImage ? { fit: "stretch", editable: directImage, inherited: inheritedImage } : undefined, effectiveBackground: this.effectiveBackground(), placeholders: this.placeholders.length, effectivePlaceholders: this.effectivePlaceholders().length, placeholderTypes: this.effectivePlaceholders().map((placeholder) => placeholder.type), slideGuides: this.slideGuides.length }; }
  toJSON() { return { id: this.id, name: this.name, type: this.type, masterId: this.masterId, background: this.background, placeholders: this.placeholders.map((placeholder) => ({ ...placeholder })), slideGuides: this.slideGuides.map((guide) => ({ ...guide })) }; }
}

class SlideLayoutCollection {
  constructor(presentation) { this.presentation = presentation; this.items = []; }
  add(config = {}) {
    const normalized = typeof config === "string" ? { name: config } : config;
    if (!normalized || typeof normalized !== "object" || Array.isArray(normalized)) throw new TypeError("Presentation layout config must be an object or name string.");
    const layout = new SlideLayoutTemplate(this.presentation, normalized);
    this.items.push(layout);
    return layout;
  }
  getItem(idOrName) { return this.items.find((layout) => layout.id === idOrName || layout.name === idOrName || layout.type === idOrName); }
  getById(id) { return this.items.find((layout) => layout.id === id); }
  inspectRecords() { return this.items.map((layout) => layout.inspectRecord()); }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

function svgInner(svg = "") {
  return String(svg || "").replace(/^<svg\b[^>]*>/i, "").replace(/<\/svg>\s*$/i, "");
}

function presentationMontageSvg(presentation, options = {}) {
  const slides = presentation.slides.items.length ? presentation.slides.items : [presentation.slides.add()];
  const gap = Number(options.gap ?? 24);
  const scale = Number(options.scale ?? 0.25);
  const columns = Math.max(1, Number(options.columns ?? 1) || 1);
  const slideW = Number(presentation.slideSize.width || 1280);
  const slideH = Number(presentation.slideSize.height || 720);
  const thumbW = slideW * scale;
  const thumbH = slideH * scale;
  const labelH = 20;
  const rows = Math.ceil(slides.length / columns);
  const width = Math.max(1, columns * thumbW + (columns + 1) * gap);
  const height = Math.max(1, rows * (thumbH + labelH) + (rows + 1) * gap);
  const thumbs = slides.map((slide, index) => {
    const col = index % columns;
    const row = Math.floor(index / columns);
    const x = gap + col * (thumbW + gap);
    const y = gap + row * (thumbH + labelH + gap);
    return `<g data-slide="${index + 1}"><rect x="${x - 1}" y="${y - 1}" width="${thumbW + 2}" height="${thumbH + 2}" fill="#ffffff" stroke="#94a3b8"/><g transform="translate(${x},${y}) scale(${scale})">${svgInner(slide.toSvg())}</g><text x="${x}" y="${y + thumbH + 15}" font-family="Arial" font-size="12" fill="#475569">Slide ${index + 1}${slide.title() ? ` — ${xmlEscape(slide.title()).slice(0, 80)}` : ""}</text></g>`;
  }).join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="#f8fafc"/>${thumbs}</svg>`;
}

export class Presentation {
  constructor(options = {}) {
    this.id = aid("pr");
    this.slideSize = options.slideSize || { width: 1280, height: 720 };
    this.commentFormat = options.commentFormat || "legacy";
    this.theme = new PresentationTheme(this, options.theme || {});
    this.masters = new PresentationSlideMasterCollection(this);
    const masterConfigs = Array.isArray(options.masters) && options.masters.length ? options.masters : [options.master || {}];
    for (const master of masterConfigs) this.masters.add(master);
    this.layouts = new SlideLayoutCollection(this);
    for (const layout of options.layouts || []) this.layouts.add(layout);
    this.slides = new SlideCollection(this);
    this.customShows = new PresentationCustomShowCollection(this);
    this.sections = new PresentationSectionCollection(this);
    Object.defineProperty(this, "_viewProperties", { value: undefined, writable: true });
    this.view = new PresentationView(this);
  }

  static create(options = {}) { return new Presentation(options); }
  get fontFamilies() { return officeFontFamilies([this.toProto()]); }
  get master() { return this.masters.items[0]; }
  set master(value) {
    const master = value instanceof PresentationSlideMaster ? value : new PresentationSlideMaster(this, value || {});
    master.presentation = this;
    if (master.theme) master.theme.presentation = this;
    if (this.masters.items.length) this.masters.items[0] = master;
    else this.masters.items.push(master);
  }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["deck", "slide", "textbox", "shape", "nativeObject", "layout"]);
    const records = [];
    if (kinds.has("deck")) records.push({ kind: "deck", id: this.id, slides: this.slides.count, customShows: this.customShows.count, sections: this.sections.count });
    if (kinds.has("theme")) records.push(this[PRESENTATION_IMPORTED_THEME_PROFILE] || this.theme.inspectRecord());
    if (kinds.has("slideMaster") || kinds.has("master")) records.push(...this.masters.items.map((master) => master.inspectRecord()));
    if (kinds.has("layout") || kinds.has("layoutTemplate")) records.push(...this.layouts.inspectRecords());
    if (kinds.has("customShow")) records.push(...this.customShows.items.map((show) => show.inspectRecord()));
    if (kinds.has("section")) records.push(...this.sections.items.map((section) => section.inspectRecord()));
    for (const slide of this.slides) records.push(...slide.inspectRecords(kinds));
    if (options.includeNativeLeaves === true || kinds.has("nativeLeaf")) {
      const capability = this[PRESENTATION_NATIVE_LEAF_CAPABILITY];
      if (!capability) {
        const error = new Error("Presentation native leaves are available only for a trusted imported PPTX source revision.");
        error.code = "presentation_native_leaf_source_required";
        throw error;
      }
      records.push(...capability.inspect());
    }
    if (options.includeComponentCandidates === true || kinds.has("componentCandidate")) {
      const capability = this[PRESENTATION_COMPONENT_CAPABILITY];
      if (!capability) {
        const error = new Error("Presentation component candidates are available only for a trusted imported PPTX source revision.");
        error.code = "presentation_component_source_required";
        throw error;
      }
      records.push(...capability.inspect());
    }
    if (options.includeImportObjects === true || kinds.has("importObject")) {
      const state = this[PRESENTATION_STATE];
      if (!state) {
        const error = new Error("Imported-object classification requires a trusted PPTX source revision.");
        error.code = "presentation_import_object_source_required";
        throw error;
      }
      records.push(...classifyImportedPresentationObjects(state, {
        nativeLeafRecords: this[PRESENTATION_NATIVE_LEAF_CAPABILITY]?.inspect?.() || [],
        componentRecords: this[PRESENTATION_COMPONENT_CAPABILITY]?.inspect?.() || [],
        includeNested: options.includeNested === true,
      }));
    }
    return ndjson(filterInspectRecords(records, options), options.maxChars ?? Infinity);
  }

  designProfile(options = {}) {
    if (!options || typeof options !== "object" || Array.isArray(options)) {
      throw new TypeError("Presentation design profile options must be an object.");
    }
    const state = this[PRESENTATION_STATE];
    const sourceRevisionSha256 = String(state?.opaqueOpc?.sourcePackage?.sha256 || state?.source?.packageSha256 || "").toLowerCase();
    return buildPresentationDesignProfile(this, {
      ...options,
      ...(sourceRevisionSha256 ? { sourceRevisionSha256 } : {}),
    });
  }

  planTemplateGeneration(request = {}) {
    if (!request || typeof request !== "object" || Array.isArray(request)) {
      throw new TypeError("Template generation plan request must be an object.");
    }
    const unsupported = Object.keys(request).filter((key) => !new Set(["slides", "maxItems"]).has(key));
    if (unsupported.length) throw new TypeError(`Template generation plan request has unsupported fields: ${unsupported.join(", ")}.`);
    const profile = this.designProfile({ maxItems: request.maxItems === undefined ? 64 : request.maxItems });
    return buildTemplateGenerationPlan(this, {
      profile,
      slides: request.slides,
      maxItems: request.maxItems === undefined ? 64 : request.maxItems,
    });
  }

  editNativeLeaf(targetId, leafId, update) {
    const capability = this[PRESENTATION_NATIVE_LEAF_CAPABILITY];
    if (!capability) {
      const error = new Error("Presentation native-leaf editing requires a trusted imported PPTX source revision.");
      error.code = "presentation_native_leaf_source_required";
      throw error;
    }
    return capability.edit(targetId, leafId, update);
  }

  editComponentOccurrence(request = {}) {
    if (!request || typeof request !== "object" || Array.isArray(request)) {
      throw new TypeError("Presentation component edit request must be an object.");
    }
    const unsupported = Object.keys(request).filter((key) =>
      !new Set(["candidateId", "occurrenceIndex", "expectedCandidate", "edits"]).has(key));
    if (unsupported.length) {
      throw new TypeError(`Presentation component edit request has unsupported fields: ${unsupported.join(", ")}.`);
    }
    const candidateId = typeof request.candidateId === "string" ? request.candidateId.trim() : "";
    if (!candidateId) throw new TypeError("Presentation component editing requires the exact inspected candidateId.");
    const candidate = this.resolveComponentCandidate(candidateId);
    if (!candidate) {
      const error = new Error(`Presentation component candidate ${candidateId} was not found in this revision.`);
      error.code = "presentation_component_candidate_not_found";
      throw error;
    }
    if (candidate.status !== "inspect-only" || candidate.editCapability?.supported !== true) {
      const error = new Error(`Presentation component candidate ${candidateId} has no bounded native-leaf edit capability.`);
      error.code = "unsupported_presentation_component_edit";
      throw error;
    }
    if (request.expectedCandidate !== undefined) {
      if (!request.expectedCandidate || typeof request.expectedCandidate !== "object" || Array.isArray(request.expectedCandidate)) {
        throw new TypeError("Presentation component edit expectedCandidate must be an inspection object.");
      }
      if (JSON.stringify(request.expectedCandidate) !== JSON.stringify(candidate)) {
        const error = new Error(`Presentation component candidate ${candidateId} ownership evidence is stale.`);
        error.code = "stale_presentation_component_candidate";
        throw error;
      }
    }
    const state = this[PRESENTATION_STATE];
    const sourceRevisionSha256 = String(state?.opaqueOpc?.sourcePackage?.sha256 || state?.source?.packageSha256 || "").toLowerCase();
    if (!sourceRevisionSha256 || sourceRevisionSha256 !== String(candidate.sourceRevisionSha256 || "").toLowerCase()) {
      const error = new Error(`Presentation component candidate ${candidateId} belongs to a different source revision.`);
      error.code = "stale_presentation_source_revision";
      throw error;
    }
    const occurrences = Array.isArray(candidate.occurrences) ? candidate.occurrences : [];
    const occurrenceIndex = request.occurrenceIndex === undefined ? 0 : Number(request.occurrenceIndex);
    if (!Number.isInteger(occurrenceIndex) || occurrenceIndex < 0 || occurrenceIndex >= occurrences.length) {
      throw new RangeError(`Presentation component candidate ${candidateId} occurrenceIndex must identify one inspected occurrence.`);
    }
    const occurrence = occurrences[occurrenceIndex];
    if (occurrence?.editCapability?.supported !== true) {
      const error = new Error(`Presentation component candidate ${candidateId} occurrence ${occurrenceIndex} has no bounded native-leaf edit capability${occurrence?.editCapability?.reason ? `: ${occurrence.editCapability.reason}` : "."}`);
      error.code = "unsupported_presentation_component_edit";
      throw error;
    }
    if (!Array.isArray(request.edits) || request.edits.length === 0 || request.edits.length > 256) {
      throw new TypeError("Presentation component editing requires one through 256 issued leaf edits.");
    }
    const issuedLeafIds = new Set(occurrence.editCapability.leafIds || []);
    for (const edit of request.edits) {
      if (!edit || typeof edit !== "object" || Array.isArray(edit) || !issuedLeafIds.has(edit.leafId)) {
        const error = new Error(`Presentation component edit contains a leaf that was not issued for occurrence ${occurrenceIndex}.`);
        error.code = "presentation_native_leaf_not_issued";
        throw error;
      }
      const editKeys = Object.keys(edit).sort();
      if (editKeys.length !== 4 || editKeys[0] !== "expectedHash" || editKeys[1] !== "leafId" || editKeys[2] !== "targetId" || editKeys[3] !== "value") {
        const error = new Error("Presentation component edits accept exactly targetId, leafId, expectedHash, and value.");
        error.code = "invalid_presentation_native_leaf_edit";
        throw error;
      }
    }
    const nativeCapability = this[PRESENTATION_NATIVE_LEAF_CAPABILITY];
    if (!nativeCapability?.editMany) {
      const error = new Error("Presentation component editing requires the native-leaf compiler.");
      error.code = "presentation_native_leaf_source_required";
      throw error;
    }
    const receipts = nativeCapability.editMany(request.edits.map(({ targetId, leafId, expectedHash, value }) => ({ targetId, leafId, expectedHash, value })));
    return Object.freeze({
      kind: "componentEdit",
      candidateId,
      occurrenceIndex,
      revisionSha256: sourceRevisionSha256,
      edits: receipts,
    });
  }

  resolveComponentCandidate(candidateId) {
    const capability = this[PRESENTATION_COMPONENT_CAPABILITY];
    if (!capability) {
      const error = new Error("Presentation component candidates are available only for a trusted imported PPTX source revision.");
      error.code = "presentation_component_source_required";
      throw error;
    }
    return capability.resolve(candidateId);
  }

  reuseSourceSlide(request = {}) {
    if (!request || typeof request !== "object" || Array.isArray(request)) {
      throw new TypeError("Presentation source-slide reuse request must be an object.");
    }
    const unsupported = Object.keys(request).filter((key) =>
      !new Set(["slideId", "sourceRevisionSha256", "expectedCloneCapability"]).has(key));
    if (unsupported.length) {
      throw new TypeError(`Presentation source-slide reuse request has unsupported fields: ${unsupported.join(", ")}.`);
    }
    const slideId = typeof request.slideId === "string" ? request.slideId.trim() : "";
    if (!slideId) throw new TypeError("Presentation source-slide reuse requires the exact inspected slideId.");
    const sourceRevisionSha256 = typeof request.sourceRevisionSha256 === "string"
      ? request.sourceRevisionSha256.trim().toLowerCase()
      : "";
    if (!/^[0-9a-f]{64}$/u.test(sourceRevisionSha256)) {
      throw new TypeError("Presentation source-slide reuse requires a 64-character sourceRevisionSha256 from inspection.");
    }
    const slide = this.slides.items.find((candidate) => candidate.id === slideId);
    if (!slide) {
      const error = new Error(`Presentation source slide ${slideId} was not found in this revision.`);
      error.code = "presentation_source_slide_not_found";
      throw error;
    }
    const capability = slide.cloneCapability;
    if (!capability.sourceBound || !capability.known || !capability.supported) {
      const error = new Error(`Presentation source slide ${slideId} cannot be reused safely${capability.blockedReason ? `: ${capability.blockedReason}` : "."}`);
      error.code = "unsupported_presentation_slide_clone";
      throw error;
    }
    if (capability.sourceRevisionSha256 !== sourceRevisionSha256) {
      const error = new Error(`Presentation source slide ${slideId} belongs to a different source revision.`);
      error.code = "stale_presentation_source_revision";
      throw error;
    }
    if (request.expectedCloneCapability !== undefined) {
      if (!request.expectedCloneCapability || typeof request.expectedCloneCapability !== "object" || Array.isArray(request.expectedCloneCapability)) {
        throw new TypeError("Presentation source-slide reuse expectedCloneCapability must be an inspection object.");
      }
      const expected = JSON.stringify(request.expectedCloneCapability);
      const actual = JSON.stringify(capability);
      if (expected !== actual) {
        const error = new Error(`Presentation source slide ${slideId} clone ownership evidence is stale.`);
        error.code = "stale_presentation_clone_capability";
        throw error;
      }
    }
    return slide.duplicate();
  }

  reuseSourceComponent(request = {}) {
    if (!request || typeof request !== "object" || Array.isArray(request)) {
      throw new TypeError("Presentation source-component reuse request must be an object.");
    }
    const unsupported = Object.keys(request).filter((key) =>
      !new Set(["candidateId", "occurrenceIndex", "expectedCandidate"]).has(key));
    if (unsupported.length) {
      throw new TypeError(`Presentation source-component reuse request has unsupported fields: ${unsupported.join(", ")}.`);
    }
    const candidateId = typeof request.candidateId === "string" ? request.candidateId.trim() : "";
    if (!candidateId) throw new TypeError("Presentation source-component reuse requires the exact inspected candidateId.");
    const candidate = this.resolveComponentCandidate(candidateId);
    if (!candidate) {
      const error = new Error(`Presentation component candidate ${candidateId} was not found in this revision.`);
      error.code = "presentation_component_candidate_not_found";
      throw error;
    }
    if (candidate.status !== "inspect-only" || candidate.mutationCapability?.supported !== false) {
      const error = new Error(`Presentation component candidate ${candidateId} is not available for bounded source reuse.`);
      error.code = "unsupported_presentation_component_reuse";
      throw error;
    }
    if (request.expectedCandidate !== undefined) {
      if (!request.expectedCandidate || typeof request.expectedCandidate !== "object" || Array.isArray(request.expectedCandidate)) {
        throw new TypeError("Presentation source-component reuse expectedCandidate must be an inspection object.");
      }
      if (JSON.stringify(request.expectedCandidate) !== JSON.stringify(candidate)) {
        const error = new Error(`Presentation component candidate ${candidateId} ownership evidence is stale.`);
        error.code = "stale_presentation_component_candidate";
        throw error;
      }
    }
    const state = this[PRESENTATION_STATE];
    const sourceRevisionSha256 = String(state?.opaqueOpc?.sourcePackage?.sha256 || state?.source?.packageSha256 || "").toLowerCase();
    if (!sourceRevisionSha256 || sourceRevisionSha256 !== String(candidate.sourceRevisionSha256 || "").toLowerCase()) {
      const error = new Error(`Presentation component candidate ${candidateId} belongs to a different source revision.`);
      error.code = "stale_presentation_source_revision";
      throw error;
    }
    const occurrences = Array.isArray(candidate.occurrences) ? candidate.occurrences : [];
    const occurrenceIndex = request.occurrenceIndex === undefined ? 0 : Number(request.occurrenceIndex);
    if (!Number.isInteger(occurrenceIndex) || occurrenceIndex < 0 || occurrenceIndex >= occurrences.length) {
      throw new RangeError(`Presentation component candidate ${candidateId} occurrenceIndex must identify one inspected occurrence.`);
    }
    const occurrence = occurrences[occurrenceIndex];
    if (!occurrence?.slideId || !occurrence.targetId || !Number.isInteger(Number(occurrence.sourceShapeTreeIndex))) {
      const error = new Error(`Presentation component candidate ${candidateId} has no safe top-level source locator.`);
      error.code = "unsupported_presentation_component_reuse";
      throw error;
    }
    if (occurrence.reuseCapability && occurrence.reuseCapability.supported !== true) {
      const error = new Error(`Presentation component candidate ${candidateId} occurrence ${occurrenceIndex} cannot be reused safely${occurrence.reuseCapability.reason ? `: ${occurrence.reuseCapability.reason}` : "."}`);
      error.code = "unsupported_presentation_component_reuse";
      throw error;
    }
    if (occurrence.ownership?.sourceBound !== true || occurrence.ownership?.closedGraph !== true || occurrence.ownership?.mutableDescendantsShared === true) {
      const error = new Error(`Presentation component candidate ${candidateId} is not backed by a closed source graph.`);
      error.code = "unsupported_presentation_component_reuse";
      throw error;
    }
    const sourceState = (state.slides || []).find((entry) => entry.wire?.id === occurrence.slideId);
    const sourceSlide = sourceState?.slide;
    if (!sourceState || !sourceSlide || sourceSlide.presentation !== this) {
      const error = new Error(`Presentation component candidate ${candidateId} source slide is not available in this revision.`);
      error.code = "presentation_component_source_not_found";
      throw error;
    }
    const sourceEntry = sourceState.entries.find((entry) => entry.wire?.id === occurrence.targetId);
    if (!sourceEntry || sourceEntry.model?.parentGroup) {
      const error = new Error(`Presentation component candidate ${candidateId} must identify a direct top-level element.`);
      error.code = "unsupported_presentation_component_reuse";
      throw error;
    }
    const directElements = directSlideModelElements(sourceSlide);
    if (!directElements.includes(sourceEntry.model)) {
      const error = new Error(`Presentation component candidate ${candidateId} does not identify a direct slide element.`);
      error.code = "unsupported_presentation_component_reuse";
      throw error;
    }
    const sourceIds = new Set(sourceState.entries.map((entry) => entry.wire?.id).filter(Boolean));
    const removedSourceIds = new Set(sourceState.entries
      .filter((entry) => entry !== sourceEntry)
      .map((entry) => entry.wire?.id)
      .filter(Boolean));
    for (const entry of sourceState.entries) {
      if (entry === sourceEntry) continue;
      const capability = entry.model?.deletionCapability;
      if (!capability?.sourceBound || capability.known !== true || capability.supported !== true) {
        const error = new Error(`Presentation component candidate ${candidateId} cannot remove source element ${entry.wire?.id || "<unknown>"} safely${capability?.blockedReason ? `: ${capability.blockedReason}` : "."}`);
        error.code = "unsupported_presentation_component_reuse";
        throw error;
      }
    }
    for (const entry of sourceState.entries) {
      if (entry.wire?.content?.case !== "connector" || removedSourceIds.has(entry.wire.id)) continue;
      const connector = entry.model;
      if ([connector.startTargetId, connector.endTargetId].some((targetId) => targetId && removedSourceIds.has(targetId))) {
        const error = new Error(`Presentation component candidate ${candidateId} would leave a retained connector pointing at a removed element.`);
        error.code = "unsupported_presentation_component_reuse";
        throw error;
      }
    }
    const clone = sourceSlide.duplicate();
    const cloneState = (state.clones || []).find((entry) => entry.slide === clone);
    if (!cloneState) {
      throw new Error("Presentation source-component reuse could not establish a clone binding.");
    }
    const cloneTarget = directSlideModelElements(clone).find((element) => cloneState.sourceIdByCloneId?.get(element.id) === occurrence.targetId);
    if (!cloneTarget) {
      const error = new Error(`Presentation component candidate ${candidateId} could not resolve its cloned top-level element.`);
      error.code = "unsupported_presentation_component_reuse";
      throw error;
    }
    const cloneElements = directSlideModelElements(clone);
    cloneState.allowedDeletedIds = new Set();
    cloneState.componentReuse = Object.freeze({ candidateId, occurrenceIndex });
    for (const element of cloneElements) {
      if (element === cloneTarget) continue;
      const sourceId = cloneState.sourceIdByCloneId?.get(element.id);
      if (!sourceId || !sourceIds.has(sourceId)) {
        throw new Error("Presentation source-component reuse encountered an unbound clone element.");
      }
      cloneState.allowedDeletedIds.add(sourceId);
      removePendingCloneDirectElement(clone, element);
    }
    if (cloneState.allowedDeletedIds.size !== sourceState.entries.length - 1) {
      throw new Error("Presentation source-component reuse did not account for every removed source element.");
    }
    return clone;
  }

  validateLayout(options = {}) {
    const issues = this.slides.items.flatMap((slide) => slide.validateLayout(options).issues);
    return { ok: issues.length === 0, issues, ...ndjson(issues, options.maxChars ?? Infinity) };
  }

  auditAccessibility(options = {}) {
    if (!options || typeof options !== "object" || Array.isArray(options)) {
      throw new TypeError("Presentation accessibility audit options must be an object.");
    }
    const records = this.slides.items.flatMap((slide) => presentationSlideElements(slide).map((element) => ({
      slide: slide.index + 1,
      id: element.id,
      name: element.name || undefined,
      kind: presentationElementKind(element),
      nativeKind: element instanceof NativePresentationObject ? element.nativeKind : undefined,
      parentGroupId: element.parentGroup?.id,
      accessibility: element.accessibility ? { ...element.accessibility } : undefined,
    })));
    return auditPresentationAccessibility(records, { ...options, slideCount: this.slides.count });
  }

  verify(options = {}) {
    const issues = [];
    if (this.slides.items.length === 0) issues.push(verificationIssue("presentation", "noSlides", "Presentation has no slides."));
    try { planPresentationCustomShows(this); }
    catch (error) { issues.push(verificationIssue("presentation", "invalidCustomShow", error.message)); }
    try { planPresentationSections(this); }
    catch (error) { issues.push(verificationIssue("presentation", "invalidSection", error.message)); }
    if (this.commentFormat === "modern" || this.slides.items.some((slide) => slide.comments.items.some((thread) => thread.nativeFormat === "modern"))) {
      try { planPresentationModernComments(this.slides.items); }
      catch (error) { issues.push(verificationIssue("presentation", "invalidModernCommentMetadata", error.message)); }
    }
    const duplicateMasterIds = this.masters.items.map((master) => master.id).filter((id, index, ids) => ids.indexOf(id) !== index);
    for (const masterId of new Set(duplicateMasterIds)) issues.push(verificationIssue("presentation", "duplicateMasterId", `Presentation contains duplicate master ID ${masterId}.`, { masterId }));
    const knownMasterIds = new Set(this.masters.items.map((master) => master.id));
    for (const layout of this.layouts.items) if (!knownMasterIds.has(layout.masterId)) issues.push(verificationIssue("presentation", "missingMaster", `Layout ${layout.name || layout.id} references missing master ${layout.masterId}.`, { id: layout.id, masterId: layout.masterId }));
    issues.push(...this.validateLayout(options).issues.map((issue) => ({ ...issue, artifactKind: "presentation" })));
    for (const slide of this.slides) {
      const slideElements = presentationSlideElements(slide);
      if (slide.layoutId && !this.layouts.getItem(slide.layoutId)) issues.push(verificationIssue("presentation", "missingLayout", `Slide ${slide.index + 1} references missing layout ${slide.layoutId}.`, { slide: slide.index + 1, layoutId: slide.layoutId }));
      for (const shape of slideElements.filter((element) => element instanceof Shape)) {
        if (shape.placeholder?.required && !shape.text.value.trim()) issues.push(verificationIssue("presentation", "placeholderMissingContent", `Required ${shape.placeholder.type || "placeholder"} placeholder ${shape.name || shape.id} on slide ${slide.index + 1} is empty.`, { slide: slide.index + 1, id: shape.id, placeholder: shape.placeholder }));
      }
      for (const table of slideElements.filter((element) => element instanceof TableElement)) {
        if (!table.rows || !table.columns || table.values.length === 0 || table.values.every((row) => row.every((cell) => String(cell ?? "").trim() === ""))) issues.push(verificationIssue("presentation", "emptyTable", `Table ${table.name || table.id} on slide ${slide.index + 1} has no visible cell data.`, { slide: slide.index + 1, id: table.id }));
        if (table.values.length !== table.rows) issues.push(verificationIssue("presentation", "tableDataMismatch", `Table ${table.name || table.id} declares ${table.rows} rows but has ${table.values.length} value rows.`, { slide: slide.index + 1, id: table.id, rows: table.rows, valueRows: table.values.length }));
        if (table.values.some((row) => row.length !== table.columns)) issues.push(verificationIssue("presentation", "raggedTableRows", `Table ${table.name || table.id} has rows that do not match its declared column count.`, { slide: slide.index + 1, id: table.id, columns: table.columns, rowLengths: table.values.map((row) => row.length) }));
        try {
          const mergePlan = presentationTableMergePlan(table.rows, table.columns, table._mergeRanges);
          for (const [key, state] of mergePlan.cells) {
            if (state.kind !== "covered") continue;
            const [row, column] = key.split(":").map(Number);
            if (String(table.values[row]?.[column] ?? "") !== "") issues.push(verificationIssue("presentation", "mergedTableCoveredCellContent", `Table ${table.name || table.id} covered cell ${row},${column} must remain empty.`, { slide: slide.index + 1, id: table.id, row, column, mergeOrigin: state.origin }));
          }
        } catch (error) {
          issues.push(verificationIssue("presentation", "invalidTableMerge", `Table ${table.name || table.id} has invalid merge topology: ${error.message}`, { slide: slide.index + 1, id: table.id }));
        }
      }
      for (const chart of slideElements.filter((element) => element instanceof ChartElement)) {
        if (!PRESENTATION_CHART_TYPES.has(String(chart.chartType).toLowerCase())) issues.push(verificationIssue("presentation", "unsupportedChartType", `Chart ${chart.name || chart.id} uses unsupported chart type ${chart.chartType}.`, { severity: "warning", slide: slide.index + 1, id: chart.id, chartType: chart.chartType }));
        if (!chart.series.length) issues.push(verificationIssue("presentation", "emptyChart", `Chart ${chart.name || chart.id} on slide ${slide.index + 1} has no data series.`, { slide: slide.index + 1, id: chart.id }));
        for (const series of chart.series) {
          const values = Array.isArray(series.values) ? series.values : [];
          if (!PRESENTATION_NUMERIC_X_CHART_TYPES.has(chart.chartType) && chart.categories.length && values.length && chart.categories.length !== values.length) issues.push(verificationIssue("presentation", "chartDataMismatch", `Chart ${chart.name || chart.id} series ${series.name || "Series"} has ${values.length} values for ${chart.categories.length} categories.`, { slide: slide.index + 1, id: chart.id, series: series.name, values: values.length, categories: chart.categories.length }));
          if (PRESENTATION_NUMERIC_X_CHART_TYPES.has(chart.chartType) && series.xValues?.length !== values.length) issues.push(verificationIssue("presentation", "chartNumericXMismatch", `Chart ${chart.name || chart.id} series ${series.name || "Series"} must contain one numeric xValue per value.`, { slide: slide.index + 1, id: chart.id, series: series.name, values: values.length, xValues: series.xValues?.length || 0 }));
          if (chart.chartType === "bubble" && (series.bubbleSizes?.length !== values.length || series.bubbleSizes.some((value) => !Number.isFinite(Number(value)) || Number(value) <= 0))) issues.push(verificationIssue("presentation", "chartBubbleSizeMismatch", `Chart ${chart.name || chart.id} series ${series.name || "Series"} must contain one positive bubbleSize per value.`, { slide: slide.index + 1, id: chart.id, series: series.name, values: values.length, bubbleSizes: series.bubbleSizes?.length || 0 }));
          if (values.some((value) => value !== "" && value != null && !Number.isFinite(Number(value)))) issues.push(verificationIssue("presentation", "chartDataNonNumeric", `Chart ${chart.name || chart.id} series ${series.name || "Series"} contains non-numeric values.`, { slide: slide.index + 1, id: chart.id, series: series.name }));
        }
      }
      for (const image of slideElements.filter((element) => element instanceof ImageElement)) {
        if (!image.dataUrl && !image.uri && !image.prompt) issues.push(verificationIssue("presentation", "emptyImage", `Image ${image.name || image.id} on slide ${slide.index + 1} has no dataUrl, uri, or prompt.`, { slide: slide.index + 1, id: image.id }));
        if (image.dataUrl && !imageDataFromDataUrl(image.dataUrl)) issues.push(verificationIssue("presentation", "invalidImageDataUrl", `Image ${image.name || image.id} on slide ${slide.index + 1} has an unsupported data URL.`, { slide: slide.index + 1, id: image.id }));
      }
      for (const object of slideElements.filter((element) => element instanceof NativePresentationObject)) {
        if (!object.rawXml) issues.push(verificationIssue("presentation", "nativeObjectMarkupMissing", `Native ${object.nativeKind} object ${object.name || object.id} on slide ${slide.index + 1} has no preserved markup.`, { slide: slide.index + 1, id: object.id, nativeKind: object.nativeKind }));
        const partPaths = new Set(object.parts.map((part) => part.path));
        const sourcePart = object.sourcePart || `ppt/slides/slide${slide.index + 1}.xml`;
        for (const relationship of object.rootRelationships) {
          if (relationship.targetMode?.toLowerCase() === "external") continue;
          const target = ooxmlSafePartPath(ooxmlResolveRelationshipTarget(sourcePart, relationship.target), "PPTX");
          if (!partPaths.has(target)) issues.push(verificationIssue("presentation", "nativeObjectPartMissing", `Native ${object.nativeKind} object ${object.name || object.id} is missing relationship target ${target}.`, { slide: slide.index + 1, id: object.id, relationshipId: relationship.id, target }));
        }
        for (const part of object.parts) for (const relationship of part.relationships || []) {
          if (relationship.targetMode?.toLowerCase() === "external") continue;
          const target = ooxmlSafePartPath(ooxmlResolveRelationshipTarget(part.path, relationship.target), "PPTX");
          if (!partPaths.has(target)) issues.push(verificationIssue("presentation", "nativeObjectPartMissing", `Native ${object.nativeKind} object ${object.name || object.id} is missing recursive relationship target ${target}.`, { slide: slide.index + 1, id: object.id, sourcePart: part.path, relationshipId: relationship.id, target }));
        }
      }
      for (const comment of slide.comments) {
        if (comment.targetId && !slide.resolve(comment.targetId)) issues.push(verificationIssue("presentation", "danglingComment", `Slide ${slide.index + 1} comment ${comment.id} targets missing element ${comment.targetId}.`, { slide: slide.index + 1, id: comment.id, targetId: comment.targetId }));
      }
    }
    return verificationResult("presentation", issues, options);
  }

  resolve(id) {
    if (id === this.id) return this;
    if (id === this.theme.id) return this.theme;
    const master = this.masters.getItem(id);
    if (master) return master;
    const layout = this.layouts.getItem(id);
    if (layout) return layout;
    const customShow = this.customShows.getItem(id);
    if (customShow) return customShow;
    const section = this.sections.getItem(id);
    if (section) return section;
    for (const slide of this.slides) {
      if (slide.id === id) return slide;
      const found = slide.resolve(id);
      if (found) return found;
    }
    return undefined;
  }

  help(query = "*", options = {}) {
    return ndjson(queryHelpRecords("presentation", query, { ...options, includeInternal: true }), options.maxChars ?? Infinity);
  }

  async export(options = {}) {
    if (options.format === "montage" || options.montage === true) return new FileBlob(presentationMontageSvg(this, options), { type: "image/svg+xml", metadata: { format: "montage", slides: this.slides.count, artifactKind: "presentation" } });
    const slide = options.slide || this.slides.getItem(0) || this.slides.add();
    if (options.format === "layout") return slide.export({ ...options, format: "layout" });
    return slide.export(options);
  }

  toProto() {
    return { id: this.id, slideSize: this.slideSize, theme: this.theme.toJSON(), master: this.master.toJSON(), masters: this.masters.items.map((master) => master.toJSON()), layouts: this.layouts.items.map((layout) => layout.toJSON()), slides: this.slides.items.map((slide) => slide.toProto()), customShows: this.customShows.items.map((show) => show.toJSON()), sections: this.sections.items.map((section) => section.toJSON()), viewProperties: this.view.toProto() };
  }
}

class ShapeCollection {
  constructor(slide, owner) { this.slide = slide; this.owner = owner; this.items = []; }
  add(config = {}) {
    if (config?.geometry === "connector") {
      const connector = connectedPresentationShapeConfig(this.slide, this.owner, config.from, config.to, config, { requireExplicitSites: true });
      return (this.owner?.connectors || this.slide.connectors).add(connector).sendToBack();
    }
    const shape = new Shape(this.slide, config);
    shape.parentGroup = this.owner;
    installPresentationElementOrdering(shape);
    this.items.push(shape);
    if (this.owner) this.owner._rememberChild(shape);
    else this.slide.elements._remember(shape);
    return shape;
  }
  connect(from, to, options = {}) {
    const connector = connectedPresentationShapeConfig(this.slide, this.owner, from, to, options);
    return (this.owner?.connectors || this.slide.connectors).add(connector).sendToBack();
  }
  getConnectionSiteIndex(target, side) { return presentationConnectionSiteIndex(this.slide, this.owner, target, side); }
  getItem(idOrName) { return this.items.find((shape) => shape.id === idOrName || shape.name === idOrName); }
  getItemAt(index) { return this.items[index]; }
  get count() { return this.items.length; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

class ElementCollection {
  constructor(slide, ElementClass, owner) { this.slide = slide; this.ElementClass = ElementClass; this.owner = owner; this.items = []; }
  add(...args) { const element = new this.ElementClass(this.slide, ...args); element.parentGroup = this.owner; installPresentationElementOrdering(element); this.items.push(element); if (this.owner) this.owner._rememberChild(element); else this.slide.elements._remember(element); return element; }
  getItemAt(index) { return this.items[index]; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

class SlideElementCollection {
  constructor(slide) { this.slide = slide; this.items = []; }
  _remember(element) {
    if (element?.slide !== this.slide || element.parentGroup) throw new Error("Direct presentation element must belong to this slide scene stack.");
    if (this.items.includes(element)) throw new Error(`Presentation element ${element.id} is already registered in the slide scene stack.`);
    this.items.push(element);
  }
  getItem(idOrName) { return this.items.find((element) => element.id === idOrName || element.name === idOrName); }
  getItemAt(index) { return this.items[index]; }
  get count() { return this.items.length; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

function normalizeFrame(config = {}, fallback = { left: 0, top: 0, width: 240, height: 160 }) {
  const source = config.position || config.frame || config;
  return {
    left: source.left ?? fallback.left,
    top: source.top ?? fallback.top,
    width: source.width ?? fallback.width,
    height: source.height ?? fallback.height,
  };
}

function resolveAutoLayoutFrame(slide, frame) {
  if (frame === "slide") return slide.frame;
  if (frame?.position) return frame.position;
  if (frame && typeof frame.left === "number" && typeof frame.top === "number" && typeof frame.width === "number" && typeof frame.height === "number") return frame;
  return slide.frame;
}

function elementFrame(element) {
  return element.position || element.frame || element.layoutJson?.().frame;
}

function elementLabel(element) {
  return element.name || element.id;
}

function overlapArea(a, b) {
  const left = Math.max(a.left, b.left);
  const top = Math.max(a.top, b.top);
  const right = Math.min(a.left + a.width, b.left + b.width);
  const bottom = Math.min(a.top + a.height, b.top + b.height);
  return Math.max(0, right - left) * Math.max(0, bottom - top);
}

function coversSlideBackground(frame, slideFrame, minimumCoverage = 0.8) {
  const slideArea = Math.max(0, slideFrame.width) * Math.max(0, slideFrame.height);
  return slideArea > 0 && overlapArea(frame, slideFrame) / slideArea >= minimumCoverage;
}

function containsFrame(container, child, tolerance = 0.5) {
  return container.left <= child.left + tolerance
    && container.top <= child.top + tolerance
    && container.left + container.width >= child.left + child.width - tolerance
    && container.top + container.height >= child.top + child.height - tolerance;
}

function isFilledContainerBackground(element, frame, elements) {
  if (!element || typeof element !== "object" || !frame) return false;
  if (typeof element.text?.value === "string" && element.text.value.trim() !== "") return false;
  const fill = typeof element.fill === "string" ? element.fill.trim().toLowerCase() : element.fill;
  if ((!fill || fill === "transparent" || fill === "none") && !isPresentationGradientFill(fill)) return false;
  if (!["rect", "roundRect", "ellipse"].includes(element.geometry)) return false;
  const containedChildren = elements.filter((candidate) => {
    if (candidate === element) return false;
    const childFrame = elementFrame(candidate);
    return childFrame && containsFrame(frame, childFrame);
  });
  return containedChildren.length >= 1 || isLineLikeFrame(frame);
}

function isLineLikeFrame(frame) {
  return frame.width >= frame.height * 8 || frame.height >= frame.width * 8;
}

function overlapsLineLikeContainer(container, child) {
  if (!isLineLikeFrame(container)) return false;
  if (container.width >= container.height * 8) {
    const childCenter = child.left + child.width / 2;
    return childCenter >= container.left && childCenter <= container.left + container.width;
  }
  const childCenter = child.top + child.height / 2;
  return childCenter >= container.top && childCenter <= container.top + container.height;
}

function isAllowedContainerOverlap(left, leftFrame, right, rightFrame, containers) {
  return (containers.has(left) && (containsFrame(leftFrame, rightFrame) || overlapsLineLikeContainer(leftFrame, rightFrame)))
    || (containers.has(right) && (containsFrame(rightFrame, leftFrame) || overlapsLineLikeContainer(rightFrame, leftFrame)));
}

function textOverflowIssue(slide, element, frame, measurementFrame = frame) {
  const text = element.text?.value || "";
  if (!text) return undefined;
  const textFrame = typeof element.textFrame === "function" ? element.textFrame(measurementFrame) : measurementFrame;
  const displayTextFrame = measurementFrame === frame || typeof element.textFrame !== "function" ? textFrame : element.textFrame(frame);
  const paragraphs = typeof element.text.effectiveParagraphs === "function" ? element.text.effectiveParagraphs() : normalizePresentationParagraphs(text);
  const requiredHeight = paragraphs.reduce((height, paragraph) => {
    const explicitFontSizes = [
      paragraph.style?.fontSize,
      element.text.style.fontSize,
      ...paragraph.runs.map((run) => run.style?.fontSize),
    ].filter((fontSize) => Number.isFinite(fontSize) && fontSize > 0);
    const paragraphFontSize = explicitFontSizes.length ? Math.max(...explicitFontSizes) : 24;
    const availableWidth = Math.max(1, textFrame.width - Math.max(0, paragraph.marginLeft || paragraph.level * 24));
    const charsPerLine = Math.max(1, Math.floor(availableWidth / (paragraphFontSize * 0.55)));
    const requiredLines = presentationParagraphsText([paragraph]).split("\n").reduce((lines, line) => lines + Math.max(1, Math.ceil(line.length / charsPerLine)), 0);
    const spacing = paragraph.lineSpacing || element.text.style.lineSpacing || 1;
    const lineHeight = spacing > 10 ? spacing : paragraphFontSize * spacing;
    return height + (paragraph.spaceBefore ?? paragraphFontSize * (paragraph.spaceBeforePercent || 0)) + requiredLines * lineHeight + (paragraph.spaceAfter ?? paragraphFontSize * (paragraph.spaceAfterPercent || 0));
  }, 0);
  if (requiredHeight <= textFrame.height) return undefined;
  const scaleY = measurementFrame === frame || !(textFrame.height > 0) ? 1 : displayTextFrame.height / textFrame.height;
  const renderedRequiredHeight = requiredHeight * scaleY;
  return {
    kind: "layoutIssue",
    type: "textOverflow",
    severity: "warning",
    slide: slide.index + 1,
    id: element.id,
    name: element.name || undefined,
    bbox: [displayTextFrame.left, displayTextFrame.top, displayTextFrame.width, displayTextFrame.height],
    requiredHeight: Math.round(renderedRequiredHeight),
    message: `Text may overflow ${elementLabel(element)}: estimated ${Math.round(renderedRequiredHeight)}px required for ${Math.round(displayTextFrame.height)}px frame.`,
  };
}

function tableOverflowIssues(slide, tableElement, frame = tableElement.position) {
  const issues = [];
  const cellW = frame.width / Math.max(1, tableElement.columns);
  const cellH = frame.height / Math.max(1, tableElement.rows);
  const fontSize = 13;
  for (let row = 0; row < tableElement.rows; row++) {
    for (let column = 0; column < tableElement.columns; column++) {
      const value = String(tableElement.values[row]?.[column] ?? "");
      const requiredWidth = value.length * fontSize * 0.55 + 12;
      if (requiredWidth > cellW || cellH < fontSize * 1.4) {
        issues.push({
          kind: "layoutIssue",
          type: "tableTextOverflow",
          severity: "warning",
          slide: slide.index + 1,
          id: tableElement.id,
          name: tableElement.name || undefined,
          row,
          column,
          bbox: [frame.left + column * cellW, frame.top + row * cellH, cellW, cellH],
          message: `Table cell ${elementLabel(tableElement)}[${row},${column}] may overflow its cell.`,
        });
      }
    }
  }
  return issues;
}

function normalizedSlideCommentTarget(slide, target) {
  if (target == null) return { targetId: undefined };
  if (typeof target === "string") return { targetId: target };
  if (target.kind === "textRange" || target.kind === "shape" || target.id) return { targetId: target.id };
  if (target.slide) return { targetId: undefined };
  if (target.element) return { targetId: target.element.id };
  if (target.textRange) return { targetId: target.textRange.id };
  if (target.textMatch) {
    const element = target.textMatch.element;
    const query = String(target.textMatch.query ?? "");
    const occurrence = Number(target.textMatch.occurrence ?? 0);
    if (!element?.id || !query || !Number.isInteger(occurrence) || occurrence < 0) throw new TypeError("Comment textMatch requires an element, a non-empty query, and a non-negative integer occurrence.");
    const text = String(element.text?.value ?? element.text ?? "");
    let offset = -1;
    let from = 0;
    for (let index = 0; index <= occurrence; index += 1) {
      offset = text.indexOf(query, from);
      if (offset < 0) throw new RangeError(`Comment textMatch query ${JSON.stringify(query)} occurrence ${occurrence} was not found in ${element.id}.`);
      from = offset + Math.max(1, query.length);
    }
    return {
      targetId: `${element.id}/text`,
      nativeAnchor: { type: "textRange", cp: offset, length: query.length },
    };
  }
  return { targetId: undefined };
}

class SlideCommentThread {
  constructor(slide, target, text, config = {}) {
    const normalizedTarget = normalizedSlideCommentTarget(slide, target);
    this.slide = slide;
    this.id = config.id || aid("pc");
    this.targetId = normalizedTarget.targetId || config.targetId;
    this.author = config.author || "User";
    this.resolved = Boolean(config.resolved);
    this.created = config.created || new Date(0).toISOString();
    this.nativeFormat = config.nativeFormat;
    this.nativeAnchor = config.nativeAnchor || normalizedTarget.nativeAnchor;
    this.position = config.position;
    this.comments = (config.comments || [{ author: this.author, text: String(text ?? ""), created: this.created }]).map((comment) => ({ ...comment, author: comment.author || this.author, text: String(comment.text ?? ""), created: comment.created || this.created }));
  }

  addReply(text, config = {}) {
    this.comments.push({ ...config, author: config.author || this.author, text: String(text ?? ""), created: config.created || new Date(0).toISOString() });
    return this;
  }

  resolve() {
    this.resolved = true;
    if (this.nativeFormat === "modern" || this.comments[0]?.status) this.comments[0].status = "resolved";
    return this;
  }
  reopen() {
    this.resolved = false;
    if (this.nativeFormat === "modern" || this.comments[0]?.status) this.comments[0].status = "active";
    return this;
  }

  inspectRecord() {
    return { kind: "comment", id: this.id, slide: this.slide.index + 1, targetId: this.targetId, author: this.author, resolved: this.resolved, nativeFormat: this.nativeFormat, nativeAnchor: this.nativeAnchor, nativeCommentIds: this.comments.map((comment) => comment.nativeId).filter(Boolean), replies: Math.max(0, this.comments.length - 1), textPreview: this.comments.map((comment) => comment.text).join("\n").slice(0, 300) };
  }

  toJSON() { return { id: this.id, targetId: this.targetId, author: this.author, resolved: this.resolved, created: this.created, nativeFormat: this.nativeFormat, nativeAnchor: this.nativeAnchor, position: this.position, comments: this.comments.map((comment) => ({ ...comment })) }; }
}

class SlideCommentCollection {
  constructor(slide) { this.slide = slide; this.items = []; }
  get capability() {
    const imported = this[PRESENTATION_LEGACY_COMMENTS_CAPABILITY];
    return imported
      ? { ...imported }
      : {
          sourceBound: false,
          format: this.slide.presentation.commentFormat,
          partPresent: this.items.length > 0,
          editable: false,
          addable: this.slide.presentation.commentFormat === "legacy",
        };
  }
  addThread(target, text, config = {}) { const thread = new SlideCommentThread(this.slide, target, text, config); this.items.push(thread); return thread; }
  add(target, text, config = {}) { return this.addThread(target, text, config); }
  getItem(id) { return this.items.find((thread) => thread.id === id); }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

const NativePresentationObject = createNativePresentationObjectClass({ normalizeFrame });

const GroupShape = createPresentationGroupShapeClass({
  createId: aid,
  createShapeCollection: (slide, owner) => new ShapeCollection(slide, owner),
  createConnectorCollection: (slide, owner) => new ElementCollection(slide, ConnectorElement, owner),
  createGroupCollection: (slide, owner, GroupClass) => new ElementCollection(slide, GroupClass, owner),
  createTableCollection: (slide, owner) => new ElementCollection(slide, TableElement, owner),
  createChartCollection: (slide, owner) => new ElementCollection(slide, ChartElement, owner),
  createImageCollection: (slide, owner) => new ElementCollection(slide, ImageElement, owner),
  createNativeObjectCollection: (slide, owner) => new ElementCollection(slide, NativePresentationObject, owner),
  isShape: (element) => element instanceof Shape,
  isConnector: (element) => element instanceof ConnectorElement,
  isGroup: (element) => element instanceof GroupShape,
  isTable: (element) => element instanceof TableElement,
  isChart: (element) => element instanceof ChartElement,
  isImage: (element) => element instanceof ImageElement,
  isNativeObject: (element) => element instanceof NativePresentationObject,
  elementKind: (element) => presentationElementKind(element),
  validateChildLayout: (element, frame) => element instanceof TableElement
    ? tableOverflowIssues(element.slide, element, frame)
    : element instanceof Shape
      ? [textOverflowIssue(element.slide, element, frame, element.position)].filter(Boolean)
      : [],
  createTextRange: (element, id) => createTextRange(element, id, { parentKind: "shape" }),
  textRangeRecord,
  elementLabel,
});
export { GroupShape };
function slideLayoutSlice(slide, layout, options = {}) {
  const targets = inspectTargetTokens(options);
  const search = String(options.search || options.searchTerm || "").trim().toLowerCase();
  if (!targets.length && !search) return layout;
  const before = Math.max(0, Number(options.before ?? options.contextBefore ?? options.context ?? 0) || 0);
  const after = Math.max(0, Number(options.after ?? options.contextAfter ?? options.context ?? 0) || 0);
  const targetsSlide = targets.some((target) => target === slide.id || target === slide.name || target === String(slide.index + 1) || target === "slide");
  if (targetsSlide && !search) return { ...layout, slice: { targets, before, after, matchedElements: layout.elements.length, returnedElements: layout.elements.length } };
  const matches = [];
  layout.elements.forEach((element, index) => {
    const matchesSearch = !search || JSON.stringify(element).toLowerCase().includes(search);
    const matchesTarget = !targets.length || targetsSlide || inspectRecordMatchesTarget(element, targets);
    if (matchesSearch && matchesTarget) matches.push(index);
  });
  const keep = new Set();
  for (const index of matches) {
    for (let i = Math.max(0, index - before); i <= Math.min(layout.elements.length - 1, index + after); i += 1) keep.add(i);
  }
  const elements = layout.elements.filter((_, index) => keep.has(index));
  return { ...layout, elements, slice: { targets, search: search || undefined, before, after, matchedElements: matches.length, returnedElements: elements.length } };
}

class SpeakerNotes {
  constructor(slide, text = "") {
    this.slide = slide;
    this.textFrame = new TextFrame(text);
  }

  get id() { return `${this.slide.id}/notes`; }
  get text() { return this.textFrame.value; }
  set text(value) { this.textFrame.set(value); }
  get capability() {
    const imported = this[PRESENTATION_SPEAKER_NOTES_CAPABILITY];
    return imported
      ? { ...imported }
      : { sourceBound: false, partPresent: Boolean(this.text), editable: true, addable: true };
  }
  setText(value) { this.textFrame.set(value); return this; }
  append(value) { this.textFrame.set(`${this.text}${String(value ?? "")}`); return this; }
  clear() { this.textFrame.set(""); return this; }
}

function orderedSlideModelElements(slide) {
  assertPresentationElementIndexes(slide, slide.elements.items);
  return [...slide.elements.items];
}

function presentationBackgroundHasImage(background) {
  return Boolean(background?.image && (background.image.dataUrl || background.image.assetId));
}

function presentationBackgroundImageSvg(background) {
  if (!background?.image?.dataUrl) return "";
  return `<image href="${xmlEscape(background.image.dataUrl)}" x="0" y="0" width="100%" height="100%" preserveAspectRatio="none"/>`;
}

export class Slide {
  constructor(presentation, options = {}) {
    if (options.hidden !== undefined && typeof options.hidden !== "boolean") {
      throw new TypeError("Presentation slide hidden must be a boolean.");
    }
    this.presentation = presentation;
    this.id = aid("sl");
    this.name = options.name || "";
    this.elements = new SlideElementCollection(this);
    this.shapes = new ShapeCollection(this);
    this.images = new ElementCollection(this, ImageElement);
    this.tables = new ElementCollection(this, TableElement);
    this.charts = new ElementCollection(this, ChartElement);
    this.connectors = new ElementCollection(this, ConnectorElement);
    this.groups = new ElementCollection(this, GroupShape);
    this.nativeObjects = new ElementCollection(this, NativePresentationObject);
    this.comments = new SlideCommentCollection(this);
    this.layoutId = options.layoutId || options.layout?.id || (typeof options.layout === "string" ? options.layout : undefined);
    this.speakerNotes = new SpeakerNotes(this, options.notes || options.speakerNotes?.text || "");
    this.background = options.background ? normalizePresentationBackground(options.background) : {};
    this.transition = new SlideTransition(this, options.transition);
    this.animations = new SlideAnimations(this, options.animations || []);
    this.morph = new SlideMorph(this, options.morph);
    this._hidden = options.hidden ?? false;
  }

  get index() { return this.presentation.slides.items.indexOf(this); }
  get hidden() { return this._hidden; }
  get visibilityCapability() {
    const imported = this[PRESENTATION_SLIDE_VISIBILITY_CAPABILITY];
    return imported ? { ...imported } : { sourceBound: false, known: true, editable: true };
  }
  get deletionCapability() {
    const imported = this[PRESENTATION_SLIDE_DELETION_CAPABILITY];
    return imported
      ? { ...imported }
      : { sourceBound: false, known: true, supported: true, blockedReason: "", ownedPartCount: 0 };
  }
  get cloneCapability() {
    const imported = this[PRESENTATION_SLIDE_CLONE_CAPABILITY];
    return imported
      ? { ...imported }
      : { sourceBound: false, known: true, supported: false, blockedReason: "Source-free slides use ordinary authoring rather than source-preserving graph clone.", clonedPartCount: 0, sharedPartCount: 0 };
  }
  get continuationCapability() {
    const imported = this[PRESENTATION_SLIDE_CONTINUATION_CAPABILITY];
    return imported
      ? { ...imported }
      : { sourceBound: false, ready: true, profile: "full-authoring", requiresExportReopen: false };
  }
  setHidden(hidden) {
    if (typeof hidden !== "boolean") throw new TypeError("Presentation slide hidden must be a boolean.");
    const capability = this.visibilityCapability;
    if (capability.sourceBound && (!capability.known || !capability.editable)) {
      throw new Error("Imported presentation slide visibility is source-bound and not safely editable.");
    }
    this._hidden = hidden;
    return this;
  }
  hide() { return this.setHidden(true); }
  show() { return this.setHidden(false); }
  moveTo(index) {
    if (!Number.isInteger(index) || index < 0 || index >= this.presentation.slides.items.length) {
      throw new RangeError("Presentation slide destination must be an existing 0-based slide index.");
    }
    const current = this.index;
    if (current < 0) throw new Error("Presentation slide must belong to its presentation before it can move.");
    if (current === index) return this;
    this.presentation.slides.items.splice(current, 1);
    this.presentation.slides.items.splice(index, 0, this);
    return this;
  }
  duplicate() {
    const duplicate = this.presentation[PRESENTATION_SLIDE_DUPLICATOR];
    if (typeof duplicate !== "function") {
      throw new Error("Presentation slide duplication is available only for a supported imported PPTX source slide.");
    }
    return duplicate(this);
  }
  delete() {
    const current = this.index;
    if (current < 0) throw new Error("Presentation slide must belong to its presentation before it can be deleted.");
    if (this.presentation.slides.items.length <= 1) {
      throw new RangeError("Presentation must retain at least one slide.");
    }
    const capability = this.deletionCapability;
    if (capability.sourceBound && (!capability.known || !capability.supported)) {
      const detail = capability.blockedReason ? `: ${capability.blockedReason}` : ".";
      const error = new Error(`Imported presentation slide cannot be safely deleted${detail}`);
      error.code = "unsupported_presentation_slide_delete";
      throw error;
    }
    this.presentation.slides.items.splice(current, 1);
  }
  get frame() { return { left: 0, top: 0, ...this.presentation.slideSize }; }
  get placeholders() {
    const items = this.shapes.items.filter((shape) => shape.placeholder);
    return {
      items,
      get count() { return items.length; },
      getItem(idOrName) {
        const key = String(idOrName ?? "");
        const index = Number(idOrName);
        return items.find((shape) => shape.id === idOrName || shape.name === idOrName ||
          shape.placeholder?.type === key || shape.placeholder?.name === key ||
          (Number.isInteger(index) && Number(shape.placeholder?.idx) === index));
      },
      [Symbol.iterator]() { return items[Symbol.iterator](); },
    };
  }

  addNotes(text) { return this.speakerNotes.setText(text); }
  addComment(target, text, config = {}) { return this.comments.addThread(target, text, config); }
  addConnector(config = {}) { return this.connectors.add(config); }
  addGroup(config = {}) { return this.groups.add(config); }
  setBackgroundImage(config = {}) {
    if (!config || typeof config !== "object" || Array.isArray(config)) throw new TypeError("Presentation background image requires an options object.");
    if (this.elements.items.some((element) => element.zOrderCapability?.sourceBound === true)) {
      const error = new Error("Imported presentation slides cannot place a new image below source-bound elements; use a source-derived slide or a capability-proven native reorder.");
      error.code = "unsupported_presentation_background_image";
      throw error;
    }
    const existing = this.images.items.find((image) => image._officeKitLayerRole === "background");
    const frame = { ...this.frame };
    if (existing) {
      existing.replace({ ...config, position: frame });
      existing.position = frame;
      existing.sendToBack();
      return existing;
    }
    const image = this.images.add({ ...config, position: frame, fit: config.fit || "cover" });
    Object.defineProperty(image, "_officeKitLayerRole", { enumerable: false, configurable: false, writable: false, value: "background" });
    image.sendToBack();
    return image;
  }
  setNativeBackgroundImage(config = {}) {
    this.background = normalizeNativePresentationBackgroundImage(config, `Presentation slide ${this.id}`, this.background);
    return this;
  }
  clearNativeBackgroundImage() {
    if (presentationBackgroundHasImage(this.background)) this.background = {};
    return this;
  }
  clearBackgroundImage() {
    const existing = this.images.items.find((image) => image._officeKitLayerRole === "background");
    if (existing) existing.delete();
    return this;
  }
  setBackground(background) { this.background = normalizePresentationBackground(background, this.background); return this; }
  clearBackground() { this.background = {}; return this; }
  setTransition(transition) { this.transition.set(transition); return this; }
  clearTransition() { this.transition.clear(); return this; }
  setMorph(morph) {
    if (morph == null) this.morph.clear();
    else this.morph.set(morph);
    return this;
  }
  clearMorph() { this.morph.clear(); return this; }
  applyLayout(layoutOrName) {
    const layout = typeof layoutOrName === "string" ? this.presentation.layouts.getItem(layoutOrName) : layoutOrName;
    if (!(layout instanceof SlideLayoutTemplate) || layout.presentation !== this.presentation) {
      throw new Error(`Unknown slide layout: ${typeof layoutOrName === "string" ? layoutOrName : "provided layout"}`);
    }
    return layout.apply(this);
  }
  setLayout(layoutOrName) { this.applyLayout(layoutOrName); return this; }
  effectiveBackground() { const layout = this.presentation.layouts.getItem(this.layoutId); return this.background.fill || this.background.gradient ? this.background : layout?.effectiveBackground() || this.presentation.master.effectiveBackground(); }
  effectiveBackgroundImage() {
    if (presentationBackgroundHasImage(this.background)) return this.background;
    if (this.background?.fill || this.background?.gradient) return undefined;
    const layout = this.presentation.layouts.getItem(this.layoutId);
    return layout?.effectiveBackgroundImage() || this.presentation.master.effectiveBackgroundImage();
  }
  effectiveTheme() { const layout = this.presentation.layouts.getItem(this.layoutId); return layout?.effectiveTheme() || this.presentation.master.effectiveTheme(); }

  inspectRecords(kinds) {
    const records = [];
    if (kinds.has("layout")) { const layout = this.presentation.layouts.getItem(this.layoutId); records.push({ kind: "layout", layoutId: this.layoutId || `${this.id}/layout`, name: layout?.name || "Blank", type: layout?.type || "blank", masterId: layout?.masterId, themeId: this.effectiveTheme().id, placeholders: layout?.placeholders.length || 0 }); }
    if (kinds.has("slide")) { const directImage = presentationBackgroundHasImage(this.background); const effectiveImage = this.effectiveBackgroundImage(); const layout = this.presentation.layouts.getItem(this.layoutId); const inheritedImage = !directImage && !this.background?.fill && !this.background?.gradient && Boolean(effectiveImage); const imageOwner = directImage ? "slide" : layout && presentationBackgroundHasImage(layout.background) ? "layout" : inheritedImage ? "master" : undefined; records.push({ kind: "slide", id: this.id, slide: this.index + 1, title: this.title(), hidden: this.hidden, visibilityCapability: this.visibilityCapability, deletionCapability: this.deletionCapability, cloneCapability: this.cloneCapability, continuationCapability: this.continuationCapability, background: this.background.fill || this.background.gradient || this.background.image ? this.background : undefined, nativeBackgroundImage: effectiveImage ? { fit: "stretch", editable: directImage, inherited: !directImage, owner: imageOwner } : undefined, effectiveBackground: this.effectiveBackground(), transition: this.transition.toJSON(), transitionCapability: this.transition.capability, textShapes: this.shapes.items.filter((s) => s.text.value).length, tables: this.tables.items.length, charts: this.charts.items.length, images: this.images.items.length, connectors: this.connectors.items.length, groups: this.groups.items.length, nativeObjects: this.nativeObjects.items.length, layerCount: this.elements.count, comments: this.comments.items.length, commentsCapability: this.comments.capability, hasNotes: Boolean(this.speakerNotes.text), notesCapability: this.speakerNotes.capability }); }
    if (kinds.has("layer") || kinds.has("zOrder")) records.push(...this.elements.items.map((element, stackIndex) => ({
      kind: "layer",
      id: element.id,
      slide: this.index + 1,
      elementKind: element instanceof GroupShape ? "groupShape" : element.kind || element.constructor?.name,
      name: element.name || undefined,
      stackIndex,
      layerRole: element._officeKitLayerRole,
      zOrderCapability: element.zOrderCapability,
    })));
    for (const shape of this.shapes) {
      if (kinds.has("textbox") && shape.text.value) records.push(shape.inspectRecord("textbox"));
      else if (kinds.has("shape")) records.push(shape.inspectRecord("shape"));
      if (kinds.has("textRange") && shape.text.value) records.push(textRangeRecord(shape, { parentKind: "shape", record: { slide: this.index + 1, bbox: [shape.position.left, shape.position.top, shape.position.width, shape.position.height], bboxUnit: "px" } }));
    }
    if (kinds.has("table")) records.push(...this.tables.items.map((table) => table.inspectRecord()));
    if (kinds.has("chart")) records.push(...this.charts.items.map((chart) => chart.inspectRecord()));
    if (kinds.has("image")) records.push(...this.images.items.map((image) => image.inspectRecord()));
    if (kinds.has("connector")) records.push(...this.connectors.items.map((connector) => connector.inspectRecord()));
    if (kinds.has("nativeObject") || kinds.has("native")) records.push(...this.nativeObjects.items.map((object) => object.inspectRecord()));
    for (const nativeKind of ["contentPart", "oleObject", "diagram", "graphicFrame"]) if (kinds.has(nativeKind)) records.push(...this.nativeObjects.items.filter((object) => object.nativeKind === nativeKind).map((object) => object.inspectRecord()));
    for (const group of this.groups) records.push(...group.inspectRecords(kinds));
    if (kinds.has("comment") || kinds.has("thread")) records.push(...this.comments.items.map((comment) => comment.inspectRecord()));
    if (kinds.has("notes")) records.push({ kind: "notes", id: `${this.id}/notes`, slide: this.index + 1, text: this.speakerNotes.text, textPreview: this.speakerNotes.text.slice(0, 300), textChars: this.speakerNotes.text.length, capability: this.speakerNotes.capability });
    if (kinds.has("transition")) records.push(this.transition.inspectRecord());
    if (kinds.has("animations")) records.push(this.animations.inspectRecord());
    if (kinds.has("animation")) records.push(...this.animations.inspectRecords());
    if (kinds.has("morph")) records.push(...this.morph.inspectRecords());
    return records;
  }

  title() { return this.shapes.items.find((shape) => shape.text.value)?.text.value || this.charts.items[0]?.title || ""; }
  resolve(id) {
    if (id === this.speakerNotes.id) return this.speakerNotes;
    if (id === this.transition.id) return this.transition;
    if (id === this.morph.id) return this.morph;
    if (id === `${this.id}/animations`) return this.animations;
    if (String(id || "").endsWith("/text")) {
      const parentId = String(id).slice(0, -5);
      const shape = this.shapes.items.find((item) => item.id === parentId);
      if (shape) return createTextRange(shape, id, { parentKind: "shape" });
    }
    const direct = [...this.elements.items, ...this.comments.items].find((element) => element.id === id);
    if (direct) return direct;
    for (const group of this.groups) {
      const nested = group.resolve(id);
      if (nested) return nested;
    }
    return undefined;
  }

  validateLayout(options = {}) {
    const issues = [];
    const slideFrame = this.frame;
    const elements = this.elements.items.filter((element) => !(element instanceof ConnectorElement));
    const connectors = this.connectors.items;
    const minOverlapArea = options.minOverlapArea ?? 64;
    const backgroundCoverage = options.backgroundCoverage ?? 0.8;
    const padding = options.boundsPadding ?? 0;
    const backgroundElements = new Set(elements.filter((element) => {
      const frame = elementFrame(element);
      return frame && coversSlideBackground(frame, slideFrame, backgroundCoverage);
    }));
    const containerBackgrounds = new Set(elements.filter((element) => {
      const frame = elementFrame(element);
      return isFilledContainerBackground(element, frame, elements);
    }));
    for (const element of elements) {
      const frame = elementFrame(element);
      if (!frame) continue;
      const offCanvas = frame.left < slideFrame.left - padding || frame.top < slideFrame.top - padding || frame.left + frame.width > slideFrame.left + slideFrame.width + padding || frame.top + frame.height > slideFrame.top + slideFrame.height + padding;
      if (offCanvas) {
        issues.push({
          kind: "layoutIssue",
          type: "offCanvas",
          severity: "error",
          slide: this.index + 1,
          id: element.id,
          name: element.name || undefined,
          bbox: [frame.left, frame.top, frame.width, frame.height],
          message: `${elementLabel(element)} extends outside the slide frame.`,
        });
      }
      const textIssue = textOverflowIssue(this, element, frame);
      if (textIssue) issues.push(textIssue);
      if (element instanceof TableElement) issues.push(...tableOverflowIssues(this, element));
    }
    for (const connector of connectors) {
      const points = [connector.start, connector.end];
      if (points.some((point) => point.x < slideFrame.left - padding || point.y < slideFrame.top - padding || point.x > slideFrame.left + slideFrame.width + padding || point.y > slideFrame.top + slideFrame.height + padding)) {
        issues.push({ kind: "layoutIssue", type: "connectorOffCanvas", severity: "error", slide: this.index + 1, id: connector.id, name: connector.name || undefined, start: connector.start, end: connector.end, message: `${elementLabel(connector)} connector endpoint extends outside the slide frame.` });
      }
    }
    for (const group of this.groups) issues.push(...group.validateLayout());
    for (let leftIndex = 0; leftIndex < elements.length; leftIndex++) {
      for (let rightIndex = leftIndex + 1; rightIndex < elements.length; rightIndex++) {
        const left = elements[leftIndex];
        const right = elements[rightIndex];
        if (backgroundElements.has(left) || backgroundElements.has(right)) continue;
        const leftFrame = elementFrame(left);
        const rightFrame = elementFrame(right);
        if (!leftFrame || !rightFrame) continue;
        if (isAllowedContainerOverlap(left, leftFrame, right, rightFrame, containerBackgrounds)) continue;
        const area = overlapArea(leftFrame, rightFrame);
        if (area >= minOverlapArea) {
          issues.push({
            kind: "layoutIssue",
            type: "overlap",
            severity: "error",
            slide: this.index + 1,
            ids: [left.id, right.id],
            names: [elementLabel(left), elementLabel(right)],
            overlapArea: Math.round(area),
            message: `${elementLabel(left)} overlaps ${elementLabel(right)} by about ${Math.round(area)}px².`,
          });
        }
      }
    }
    return { ok: issues.length === 0, issues, ...ndjson(issues, options.maxChars ?? Infinity) };
  }

  async export(options = {}) {
    if (options.format === "layout" || options.format === LAYOUT_MIME) return new FileBlob(JSON.stringify(this.layoutJson(options), null, 2), { type: LAYOUT_MIME, metadata: { artifactKind: "presentation", format: "layout", slide: this.index + 1, target: options.target ?? options.targetId ?? options.id ?? options.anchor, search: options.search ?? options.searchTerm } });
    return new FileBlob(this.toSvg(), { type: "image/svg+xml" });
  }

  layoutJson(options = {}) {
    const elements = orderedSlideModelElements(this).map((element) => {
      const record = element.layoutJson();
      const comments = this.comments.items.filter((comment) => comment.targetId === element.id);
      return {
        ...record,
        slide: this.index + 1,
        textRangeId: element.text?.value ? `${element.id}/text` : undefined,
        commentIds: comments.length ? comments.map((comment) => comment.id) : undefined,
        commentTextPreview: comments.length ? comments.flatMap((comment) => comment.comments.map((item) => item.text)).join("\n").slice(0, 300) : undefined,
      };
    });
    const effectiveBackground = this.effectiveBackgroundImage() || this.effectiveBackground();
    return slideLayoutSlice(this, {
      schema: "office-kit-artifact.layout/v1",
      unit: "px",
      slide: { id: this.id, slide: this.index + 1, frame: this.frame, hidden: this.hidden, background: effectiveBackground, transition: this.transition.toJSON(), animations: this.animations.toJSON(), morph: this.morph.toJSON(), notes: this.speakerNotes.text || undefined },
      elements,
    }, options);
  }

  toSvg() {
    const { width, height } = this.presentation.slideSize;
    const elements = orderedSlideModelElements(this).map((element) => element.toSvg()).join("");
    const backgroundImage = this.effectiveBackgroundImage();
    const imageSvg = presentationBackgroundImageSvg(backgroundImage);
    const effectiveBackground = this.effectiveBackground();
    const backgroundGradient = effectiveBackground?.gradient
      ? presentationGradientFillSvg(effectiveBackground.gradient, `${this.id}-background`, `Presentation slide ${this.id} background gradient`)
      : undefined;
    const background = imageSvg
      ? imageSvg
      : backgroundGradient
        ? `${backgroundGradient.defs}<rect width="100%" height="100%" fill="${xmlEscape(backgroundGradient.paint)}"/>`
        : `<rect width="100%" height="100%" fill="${xmlEscape(resolvePresentationBackgroundColor(effectiveBackground, this.effectiveTheme()))}"/>`;
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${background}${elements}</svg>`;
  }

  toProto() { return { id: this.id, layoutId: this.layoutId, hidden: this.hidden, background: this.background.fill || this.background.gradient || this.background.image ? this.background : undefined, transition: this.transition.toJSON(), animations: this.animations.toJSON(), morph: this.morph.toJSON(), notes: this.speakerNotes.text || undefined, comments: this.comments.items.map((comment) => comment.toJSON()), elements: orderedSlideModelElements(this).filter((element) => !(element instanceof GroupShape)).map((element) => element.layoutJson()), groups: this.groups.items.map((group) => group.toProto()) }; }

  compose(composeNode, options = {}) {
    const frame = options.frame || { left: 72, top: 64, width: this.presentation.slideSize.width - 144, height: this.presentation.slideSize.height - 128 };
    return materializeComposeNode(this, composeNode, frame);
  }

  autoLayout(shapes, options = {}) {
    const items = Array.from(shapes || []).filter(Boolean);
    if (items.length === 0) return items;
    const frame = resolveAutoLayoutFrame(this, options.frame || "slide");
    const inner = innerFrame(frame, {
      left: options.horizontalPadding ?? 0,
      right: options.horizontalPadding ?? 0,
      top: options.verticalPadding ?? 0,
      bottom: options.verticalPadding ?? 0,
    });
    const direction = options.direction || "horizontal";
    const horizontal = direction === "horizontal";
    const mainSize = horizontal ? "width" : "height";
    const crossSize = horizontal ? "height" : "width";
    const requestedGap = horizontal ? options.horizontalGap : options.verticalGap;
    const totalMain = items.reduce((sum, shape) => sum + (shape.position?.[mainSize] ?? 0), 0);
    const gap = requestedGap === "auto"
      ? items.length > 1 ? Math.max(0, (inner[mainSize] - totalMain) / (items.length - 1)) : 0
      : Number(requestedGap ?? 0);
    const usedMain = totalMain + gap * Math.max(0, items.length - 1);
    const align = options.align || "center";
    const mainStart = align.includes("Right") || align === "right" || align.includes("Bottom")
      ? inner[horizontal ? "left" : "top"] + inner[mainSize] - usedMain
      : align === "center" || align === "left" || align === "right"
        ? inner[horizontal ? "left" : "top"] + Math.max(0, (inner[mainSize] - usedMain) / 2)
        : inner[horizontal ? "left" : "top"];
    let cursor = mainStart;
    for (const shape of items) {
      const crossStart = align.includes("Bottom")
        ? inner[horizontal ? "top" : "left"] + inner[crossSize] - shape.position[crossSize]
        : align.includes("Center") || align === "center" || align === "left" || align === "right"
          ? inner[horizontal ? "top" : "left"] + Math.max(0, (inner[crossSize] - shape.position[crossSize]) / 2)
          : inner[horizontal ? "top" : "left"];
      shape.position = horizontal
        ? { ...shape.position, left: cursor, top: crossStart }
        : { ...shape.position, left: crossStart, top: cursor };
      cursor += shape.position[mainSize] + gap;
    }
    return items;
  }
}

class TextFrame {
  constructor(text = "", bodyProperties, { defaultBodyProperties = false } = {}) { this._paragraphs = normalizePresentationParagraphs(text); this.style = {}; this.inheritedParagraphStyles = {}; this.bodyProperties = normalizePresentationTextBodyProperties(bodyProperties, { defaults: defaultBodyProperties }); }
  get value() { return presentationParagraphsText(this._paragraphs); }
  set value(text) { this._paragraphs = normalizePresentationParagraphs(text); }
  get paragraphs() { return normalizePresentationParagraphs(this._paragraphs); }
  set paragraphs(value) { this._paragraphs = normalizePresentationParagraphs(value); }
  effectiveParagraphs() { return inheritPresentationParagraphs(this._paragraphs, this.inheritedParagraphStyles); }
  set(text) { this._paragraphs = normalizePresentationParagraphs(text); return this; }
  setText(text) { return this.set(text); }
  replace(search, replacement) { replacePresentationParagraphText(this._paragraphs, search, replacement); return this; }
  toString() { return this.value; }
}

function normalizePresentationShapeFill(fill, label) {
  if (typeof fill === "string") return fill;
  if (!fill || typeof fill !== "object" || Array.isArray(fill)) throw new TypeError(`${label} must be a color string or fill object.`);
  if (fill.type === "gradient") return normalizePresentationGradientFill(fill, label);
  const color = fill.color ?? fill.fill;
  if (typeof color !== "string" || color.length === 0) throw new TypeError(`${label}.color must be a non-empty color string.`);
  if (fill.opacity != null) {
    const opacity = Number(fill.opacity);
    if (!Number.isFinite(opacity) || opacity < 0 || opacity > 1) throw new RangeError(`${label}.opacity must be from 0 through 1.`);
  }
  return { ...fill, color };
}

export class Shape {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("sh");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
    this.name = config.name || "";
    this.position = config.position || { left: 0, top: 0, width: 160, height: 80 };
    this.geometry = config.geometry || "rect";
    const formulaGraph = normalizePresentationCustomGeometryFormulaGraph({ adjustments: config.customAdjustments, guides: config.customGuides });
    const customGeometryContext = {
      ...formulaGraph,
      widthEmu: Math.round(Number(this.position.width) * EMU_PER_PIXEL),
      heightEmu: Math.round(Number(this.position.height) * EMU_PER_PIXEL),
    };
    this.customAdjustments = formulaGraph.adjustments;
    this.customGuides = formulaGraph.guides;
    this.customPaths = normalizePresentationCustomPaths(config.customPaths, customGeometryContext);
    this.customConnectionSites = normalizePresentationCustomConnectionSites(config.customConnectionSites, customGeometryContext);
    this.customAdjustmentHandles = normalizePresentationCustomAdjustmentHandles(config.customAdjustmentHandles, customGeometryContext);
    this.textRectangle = normalizePresentationCustomTextRectangle(config.textRectangle, customGeometryContext);
    if (this.geometry !== "custom" && (this.customPaths.length || this.customConnectionSites.length || this.customAdjustmentHandles.length || this.customAdjustments.length || this.customGuides.length || this.textRectangle)) {
      throw new TypeError("Presentation customPaths, customConnectionSites, customAdjustmentHandles, customAdjustments, customGuides, and textRectangle are available only for custom geometry shapes.");
    }
    this.transform = config.transform == null ? undefined : normalizePresentationPlaceholderTransform(config.transform, `Presentation shape ${this.name || this.id} transform`);
    this.fill = normalizePresentationShapeFill(config.fill || "transparent", `Presentation shape ${this.name || this.id} fill`);
    this.line = config.line || (this.geometry === "textbox"
      ? { fill: "transparent", width: 0 }
      : { fill: "#334155", width: 1 });
    this.borderRadius = config.borderRadius;
    this.shadow = config.shadow ? { ...config.shadow } : undefined;
    this.placeholder = config.placeholder;
    this.accessibility = initializePresentationAccessibility(this, config, `Presentation shape ${this.id}`);
    if (config._officeKitUseBackgroundFill !== undefined) importedShapeBackgroundFill.set(this, Boolean(config._officeKitUseBackgroundFill));
    this._text = new TextFrame(config.text ?? "", config.textBodyProperties, { defaultBodyProperties: config.textBodyProperties === undefined });
    this._text.style = { ...(config.textStyle || config.style?.text || {}) };
  }

  get text() { return this._text; }
  set text(value) { this._text.set(value); }
  get useBackgroundFill() { return importedShapeBackgroundFill.get(this); }
  get accessibilityCapability() { return presentationAccessibilityCapability(this); }
  get deletionCapability() { return presentationElementDeletionCapability(this, "shape"); }

  delete() {
    const owner = this.parentGroup;
    const collection = owner?.shapes || this.slide?.shapes;
    return deletePresentationElement(this, collection, "shape");
  }

  setAccessibilityMetadata(update) {
    this.accessibility = setPresentationAccessibilityMetadata(this, this.accessibility, update, `Presentation shape ${this.id}`);
    return this;
  }

  #normalizedTextRectangle(formulaGraph = normalizePresentationCustomGeometryFormulaGraph({ adjustments: this.customAdjustments, guides: this.customGuides })) {
    const rectangle = normalizePresentationCustomTextRectangle(this.textRectangle, {
      ...formulaGraph,
      widthEmu: Math.round(Number(this.position?.width) * EMU_PER_PIXEL),
      heightEmu: Math.round(Number(this.position?.height) * EMU_PER_PIXEL),
    });
    if (rectangle && this.geometry !== "custom") throw new TypeError("Presentation textRectangle is available only for custom geometry shapes.");
    return rectangle;
  }

  #normalizedCustomGeometry() {
    const graph = normalizePresentationCustomGeometryFormulaGraph({ adjustments: this.customAdjustments, guides: this.customGuides });
    const paths = this.customPaths?.length ? normalizePresentationCustomPaths(this.customPaths, {
      ...graph,
      widthEmu: Math.round(Number(this.position?.width) * EMU_PER_PIXEL),
      heightEmu: Math.round(Number(this.position?.height) * EMU_PER_PIXEL),
    }) : [];
    const connectionSites = normalizePresentationCustomConnectionSites(this.customConnectionSites, {
      ...graph,
      widthEmu: Math.round(Number(this.position?.width) * EMU_PER_PIXEL),
      heightEmu: Math.round(Number(this.position?.height) * EMU_PER_PIXEL),
    });
    const adjustmentHandles = normalizePresentationCustomAdjustmentHandles(this.customAdjustmentHandles, {
      ...graph,
      widthEmu: Math.round(Number(this.position?.width) * EMU_PER_PIXEL),
      heightEmu: Math.round(Number(this.position?.height) * EMU_PER_PIXEL),
    });
    const textRectangle = this.#normalizedTextRectangle(graph);
    if (this.geometry !== "custom" && (paths.length || connectionSites.length || adjustmentHandles.length || graph.adjustments.length || graph.guides.length || textRectangle)) {
      throw new TypeError("Presentation custom paths, connectionSites, adjustmentHandles, adjustments, guides, and textRectangle are available only for custom geometry shapes.");
    }
    return { ...graph, paths, connectionSites, adjustmentHandles, textRectangle };
  }

  inspectRecord(kind = "shape") {
    const p = this.position;
    const paragraphs = this.text.effectiveParagraphs();
    const custom = this.#normalizedCustomGeometry();
    return { kind, id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, text: this.text.value || undefined, textPreview: this.text.value || undefined, textChars: this.text.value.length || undefined, textLines: this.text.value ? this.text.value.split("\n").length : undefined, paragraphs: presentationParagraphsNeedSerialization(paragraphs) ? paragraphs : undefined, bodyProperties: this.text.bodyProperties, customPathCount: custom.paths.length || undefined, customAdjustmentCount: custom.adjustments.length || undefined, customGuideCount: custom.guides.length || undefined, customConnectionSiteCount: custom.connectionSites.length || undefined, customAdjustmentHandleCount: custom.adjustmentHandles.length || undefined, customConnectionSites: custom.connectionSites.length ? custom.connectionSites : undefined, customAdjustmentHandles: custom.adjustmentHandles.length ? custom.adjustmentHandles : undefined, textRectangle: custom.textRectangle, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", transform: this.transform, shadow: this.shadow, placeholder: this.placeholder || undefined, accessibility: this.accessibility ? { ...this.accessibility } : undefined, accessibilityCapability: this.accessibilityCapability, deletionCapability: this.deletionCapability, useBackgroundFill: this.useBackgroundFill };
  }

  layoutJson() { const paragraphs = this.text.effectiveParagraphs(); const custom = this.#normalizedCustomGeometry(); return { kind: this.text.value ? "textbox" : "shape", id: this.id, name: this.name, geometry: this.geometry, customAdjustments: custom.adjustments.length ? custom.adjustments : undefined, customGuides: custom.guides.length ? custom.guides : undefined, customConnectionSites: custom.connectionSites.length ? custom.connectionSites : undefined, customAdjustmentHandles: custom.adjustmentHandles.length ? custom.adjustmentHandles : undefined, customPaths: custom.paths.length ? custom.paths : undefined, textRectangle: custom.textRectangle, frame: this.position, transform: this.transform, text: this.text.value, paragraphs: presentationParagraphsNeedSerialization(paragraphs) ? paragraphs : undefined, bodyProperties: this.text.bodyProperties, placeholder: this.placeholder, accessibility: this.accessibility ? { ...this.accessibility } : undefined, style: { fill: this.fill, line: this.line, borderRadius: this.borderRadius, shadow: this.shadow, text: this.text.style, useBackgroundFill: this.useBackgroundFill } }; }

  textFrame(frame = this.position) {
    const graph = normalizePresentationCustomGeometryFormulaGraph({ adjustments: this.customAdjustments, guides: this.customGuides });
    return presentationCustomTextRectangleFrame(this.#normalizedTextRectangle(graph), frame, this.position, graph);
  }

  toSvg() {
    const p = this.position;
    const custom = this.#normalizedCustomGeometry();
    const textFrame = this.textFrame(p);
    const gradient = isPresentationGradientFill(this.fill)
      ? presentationGradientFillSvg(this.fill, `${this.id}-fill`, `Presentation shape ${this.name || this.id} fill`)
      : undefined;
    const fill = this.useBackgroundFill === true
      ? resolvePresentationBackgroundColor(this.slide.effectiveBackground(), this.slide.effectiveTheme())
      : gradient?.paint || (typeof this.fill === "string" ? resolveColorToken(this.fill, this.fill) : this.fill?.color || "transparent");
    const fillOpacity = typeof this.fill === "object" && this.fill.opacity != null ? Number(this.fill.opacity) : 1;
    const fillOpacityAttribute = fillOpacity === 1 ? "" : ` fill-opacity="${fillOpacity}"`;
    const outline = this.geometry === "line"
      ? ""
      : presentationShapeLineSvgAttributes(this.line, `Presentation shape ${this.name || this.id} line`);
    const visual = this.geometry === "custom"
      ? `<g fill="${xmlEscape(fill)}"${fillOpacityAttribute} ${outline}>${presentationCustomPathsSvg(custom.paths, p, { escape: xmlEscape, adjustments: custom.adjustments, guides: custom.guides, sourceFrame: this.position })}</g>`
      : this.geometry === "line"
      ? presentationFreeLineSvg(this.line, p, `Presentation shape ${this.name || this.id}`, this.id)
      : this.geometry === "ellipse"
      ? `<ellipse cx="${p.left + p.width / 2}" cy="${p.top + p.height / 2}" rx="${p.width / 2}" ry="${p.height / 2}" fill="${xmlEscape(fill)}"${fillOpacityAttribute} ${outline}/>`
      : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" rx="${this.borderRadius ? 12 : 0}" fill="${xmlEscape(fill)}"${fillOpacityAttribute} ${outline}/>`;
    const text = this.text.value ? presentationParagraphsSvg(this.text.effectiveParagraphs(), textFrame, this.text.style, { escape: xmlEscape }) : "";
    if (!this.transform) return `${gradient?.defs || ""}${visual}${text}`;
    const cx = p.left + p.width / 2;
    const cy = p.top + p.height / 2;
    const rotation = Number(this.transform.rotationDegrees || 0);
    const flipHorizontal = this.transform.flipHorizontal === true ? -1 : 1;
    const flipVertical = this.transform.flipVertical === true ? -1 : 1;
    return `${gradient?.defs || ""}<g transform="translate(${cx} ${cy}) rotate(${rotation}) scale(${flipHorizontal} ${flipVertical}) translate(${-cx} ${-cy})">${visual}${text}</g>`;
  }

}

function presentationTableCellKey(row, column) { return `${row}:${column}`; }

function normalizePresentationTableMergeRange(range, rows, columns) {
  if (!range || typeof range !== "object" || Array.isArray(range)) throw new TypeError("Presentation table merge requires a range object.");
  const normalized = {
    startRow: Number(range.startRow),
    endRow: Number(range.endRow),
    startColumn: Number(range.startColumn),
    endColumn: Number(range.endColumn),
  };
  if (Object.values(normalized).some((value) => !Number.isInteger(value))) throw new TypeError("Presentation table merge coordinates must be integers.");
  if (normalized.startRow < 0 || normalized.startColumn < 0 || normalized.endRow < normalized.startRow || normalized.endColumn < normalized.startColumn ||
      normalized.endRow >= rows || normalized.endColumn >= columns) {
    throw new RangeError(`Presentation table merge ${normalized.startRow}:${normalized.startColumn}-${normalized.endRow}:${normalized.endColumn} is outside the ${rows}x${columns} grid.`);
  }
  if (normalized.startRow === normalized.endRow && normalized.startColumn === normalized.endColumn) throw new RangeError("Presentation table merge must span at least two cells.");
  return normalized;
}

function presentationTableMergePlan(rows, columns, ranges = []) {
  if (!Number.isInteger(rows) || rows < 1 || !Number.isInteger(columns) || columns < 1) throw new RangeError("Presentation table merges require a non-empty rectangular grid.");
  if (!Array.isArray(ranges)) throw new TypeError("Presentation table mergeRanges must be an array.");
  const cells = new Map();
  const normalizedRanges = ranges.map((range) => normalizePresentationTableMergeRange(range, rows, columns));
  for (const range of normalizedRanges) {
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        const key = presentationTableCellKey(row, column);
        if (cells.has(key)) throw new RangeError(`Presentation table merge ranges overlap at cell ${row},${column}.`);
        const origin = { row: range.startRow, column: range.startColumn };
        cells.set(key, row === range.startRow && column === range.startColumn
          ? { kind: "origin", origin, rowSpan: range.endRow - range.startRow + 1, columnSpan: range.endColumn - range.startColumn + 1, range }
          : { kind: "covered", origin, rowSpan: 0, columnSpan: 0, range });
      }
    }
  }
  return { cells, ranges: normalizedRanges };
}

class TableCellFacade {
  constructor(table, row, column) { this.table = table; this.row = row; this.column = column; this.text = new TextFrame(); }
  get value() { return this.table.values[this.row]?.[this.column] ?? ""; }
  set value(value) {
    const state = this.table.mergeState(this.row, this.column);
    if (state.kind === "covered") throw new RangeError(`Presentation table cell ${this.row},${this.column} is covered by merge origin ${state.origin.row},${state.origin.column} and is read-only.`);
    this.table.values[this.row][this.column] = value;
  }
  get editable() { return this.table.mergeState(this.row, this.column).kind !== "covered"; }
  get mergeOrigin() { return { ...this.table.mergeState(this.row, this.column).origin }; }
  get rowSpan() { return this.table.mergeState(this.row, this.column).rowSpan; }
  get columnSpan() { return this.table.mergeState(this.row, this.column).columnSpan; }
}

export class TableElement {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("tb");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
    this.name = config.name || "";
    this.rows = Number(config.rows || config.values?.length || 1);
    this.columns = Number(config.columns || config.values?.[0]?.length || 1);
    this.position = normalizeFrame(config, { left: 0, top: 0, width: 320, height: 160 });
    this.values = Array.from({ length: this.rows }, (_, r) => Array.from({ length: this.columns }, (_, c) => config.values?.[r]?.[c] ?? ""));
    this.style = config.style;
    this.styleOptions = config.styleOptions || {};
    this.accessibility = initializePresentationAccessibility(this, config, `Presentation table ${this.id}`);
    this._mergeRanges = [];
    for (const range of config.mergeRanges || (config.mergeRange ? [config.mergeRange] : [])) this._appendMergeRange(range);
    this.cells = { set: (row, column, value) => { this.getCell(row, column).value = value; }, block: (range) => ({ table: this, range }) };
    this.borders = { assign: (configValue) => { this.border = configValue; } };
  }

  _assertCell(row, column) {
    if (!Number.isInteger(row) || !Number.isInteger(column) || row < 0 || column < 0 || row >= this.rows || column >= this.columns) {
      throw new RangeError(`Presentation table cell ${row},${column} is outside the ${this.rows}x${this.columns} grid.`);
    }
  }

  _appendMergeRange(range) {
    const next = presentationTableMergePlan(this.rows, this.columns, [...this._mergeRanges, range]);
    const normalized = next.ranges.at(-1);
    for (let row = normalized.startRow; row <= normalized.endRow; row += 1) {
      for (let column = normalized.startColumn; column <= normalized.endColumn; column += 1) {
        if (row !== normalized.startRow || column !== normalized.startColumn) this.values[row][column] = "";
      }
    }
    this._mergeRanges.push(normalized);
  }

  get mergeRanges() { return this._mergeRanges.map((range) => ({ ...range })); }
  mergeState(row, column) {
    this._assertCell(row, column);
    return presentationTableMergePlan(this.rows, this.columns, this._mergeRanges).cells.get(presentationTableCellKey(row, column)) || {
      kind: "cell",
      origin: { row, column },
      rowSpan: 1,
      columnSpan: 1,
    };
  }
  getCell(row, column) { this._assertCell(row, column); return new TableCellFacade(this, row, column); }
  merge(range) { this._appendMergeRange(range); return this; }
  get accessibilityCapability() { return presentationAccessibilityCapability(this); }
  get deletionCapability() { return presentationElementDeletionCapability(this, "table"); }

  delete() {
    const owner = this.parentGroup;
    const collection = owner?.tables || this.slide?.tables;
    return deletePresentationElement(this, collection, "table");
  }

  setAccessibilityMetadata(update) {
    this.accessibility = setPresentationAccessibilityMetadata(this, this.accessibility, update, `Presentation table ${this.id}`);
    return this;
  }

  inspectRecord() {
    const p = this.position;
    return { kind: "table", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, rows: this.rows, cols: this.columns, mergeRanges: this.mergeRanges.length ? this.mergeRanges : undefined, accessibility: this.accessibility ? { ...this.accessibility } : undefined, accessibilityCapability: this.accessibilityCapability, deletionCapability: this.deletionCapability, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", values: this.values };
  }

  layoutJson() { return { kind: "table", id: this.id, name: this.name, frame: this.position, rows: this.rows, columns: this.columns, values: this.values, mergeRanges: this.mergeRanges.length ? this.mergeRanges : undefined, accessibility: this.accessibility ? { ...this.accessibility } : undefined, style: this.style, styleOptions: this.styleOptions }; }

  toSvg() {
    const p = this.position;
    const cellW = p.width / Math.max(1, this.columns);
    const cellH = p.height / Math.max(1, this.rows);
    const plan = presentationTableMergePlan(this.rows, this.columns, this._mergeRanges);
    const parts = [`<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/>`];
    for (let r = 0; r < this.rows; r++) {
      for (let c = 0; c < this.columns; c++) {
        const state = plan.cells.get(presentationTableCellKey(r, c));
        if (state?.kind === "covered") continue;
        const x = p.left + c * cellW;
        const y = p.top + r * cellH;
        const width = cellW * (state?.columnSpan || 1);
        const height = cellH * (state?.rowSpan || 1);
        const fill = this.styleOptions.headerRow && r === 0 ? "#0f172a" : r % 2 ? "#f8fafc" : "#ffffff";
        const color = this.styleOptions.headerRow && r === 0 ? "#ffffff" : "#0f172a";
        parts.push(`<rect x="${x}" y="${y}" width="${width}" height="${height}" fill="${fill}" stroke="#cbd5e1"/>`);
        parts.push(`<text x="${x + 6}" y="${y + Math.min(22, height - 6)}" font-family="Arial" font-size="13" fill="${color}">${xmlEscape(this.values[r]?.[c] ?? "")}</text>`);
      }
    }
    return parts.join("");
  }

}

const PRESENTATION_CHART_TYPES = new Set(["bar", "line", "pie", "area", "doughnut", "scatter", "bubble", "combo"]);
const PRESENTATION_NUMERIC_X_CHART_TYPES = new Set(["scatter", "bubble"]);
const PRESENTATION_CIRCULAR_CHART_TYPES = new Set(["pie", "doughnut"]);

function normalizeChartSeries(seriesItems = [], chartType = "bar") {
  return (seriesItems || []).map((series, index) => {
    const values = (series.values || series.data || []).map((value) => value);
    const xValues = (series.xValues || []).map((value) => value);
    const bubbleSizes = (series.bubbleSizes || []).map((value) => value);
    const style = normalizePresentationChartSeriesStyle(series, values.length);
    const seriesChartType = chartType === "combo" ? String(series.chartType || series.type || "").toLowerCase() : undefined;
    if (chartType === "combo" && !new Set(["bar", "line"]).has(seriesChartType)) throw new TypeError("Presentation combo chart series chartType must be bar or line.");
    const rawAxisGroup = series.axisGroup ?? series.axis ?? (series.secondaryAxis === true ? "secondary" : "primary");
    const axisGroup = normalizePresentationChartAxisGroup(rawAxisGroup === "y2" ? "secondary" : rawAxisGroup === "y1" ? "primary" : String(rawAxisGroup).toLowerCase(), seriesChartType || chartType);
    return {
      name: series.name || `Series ${index + 1}`,
      values,
      ...(xValues.length ? { xValues } : {}),
      ...(bubbleSizes.length ? { bubbleSizes } : {}),
      categories: series.categories,
      color: style.color || ["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4],
      ...(style.line ? { line: style.line } : {}),
      ...(style.points.length ? { points: style.points } : {}),
      ...(style.marker ? { marker: style.marker } : {}),
      ...(style.smooth == null ? {} : { smooth: style.smooth }),
      ...(series.dataLabels === undefined ? {} : { dataLabels: normalizePresentationChartDataLabels(series.dataLabels) }),
      ...((series.trendlines ?? series.trendline) == null ? {} : { trendlines: normalizePresentationChartTrendlines(series.trendlines ?? series.trendline, values.length, seriesChartType || chartType) }),
      ...(series.errorBars == null ? {} : { errorBars: normalizePresentationChartErrorBars(series.errorBars, seriesChartType || chartType, values.length) }),
      ...(seriesChartType ? { chartType: seriesChartType } : {}),
      ...(axisGroup === "secondary" ? { axisGroup } : {}),
    };
  });
}

function normalizeChartAxes(config = {}, hasSecondary = false) {
  const axes = config.axes || {};
  const axisTitles = config.axisTitles || {};
  const secondary = axes.secondary || {};
  const secondaryAxisTitles = axisTitles.secondary || config.secondaryAxisTitles || {};
  return {
    category: { ...(axes.category || axes.x || config.xAxis || {}), title: axes.category?.title || axes.x?.title || config.xAxis?.title || axisTitles.category || axisTitles.x || config.categoryAxisTitle || config.xAxisTitle || "" },
    value: { ...(axes.value || axes.y || config.yAxis || {}), title: axes.value?.title || axes.y?.title || config.yAxis?.title || axisTitles.value || axisTitles.y || config.valueAxisTitle || config.yAxisTitle || "" },
    ...(hasSecondary ? {
      secondary: {
        category: { ...(secondary.category || secondary.x || axes.secondaryCategory || {}), title: secondary.category?.title || secondary.x?.title || axes.secondaryCategory?.title || secondaryAxisTitles.category || secondaryAxisTitles.x || config.secondaryCategoryAxisTitle || config.secondaryXAxisTitle || "" },
        value: { ...(secondary.value || secondary.y || axes.secondaryValue || axes.y2 || {}), title: secondary.value?.title || secondary.y?.title || axes.secondaryValue?.title || axes.y2?.title || secondaryAxisTitles.value || secondaryAxisTitles.y || config.secondaryValueAxisTitle || config.secondaryYAxisTitle || "" },
      },
    } : {}),
  };
}

function normalizeChartLegend(config = {}, seriesLength = 0) {
  const normalizePosition = (value) => ({ t: "top", b: "bottom", l: "left", r: "right" }[String(value || "r")] || String(value || "r"));
  const raw = config.legend;
  if (raw === false || config.hasLegend === false) return { visible: false, position: "" };
  if (typeof raw === "string") return { visible: true, position: normalizePosition(raw) };
  const visible = raw?.visible ?? config.hasLegend ?? seriesLength > 1;
  return { visible, position: visible ? normalizePosition(raw?.position || config.legendPosition || "r") : "" };
}

function normalizeChartDataLabels(config = {}) {
  const raw = config.dataLabels ?? config.labels ?? {};
  if (raw === true || raw === false) return normalizePresentationChartDataLabels(raw);
  return normalizePresentationChartDataLabels({
    ...raw,
    showValue: raw.showValue ?? config.showValues,
    showCategoryName: raw.showCategoryName ?? raw.showCategory ?? config.showCategoryLabels,
  });
}

function pieSlicePath(cx, cy, radius, startAngle, endAngle) {
  const startX = cx + radius * Math.cos(startAngle);
  const startY = cy + radius * Math.sin(startAngle);
  const endX = cx + radius * Math.cos(endAngle);
  const endY = cy + radius * Math.sin(endAngle);
  const largeArc = endAngle - startAngle > Math.PI ? 1 : 0;
  return `M ${cx} ${cy} L ${startX} ${startY} A ${radius} ${radius} 0 ${largeArc} 1 ${endX} ${endY} Z`;
}

function doughnutSlicePath(cx, cy, outerRadius, innerRadius, startAngle, endAngle) {
  const outerStartX = cx + outerRadius * Math.cos(startAngle);
  const outerStartY = cy + outerRadius * Math.sin(startAngle);
  const outerEndX = cx + outerRadius * Math.cos(endAngle);
  const outerEndY = cy + outerRadius * Math.sin(endAngle);
  const innerStartX = cx + innerRadius * Math.cos(startAngle);
  const innerStartY = cy + innerRadius * Math.sin(startAngle);
  const innerEndX = cx + innerRadius * Math.cos(endAngle);
  const innerEndY = cy + innerRadius * Math.sin(endAngle);
  const largeArc = endAngle - startAngle > Math.PI ? 1 : 0;
  return `M ${outerStartX} ${outerStartY} A ${outerRadius} ${outerRadius} 0 ${largeArc} 1 ${outerEndX} ${outerEndY} L ${innerEndX} ${innerEndY} A ${innerRadius} ${innerRadius} 0 ${largeArc} 0 ${innerStartX} ${innerStartY} Z`;
}

function presentationChartMarkerSvg(marker, x, y, color) {
  if (!marker || marker.symbol === "none") return "";
  const size = Math.max(2, Number(marker.size) || 5);
  const radius = size / 2;
  const stroke = xmlEscape(color);
  if (marker.symbol === "square") return `<rect x="${x - radius}" y="${y - radius}" width="${size}" height="${size}" fill="${stroke}"/>`;
  if (marker.symbol === "diamond") return `<path d="M ${x} ${y - radius} L ${x + radius} ${y} L ${x} ${y + radius} L ${x - radius} ${y} Z" fill="${stroke}"/>`;
  if (marker.symbol === "triangle") return `<path d="M ${x} ${y - radius} L ${x + radius} ${y + radius} L ${x - radius} ${y + radius} Z" fill="${stroke}"/>`;
  if (marker.symbol === "x") return `<path d="M ${x - radius} ${y - radius} L ${x + radius} ${y + radius} M ${x + radius} ${y - radius} L ${x - radius} ${y + radius}" fill="none" stroke="${stroke}" stroke-width="1.5"/>`;
  if (marker.symbol === "plus") return `<path d="M ${x - radius} ${y} L ${x + radius} ${y} M ${x} ${y - radius} L ${x} ${y + radius}" fill="none" stroke="${stroke}" stroke-width="1.5"/>`;
  if (marker.symbol === "dash") return `<line x1="${x - radius}" y1="${y}" x2="${x + radius}" y2="${y}" stroke="${stroke}" stroke-width="2"/>`;
  return `<circle cx="${x}" cy="${y}" r="${marker.symbol === "dot" ? Math.max(1, radius / 2) : radius}" fill="${stroke}"/>`;
}

function presentationChartDataLabelText(dataLabels, category, value, context = {}) {
  if (!dataLabels?.showValue && !dataLabels?.showCategoryName && !dataLabels?.showSeriesName && !dataLabels?.showPercent) return "";
  const total = Number(context.total);
  const percent = dataLabels.showPercent && Number.isFinite(total) && total !== 0
    ? `${Math.round((Number(value) / total) * 1000) / 10}%`
    : undefined;
  return [
    dataLabels.showSeriesName ? context.seriesName : undefined,
    dataLabels.showCategoryName ? category : undefined,
    dataLabels.showValue ? value : undefined,
    percent,
  ].filter((item) => item != null && item !== "").map(String).join(": ");
}

function presentationChartErrorBarsSvg(series, points, plot, max, seriesIndex = 0, magnitudeDivisors, xValueAxis = false) {
  const errorBars = series.errorBars;
  if (!errorBars || !points.length) return "";
  const magnitudes = chartErrorBarMagnitudes(series.values || [], errorBars);
  const attributes = presentationChartLineSvgAttributes(errorBars.line || { fill: series.color || "#475569", width: 1, style: "solid" });
  return points.map((point, index) => {
    const pointIndex = point.index ?? index;
    const scale = errorBars.direction === "x"
      ? xValueAxis
        ? plot.width / Math.max(1, max)
        : plot.width / Math.max(1, series.values?.length || 0)
      : plot.height / Math.max(1, max);
    const divisor = Number(magnitudeDivisors?.[pointIndex]) || 1;
    const minus = (magnitudes[pointIndex]?.minus || 0) / divisor * scale;
    const plus = (magnitudes[pointIndex]?.plus || 0) / divisor * scale;
    const x1 = errorBars.direction === "x" ? point.x - minus : point.x;
    const x2 = errorBars.direction === "x" ? point.x + plus : point.x;
    const y1 = errorBars.direction === "y" ? point.y + minus : point.y;
    const y2 = errorBars.direction === "y" ? point.y - plus : point.y;
    const caps = errorBars.noEndCap ? "" : errorBars.direction === "x"
      ? `${minus > 0 ? `<line x1="${x1}" y1="${point.y - 4}" x2="${x1}" y2="${point.y + 4}"${attributes}/>` : ""}${plus > 0 ? `<line x1="${x2}" y1="${point.y - 4}" x2="${x2}" y2="${point.y + 4}"${attributes}/>` : ""}`
      : `${minus > 0 ? `<line x1="${point.x - 4}" y1="${y1}" x2="${point.x + 4}" y2="${y1}"${attributes}/>` : ""}${plus > 0 ? `<line x1="${point.x - 4}" y1="${y2}" x2="${point.x + 4}" y2="${y2}"${attributes}/>` : ""}`;
    return `<line data-error-bars-series="${seriesIndex}" data-error-bars-index="${pointIndex}" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}"${attributes}/>${caps}`;
  }).join("");
}

export class ChartElement {
  constructor(slide, chartType = "bar", config = {}) {
    this.slide = slide;
    this.id = config.id || aid("ch");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
    this.name = config.name || "";
    this.chartType = String(chartType || config.chartType || "bar").toLowerCase();
    if (!PRESENTATION_CHART_TYPES.has(this.chartType)) throw new TypeError(`Presentation chart type must be one of: ${[...PRESENTATION_CHART_TYPES].join(", ")}.`);
    this.position = normalizeFrame(config, { left: 0, top: 0, width: 360, height: 220 });
    this.title = config.title || "";
    this.categories = config.categories || [];
    this.series = normalizeChartSeries(config.series || [], this.chartType);
    if (PRESENTATION_NUMERIC_X_CHART_TYPES.has(this.chartType) && this.categories.length) throw new TypeError(`Presentation ${this.chartType} charts use per-series xValues rather than shared categories.`);
    if (!PRESENTATION_NUMERIC_X_CHART_TYPES.has(this.chartType) && this.series.some((series) => series.xValues?.length || series.bubbleSizes?.length)) throw new TypeError(`Presentation ${this.chartType} charts cannot carry xValues or bubbleSizes.`);
    for (const [index, series] of this.series.entries()) {
      const seriesType = this.chartType === "combo" ? series.chartType : this.chartType;
      if (PRESENTATION_NUMERIC_X_CHART_TYPES.has(this.chartType) && series.xValues?.length !== series.values.length) throw new TypeError(`Presentation ${this.chartType} series ${index + 1} requires one xValue per value.`);
      if (this.chartType === "bubble" && (series.bubbleSizes?.length !== series.values.length || series.bubbleSizes.some((value) => !Number.isFinite(Number(value)) || Number(value) <= 0))) throw new TypeError(`Presentation bubble series ${index + 1} requires one positive bubbleSize per value.`);
      if (this.chartType !== "bubble" && series.bubbleSizes?.length) throw new TypeError(`Presentation ${this.chartType} charts cannot carry bubbleSizes.`);
      if (series.marker && !["line", "scatter"].includes(seriesType)) throw new TypeError(`Presentation ${seriesType} series ${index + 1} cannot carry a marker.`);
    }
    this.externalData = normalizePresentationChartExternalData(config.externalData ?? config.sourceWorkbook);
    if (presentationChartUsesFormulaReferences(this) && !this.externalData) throw new TypeError("Presentation chart formula references require externalData with an embedded workbook or external workbook URI.");
    if (this.chartType === "combo" && (!this.series.some((series) => series.chartType === "bar") || !this.series.some((series) => series.chartType === "line"))) throw new TypeError("Presentation combo chart requires at least one bar series and one line series.");
    const hasSecondary = this.series.some((series) => series.axisGroup === "secondary");
    const hasConfiguredSecondaryAxes = Boolean(config.axes?.secondary || config.axes?.secondaryCategory || config.axes?.secondaryValue || config.axes?.y2 || config.secondaryAxisTitles || config.secondaryCategoryAxisTitle || config.secondaryValueAxisTitle || config.secondaryXAxisTitle || config.secondaryYAxisTitle);
    if (hasConfiguredSecondaryAxes && !hasSecondary) throw new TypeError("Presentation secondary axes require at least one chart series with axisGroup secondary.");
    if (hasSecondary && !this.series.some((series) => series.axisGroup !== "secondary")) throw new TypeError("Presentation secondary-axis charts require at least one primary-axis series.");
    this.axes = normalizeChartAxes(config, hasSecondary);
    this.legend = normalizeChartLegend(config, this.series.length);
    this.hasLegend = this.legend.visible;
    this.dataLabels = normalizeChartDataLabels(config);
    this.accessibility = initializePresentationAccessibility(this, config, `Presentation chart ${this.id}`);
    if (this.dataLabels.showPercent && !PRESENTATION_CIRCULAR_CHART_TYPES.has(this.chartType)) throw new TypeError("Presentation percentage data labels require a pie or doughnut chart.");
    if (PRESENTATION_CIRCULAR_CHART_TYPES.has(this.chartType) && (config.axes || config.xAxis || config.yAxis || config.axisTitles || config.categoryAxisTitle || config.valueAxisTitle || config.xAxisTitle || config.yAxisTitle)) throw new TypeError(`Presentation ${this.chartType} charts cannot carry axes.`);
    Object.assign(this, normalizePresentationChartStyle(this.chartType, config));
  }

  get accessibilityCapability() { return presentationAccessibilityCapability(this); }
  get deletionCapability() { return presentationElementDeletionCapability(this, "chart"); }

  delete() {
    const owner = this.parentGroup;
    const collection = owner?.charts || this.slide?.charts;
    return deletePresentationElement(this, collection, "chart");
  }

  setAccessibilityMetadata(update) {
    this.accessibility = setPresentationAccessibilityMetadata(this, this.accessibility, update, `Presentation chart ${this.id}`);
    return this;
  }

  inspectRecord() {
    const p = this.position;
    return { kind: "chart", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, chartType: this.chartType, title: this.title, categories: this.categories, series: this.series.length, seriesDetails: this.series, axes: PRESENTATION_CIRCULAR_CHART_TYPES.has(this.chartType) ? undefined : this.axes, legend: this.legend, dataLabels: this.dataLabels, accessibility: this.accessibility ? { ...this.accessibility } : undefined, accessibilityCapability: this.accessibilityCapability, deletionCapability: this.deletionCapability, externalData: this.externalData ? { embedded: Boolean(this.externalData.bytes), uri: this.externalData.uri, autoUpdate: this.externalData.autoUpdate, bytes: this.externalData.bytes?.byteLength } : undefined, styleId: this.styleId, varyColors: this.varyColors, barOptions: ["bar", "combo"].includes(this.chartType) ? this.barOptions : undefined, lineOptions: ["line", "combo"].includes(this.chartType) ? this.lineOptions : undefined, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px" };
  }

  layoutJson() { return { kind: "chart", id: this.id, name: this.name, chartType: this.chartType, title: this.title, frame: this.position, categories: this.categories, series: this.series, axes: PRESENTATION_CIRCULAR_CHART_TYPES.has(this.chartType) ? undefined : this.axes, legend: this.legend, dataLabels: this.dataLabels, accessibility: this.accessibility ? { ...this.accessibility } : undefined, externalData: this.externalData ? { embedded: Boolean(this.externalData.bytes), uri: this.externalData.uri, autoUpdate: this.externalData.autoUpdate, bytes: this.externalData.bytes?.byteLength } : undefined, styleId: this.styleId, varyColors: this.varyColors, barOptions: ["bar", "combo"].includes(this.chartType) ? this.barOptions : undefined, lineOptions: ["line", "combo"].includes(this.chartType) ? this.lineOptions : undefined }; }

  toSvg() {
    const p = this.position;
    const categories = this.categories.length ? this.categories : Array.from({ length: Math.max(0, ...this.series.map((series) => series.values?.length || 0)) }, (_, index) => String(index + 1));
    const barSeries = this.chartType === "combo" ? this.series.filter((series) => series.chartType === "bar") : this.chartType === "bar" ? this.series : [];
    const lineSeries = this.chartType === "combo" ? this.series.filter((series) => series.chartType === "line") : this.chartType === "line" ? this.series : [];
    const stackedBars = barSeries.length > 0 && this.barOptions.grouping !== "clustered";
    const stackedLines = lineSeries.length > 0 && this.lineOptions.grouping !== "standard";
    const horizontal = barSeries.length > 0 && this.barOptions.direction === "bar";
    const forAxisGroup = (series, axisGroup) => series.filter((item) => (item.axisGroup || "primary") === axisGroup);
    const stackedTotals = (series) => categories.map((_, categoryIndex) => series.reduce((sum, item) => sum + Math.max(0, Number(item.values?.[categoryIndex]) || 0), 0));
    const barByAxis = { primary: forAxisGroup(barSeries, "primary"), secondary: forAxisGroup(barSeries, "secondary") };
    const lineByAxis = { primary: forAxisGroup(lineSeries, "primary"), secondary: forAxisGroup(lineSeries, "secondary") };
    const barStackedMax = { primary: stackedTotals(barByAxis.primary), secondary: stackedTotals(barByAxis.secondary) };
    const lineStackedMax = { primary: stackedTotals(lineByAxis.primary), secondary: stackedTotals(lineByAxis.secondary) };
    const groupMax = (series, stacked, stackedValues, percentStacked) => percentStacked
      ? 1
      : Math.max(0, ...(stacked ? stackedValues : series.flatMap((item) => item.values || []).map((value) => Math.max(0, Number(value) || 0))));
    const groupErrorBarMax = (series, stacked, percentStacked, numericDirection) => {
      const totals = stackedTotals(series);
      return Math.max(0, ...series.flatMap((item, seriesIndex) => {
        if (item.errorBars?.direction !== numericDirection) return [];
        const magnitudes = chartErrorBarMagnitudes(item.values || [], item.errorBars);
        return (item.values || []).map((rawValue, pointIndex) => {
          const base = stacked
            ? series.slice(0, seriesIndex + 1).reduce((sum, candidate) => sum + Math.max(0, Number(candidate.values?.[pointIndex]) || 0), 0)
            : Math.max(0, Number(rawValue) || 0);
          const divisor = percentStacked ? totals[pointIndex] || 1 : 1;
          return (base + (magnitudes[pointIndex]?.plus || 0)) / divisor;
        });
      }));
    };
    const barMax = {
      primary: groupMax(barByAxis.primary, stackedBars, barStackedMax.primary, this.barOptions?.grouping === "percentStacked"),
      secondary: groupMax(barByAxis.secondary, stackedBars, barStackedMax.secondary, this.barOptions?.grouping === "percentStacked"),
    };
    const lineMax = {
      primary: groupMax(lineByAxis.primary, stackedLines, lineStackedMax.primary, this.lineOptions?.grouping === "percentStacked"),
      secondary: groupMax(lineByAxis.secondary, stackedLines, lineStackedMax.secondary, this.lineOptions?.grouping === "percentStacked"),
    };
    const errorBarMax = {
      primary: Math.max(
        groupErrorBarMax(barByAxis.primary, stackedBars, this.barOptions?.grouping === "percentStacked", horizontal ? "x" : "y"),
        groupErrorBarMax(lineByAxis.primary, stackedLines, this.lineOptions?.grouping === "percentStacked", "y"),
      ),
      secondary: Math.max(
        groupErrorBarMax(barByAxis.secondary, stackedBars, this.barOptions?.grouping === "percentStacked", horizontal ? "x" : "y"),
        groupErrorBarMax(lineByAxis.secondary, stackedLines, this.lineOptions?.grouping === "percentStacked", "y"),
      ),
    };
    const maxForAxisGroup = (axisGroup) => {
      const dataMaximum = Math.max(1, barMax[axisGroup], lineMax[axisGroup]);
      const errorMaximum = errorBarMax[axisGroup];
      return errorMaximum > dataMaximum ? errorMaximum * 1.08 : dataMaximum;
    };
    const primaryMax = maxForAxisGroup("primary");
    const secondaryMax = maxForAxisGroup("secondary");
    const hasSecondary = this.series.some((series) => series.axisGroup === "secondary");
    const plot = { left: p.left + 42, top: p.top + 42, width: Math.max(0, p.width - 72), height: Math.max(0, p.height - 82) };
    const title = `<text x="${p.left + 12}" y="${p.top + 24}" font-family="Arial" font-size="16" font-weight="700" fill="#0f172a">${xmlEscape(this.title || this.chartType)}</text>`;
    const axes = `<line x1="${plot.left}" y1="${plot.top + plot.height}" x2="${plot.left + plot.width}" y2="${plot.top + plot.height}" stroke="#94a3b8"/><line x1="${plot.left}" y1="${plot.top}" x2="${plot.left}" y2="${plot.top + plot.height}" stroke="#94a3b8"/>${hasSecondary ? `<line x1="${plot.left}" y1="${plot.top}" x2="${plot.left + plot.width}" y2="${plot.top}" stroke="#64748b"/><line x1="${plot.left + plot.width}" y1="${plot.top}" x2="${plot.left + plot.width}" y2="${plot.top + plot.height}" stroke="#64748b"/>` : ""}${this.axes.category.title ? `<text x="${plot.left + plot.width / 2 - 24}" y="${p.top + p.height - 4}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.category.title)}</text>` : ""}${this.axes.value.title ? `<text x="${p.left + 8}" y="${plot.top + 10}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.value.title)}</text>` : ""}${this.axes.secondary?.category?.title ? `<text x="${plot.left + plot.width / 2 - 24}" y="${plot.top - 4}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.secondary.category.title)}</text>` : ""}${this.axes.secondary?.value?.title ? `<text x="${plot.left + plot.width - 2}" y="${plot.top + 10}" text-anchor="end" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.secondary.value.title)}</text>` : ""}`;
    const legend = this.legend.visible ? this.series.map((series, index) => `<rect x="${p.left + p.width - 82}" y="${p.top + 18 + index * 16}" width="10" height="10" fill="${xmlEscape(resolveColorToken(series.color, series.color))}"/><text x="${p.left + p.width - 68}" y="${p.top + 27 + index * 16}" font-family="Arial" font-size="10" fill="#334155">${xmlEscape(series.name)}</text>`).join("") : "";
    if (PRESENTATION_CIRCULAR_CHART_TYPES.has(this.chartType)) {
      const series = this.series[0] || { values: [] };
      const values = (series.values || []).map((value) => Math.max(0, Number(value) || 0));
      const total = values.reduce((sum, value) => sum + value, 0) || 1;
      const radius = Math.max(8, Math.min(plot.width, plot.height) / 2);
      const innerRadius = this.chartType === "doughnut" ? radius * 0.5 : 0;
      const cx = plot.left + plot.width / 2;
      const cy = plot.top + plot.height / 2;
      let angle = -Math.PI / 2;
      const slices = values.map((value, index) => {
        const next = angle + (value / total) * Math.PI * 2;
        const point = series.points?.find((item) => item.idx === index);
        const color = resolveColorToken(point?.fill || ["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4], "#0ea5e9");
        const effectiveLabels = series.dataLabels || this.dataLabels;
        const labelText = presentationChartDataLabelText(effectiveLabels, categories[index], value, { total, seriesName: series.name });
        const label = labelText ? `<text x="${cx + (radius + 8) * Math.cos((angle + next) / 2)}" y="${cy + (radius + 8) * Math.sin((angle + next) / 2)}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(labelText)}</text>` : "";
        const geometry = innerRadius > 0 ? doughnutSlicePath(cx, cy, radius, innerRadius, angle, next) : pieSlicePath(cx, cy, radius, angle, next);
        const path = `<path d="${geometry}" fill="${xmlEscape(color)}"${presentationChartLineSvgAttributes(point?.line || series.line) || ' stroke="#ffffff"'}/>${label}`;
        angle = next;
        return path;
      }).join("");
      const categoryLegend = categories.map((category, index) => `<rect x="${p.left + p.width - 82}" y="${p.top + 18 + index * 16}" width="10" height="10" fill="${xmlEscape(resolveColorToken(series.points?.find((item) => item.idx === index)?.fill || ["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4], "#0ea5e9"))}"/><text x="${p.left + p.width - 68}" y="${p.top + 27 + index * 16}" font-family="Arial" font-size="10" fill="#334155">${xmlEscape(category)}</text>`).join("");
      return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/>${title}${slices}${this.legend.visible ? categoryLegend : ""}`;
    }
    if (PRESENTATION_NUMERIC_X_CHART_TYPES.has(this.chartType)) {
      const xValues = this.series.flatMap((series) => series.xValues || []).map(Number).filter(Number.isFinite);
      const yValues = this.series.flatMap((series) => series.values || []).map(Number).filter(Number.isFinite);
      const configuredXMin = Number(this.axes.category?.min);
      const configuredXMax = Number(this.axes.category?.max);
      const configuredYMin = Number(this.axes.value?.min);
      const configuredYMax = Number(this.axes.value?.max);
      const xMin = Number.isFinite(configuredXMin) ? configuredXMin : Math.min(0, ...xValues);
      const xMax = Number.isFinite(configuredXMax) ? configuredXMax : Math.max(1, ...xValues);
      const yMin = Number.isFinite(configuredYMin) ? configuredYMin : Math.min(0, ...yValues);
      const yMax = Number.isFinite(configuredYMax) ? configuredYMax : Math.max(1, ...yValues);
      const mapX = (value) => plot.left + ((Number(value) - xMin) / Math.max(Number.EPSILON, xMax - xMin)) * plot.width;
      const mapY = (value) => plot.top + plot.height - ((Number(value) - yMin) / Math.max(Number.EPSILON, yMax - yMin)) * plot.height;
      const maxBubble = Math.max(1, ...this.series.flatMap((series) => series.bubbleSizes || []).map(Number).filter(Number.isFinite));
      const body = this.series.map((series) => {
        const color = resolveColorToken(series.color || series.fill, "#0ea5e9");
        return (series.values || []).map((value, index) => {
          const x = mapX(series.xValues?.[index]);
          const y = mapY(value);
          if (this.chartType === "bubble") {
            const radius = Math.max(3, Math.sqrt(Number(series.bubbleSizes?.[index]) / maxBubble) * Math.min(28, Math.max(8, Math.min(plot.width, plot.height) / 5)));
            const label = presentationChartDataLabelText(series.dataLabels || this.dataLabels, series.xValues?.[index], value, { seriesName: series.name });
            return `<circle cx="${x}" cy="${y}" r="${radius}" fill="${xmlEscape(color)}" fill-opacity="0.72"${presentationChartLineSvgAttributes(series.line)}/>${label ? `<text x="${x + radius + 3}" y="${y - 3}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(label)}</text>` : ""}`;
          }
          const marker = series.marker || { symbol: "circle", size: 7 };
          const label = presentationChartDataLabelText(series.dataLabels || this.dataLabels, series.xValues?.[index], value, { seriesName: series.name });
          return `${presentationChartMarkerSvg(marker, x, y, color)}${label ? `<text x="${x + 5}" y="${y - 5}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(label)}</text>` : ""}`;
        }).join("");
      }).join("");
      return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/>${title}${axes}${body}${legend}`;
    }
    if (this.chartType === "area") {
      const max = Math.max(1, ...this.series.flatMap((series) => series.values || []).map((value) => Math.max(0, Number(value) || 0)));
      const body = this.series.map((series) => {
        const points = (series.values || []).map((value, index) => ({
          x: plot.left + (categories.length <= 1 ? plot.width / 2 : (index / Math.max(1, categories.length - 1)) * plot.width),
          y: plot.top + plot.height - (Math.max(0, Number(value) || 0) / max) * plot.height,
        }));
        if (!points.length) return "";
        const color = resolveColorToken(series.color, "#0ea5e9");
        const path = `M ${plot.left} ${plot.top + plot.height} L ${points.map((point) => `${point.x} ${point.y}`).join(" L ")} L ${plot.left + plot.width} ${plot.top + plot.height} Z`;
        return `<path d="${path}" fill="${xmlEscape(color)}" fill-opacity="0.45"${presentationChartLineSvgAttributes(series.line) || ` stroke="${xmlEscape(color)}" stroke-width="1.5"`}/>`;
      }).join("");
      const labels = categories.map((category, index) => `<text x="${plot.left + index * (plot.width / Math.max(1, categories.length))}" y="${p.top + p.height - 18}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(category)}</text>`).join("");
      return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/>${title}${axes}${body}${labels}${legend}`;
    }
    const lineBody = lineSeries.map((series, seriesIndex) => {
        const axisGroup = series.axisGroup || "primary";
        const seriesMax = axisGroup === "secondary" ? secondaryMax : primaryMax;
        const points = (series.values || []).map((value, index) => {
          const stackedValue = stackedLines ? lineSeries.slice(0, seriesIndex + 1).filter((item) => (item.axisGroup || "primary") === axisGroup).reduce((sum, item) => sum + Math.max(0, Number(item.values?.[index]) || 0), 0) : Number(value) || 0;
          const plottedValue = this.lineOptions.grouping === "percentStacked" ? stackedValue / (lineStackedMax[axisGroup][index] || 1) : stackedValue;
          const x = plot.left + (categories.length <= 1 ? plot.width / 2 : (index / Math.max(1, categories.length - 1)) * plot.width);
          const y = plot.top + plot.height - (plottedValue / seriesMax) * plot.height;
          return { x, y, index };
        });
        const color = resolveColorToken(series.line?.fill || series.color, series.color);
        const smooth = series.smooth ?? this.lineOptions.smooth;
        const strokeAttributes = presentationChartLineSvgAttributes(series.line) || ` stroke="${xmlEscape(color)}" stroke-width="2"`;
        const line = smooth && points.length > 2
          ? `<path d="M ${points[0].x} ${points[0].y} ${points.slice(1, -1).map((point, index) => { const next = points[index + 2]; return `Q ${point.x} ${point.y} ${(point.x + next.x) / 2} ${(point.y + next.y) / 2}`; }).join(" ")} T ${points.at(-1).x} ${points.at(-1).y}" fill="none"${strokeAttributes}/>`
          : `<polyline points="${points.map((point) => `${point.x},${point.y}`).join(" ")}" fill="none"${strokeAttributes}/>`;
        const marker = series.marker || this.lineOptions.marker;
        const effectiveLabels = series.dataLabels || this.dataLabels;
        const labels = points.map((point, index) => {
          const label = presentationChartDataLabelText(effectiveLabels, categories[index], series.values?.[index]);
          return label ? `<text x="${point.x + 4}" y="${point.y - 4}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(label)}</text>` : "";
        }).join("");
        return `${line}${presentationChartErrorBarsSvg(series, points, plot, seriesMax, this.series.indexOf(series), this.lineOptions.grouping === "percentStacked" ? lineStackedMax[axisGroup] : undefined)}${points.map((point, index) => presentationChartMarkerSvg(marker, point.x, point.y, resolveColorToken(series.points?.find((item) => item.idx === index)?.fill || color, color))).join("")}${labels}`;
      }).join("");
    const barBody = (() => {
      const groupExtent = categories.length ? (horizontal ? plot.height : plot.width) / categories.length : 0;
      const gapRatio = Math.max(0.12, 100 / (100 + this.barOptions.gapWidth));
      const barExtent = stackedBars ? groupExtent * gapRatio : groupExtent * gapRatio / Math.max(1, barSeries.length);
      const offsets = { primary: categories.map(() => 0), secondary: categories.map(() => 0) };
      return barSeries.flatMap((series, seriesIndex) => (series.values || []).map((rawValue, categoryIndex) => {
        const axisGroup = series.axisGroup || "primary";
        const seriesMax = axisGroup === "secondary" ? secondaryMax : primaryMax;
        const total = barStackedMax[axisGroup][categoryIndex] || 1;
        const value = Math.max(0, Number(rawValue) || 0);
        const ratio = this.barOptions.grouping === "percentStacked" ? value / total : value / seriesMax;
        const offset = offsets[axisGroup][categoryIndex];
        offsets[axisGroup][categoryIndex] += ratio;
        const point = series.points?.find((item) => item.idx === categoryIndex);
        const color = xmlEscape(resolveColorToken(point?.fill || series.color, series.color));
        const stroke = presentationChartLineSvgAttributes(point?.line || series.line);
        const labelText = presentationChartDataLabelText(series.dataLabels || this.dataLabels, categories[categoryIndex], rawValue);
        if (horizontal) {
          const width = plot.width * ratio;
          const x = plot.left + (stackedBars ? plot.width * offset : 0);
          const y = plot.top + categoryIndex * groupExtent + (stackedBars ? (groupExtent - barExtent) / 2 : (groupExtent - barExtent * barSeries.length) / 2 + seriesIndex * barExtent);
          const label = labelText ? `<text x="${x + width + 3}" y="${y + barExtent - 2}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(labelText)}</text>` : "";
          const errorBars = presentationChartErrorBarsSvg(series, [{ x: x + width, y: y + Math.max(1, barExtent - 2) / 2, index: categoryIndex }], plot, seriesMax, this.series.indexOf(series), this.barOptions.grouping === "percentStacked" ? barStackedMax[axisGroup] : undefined, true);
          return `<rect x="${x}" y="${y}" width="${width}" height="${Math.max(1, barExtent - 2)}" fill="${color}"${stroke}/>${errorBars}${label}`;
        }
        const height = plot.height * ratio;
        const x = plot.left + categoryIndex * groupExtent + (stackedBars ? (groupExtent - barExtent) / 2 : (groupExtent - barExtent * barSeries.length) / 2 + seriesIndex * barExtent);
        const y = plot.top + plot.height - height - (stackedBars ? plot.height * offset : 0);
        const label = labelText ? `<text x="${x}" y="${y - 4}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(labelText)}</text>` : "";
        const errorBars = presentationChartErrorBarsSvg(series, [{ x: x + Math.max(1, barExtent - 2) / 2, y, index: categoryIndex }], plot, seriesMax, this.series.indexOf(series), this.barOptions.grouping === "percentStacked" ? barStackedMax[axisGroup] : undefined);
        return `<rect x="${x}" y="${y}" width="${Math.max(1, barExtent - 2)}" height="${height}" fill="${color}"${stroke}/>${errorBars}${label}`;
      })).join("");
    })();
    const trendlineBody = `${barSeries.map((series) => presentationChartTrendlinesSvg(series, plot, series.axisGroup === "secondary" ? secondaryMax : primaryMax, categories.length, { horizontal, centered: true })).join("")}${lineSeries.map((series) => presentationChartTrendlinesSvg(series, plot, series.axisGroup === "secondary" ? secondaryMax : primaryMax, categories.length)).join("")}`;
    const body = `${barBody}${lineBody}${trendlineBody}`;
    const labels = this.chartType === "bar" && horizontal
      ? categories.map((category, index) => `<text x="${plot.left - 4}" y="${plot.top + (index + 0.6) * (plot.height / Math.max(1, categories.length))}" text-anchor="end" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(category)}</text>`).join("")
      : categories.map((category, index) => `<text x="${plot.left + index * (plot.width / Math.max(1, categories.length))}" y="${p.top + p.height - 18}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(category)}</text>`).join("");
    return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/>${title}${axes}${body}${labels}${legend}`;
  }

}

function presentationSvgLeafScope(image) {
  const presentation = image.slide?.presentation;
  const state = presentation?.[PRESENTATION_STATE];
  const sourceRevisionSha256 = String(state?.opaqueOpc?.sourcePackage?.sha256 || state?.source?.packageSha256 || "").toLowerCase();
  return {
    scopeId: `${sourceRevisionSha256 || "authored"}\0${presentation?.id || ""}\0${image.slide?.id || ""}\0${image.id}`,
    ...(/^[0-9a-f]{64}$/u.test(sourceRevisionSha256) ? { sourceRevisionSha256 } : {}),
  };
}

const PRESENTATION_EMBEDDED_IMAGE_CONTENT_TYPES = new Set([
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/svg+xml",
]);

function presentationImageDataUrlFromBlob(blob, contentType, label) {
  const blobType = blob instanceof FileBlob || blob?.bytes instanceof Uint8Array
    ? blob.type
    : undefined;
  const bytes = blob instanceof FileBlob || blob?.bytes instanceof Uint8Array
    ? blob.bytes
    : toUint8Array(blob);
  const resolvedContentType = String(contentType || blobType || "").trim().toLowerCase();
  if (!PRESENTATION_EMBEDDED_IMAGE_CONTENT_TYPES.has(resolvedContentType)) {
    throw new TypeError(`${label} blob requires contentType image/png, image/jpeg, image/gif, or image/svg+xml.`);
  }
  if (bytes.byteLength === 0) throw new TypeError(`${label} blob cannot be empty.`);
  return {
    contentType: resolvedContentType,
    dataUrl: `data:${resolvedContentType};base64,${Buffer.from(bytes).toString("base64")}`,
  };
}

function normalizeNativePresentationBackgroundImage(config, label, previous) {
  if (!config || typeof config !== "object" || Array.isArray(config)) throw new TypeError(`${label} requires an options object.`);
  if (config.uri != null || config.fit != null && String(config.fit) !== "stretch" || config.crop != null || config.transform != null) {
    throw new TypeError(`${label} only supports an embedded stretch image without crop or transform.`);
  }
  const assetId = config.assetId == null ? undefined : String(config.assetId).trim();
  const alphaModulationFixed = config.alphaModulationFixed === undefined
    ? previous?.image?.alphaModulationFixed === true
    : config.alphaModulationFixed;
  if (typeof alphaModulationFixed !== "boolean") throw new TypeError(`${label} alphaModulationFixed must be boolean.`);
  const embedded = config.blob == null
    ? { dataUrl: config.dataUrl }
    : presentationImageDataUrlFromBlob(config.blob, config.contentType, `${label} image`);
  if (!embedded.dataUrl && !assetId) throw new TypeError(`${label} requires dataUrl, blob, or assetId.`);
  return normalizePresentationBackground({ image: { ...(assetId ? { assetId } : {}), ...(embedded.dataUrl ? { dataUrl: embedded.dataUrl } : {}), fit: "stretch", ...(alphaModulationFixed ? { alphaModulationFixed: true } : {}) } });
}

export class ImageElement {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("im");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
    this.name = config.name || "";
    this.position = normalizeFrame(config, { left: 0, top: 0, width: 320, height: 180 });
    const hasLegacyAlt = Object.hasOwn(config, "alt");
    const legacyAlt = config.alt == null ? "" : config.alt;
    if (hasLegacyAlt && config.accessibility?.description != null && config.accessibility.description !== legacyAlt) {
      throw new TypeError(`Presentation image ${this.id} alt and accessibility.description must match when both are provided.`);
    }
    const accessibility = hasLegacyAlt
      ? {
          ...(config.accessibility || {}),
          ...(legacyAlt === "" ? {} : { description: legacyAlt }),
        }
      : config.accessibility;
    this.accessibility = initializePresentationAccessibility(
      this,
      { ...config, accessibility },
      `Presentation image ${this.id}`,
    );
    const embedded = config.blob == null
      ? undefined
      : presentationImageDataUrlFromBlob(config.blob, config.contentType, `Presentation image ${this.id}`);
    if (embedded && (config.dataUrl != null || config.uri != null)) {
      throw new TypeError(`Presentation image ${this.id} blob cannot be combined with dataUrl or uri.`);
    }
    this.prompt = config.prompt;
    this.uri = config.uri;
    const importedDataUrlSource = config._officeKitDataUrlSource;
    if (embedded && importedDataUrlSource) {
      throw new TypeError(`Presentation image ${this.id} blob cannot replace an imported lazy dataUrl source.`);
    }
    if (importedDataUrlSource) {
      if (typeof importedDataUrlSource.resolve !== "function" || !importedDataUrlSource.asset) {
        throw new TypeError(`Presentation image ${this.id} received an invalid imported dataUrl source.`);
      }
      const binding = { source: importedDataUrlSource, modified: false };
      let resolved = false;
      let value;
      binding.get = function getImportedDataUrl() {
        if (!resolved) {
          value = binding.source.resolve();
          resolved = true;
        }
        return value;
      };
      binding.set = function setImportedDataUrl(next) {
        value = next;
        resolved = true;
        binding.modified = true;
      };
      Object.defineProperty(this, PRESENTATION_IMAGE_DATA_URL_SOURCE, { value: binding });
      Object.defineProperty(this, "dataUrl", {
        configurable: true,
        enumerable: true,
        get: binding.get,
        set: binding.set,
      });
    } else {
      this.dataUrl = embedded?.dataUrl ?? config.dataUrl;
    }
    const importedSvgDataUrlSource = config._officeKitSvgDataUrlSource;
    if (importedSvgDataUrlSource) {
      if (typeof importedSvgDataUrlSource.resolve !== "function" || !importedSvgDataUrlSource.asset) {
        throw new TypeError(`Presentation image ${this.id} received an invalid imported SVG dataUrl source.`);
      }
      const binding = { source: importedSvgDataUrlSource, modified: false };
      let resolved = false;
      let value;
      binding.get = function getImportedSvgDataUrl() {
        if (!resolved) {
          value = binding.source.resolve();
          resolved = true;
        }
        return value;
      };
      binding.set = function setImportedSvgDataUrl(next) {
        value = next;
        resolved = true;
        binding.modified = true;
      };
      Object.defineProperty(this, PRESENTATION_IMAGE_SVG_DATA_URL_SOURCE, { value: binding });
      Object.defineProperty(this, "svgDataUrl", {
        configurable: true,
        enumerable: true,
        get: binding.get,
        set: binding.set,
      });
    } else {
      this.svgDataUrl = config.svgDataUrl;
    }
    this.contentType = embedded?.contentType ?? config.contentType;
    this.fit = config.fit || "contain";
    this.crop = config.crop;
    this.geometry = config.geometry || "rect";
    this.borderRadius = config.borderRadius;
    this.transform = config.transform == null ? undefined : normalizePresentationPlaceholderTransform(config.transform, `Presentation image ${this.name || this.id} transform`);
  }

  get frame() { return this.position; }
  set frame(value) { this.position = normalizeFrame(value, this.position); }
  get alt() { return this.accessibility?.description || ""; }
  set alt(value) {
    const next = value == null ? "" : value;
    if (next === this.alt) return;
    this.accessibility = setPresentationAccessibilityMetadata(
      this,
      this.accessibility,
      { description: next === "" ? null : next },
      `Presentation image ${this.id}`,
    );
  }
  get accessibilityCapability() { return presentationAccessibilityCapability(this); }
  get deletionCapability() { return presentationElementDeletionCapability(this, "image"); }
  delete() {
    const owner = this.parentGroup;
    const collection = owner?.images || this.slide?.images;
    return deletePresentationElement(this, collection, "image");
  }
  get fit() { return this._fit; }
  set fit(value) { this._fit = normalizePresentationImageFit(value); }
  get crop() { return this._crop; }
  set crop(value) { this._crop = normalizePresentationImageCrop(value); }
  setAccessibilityMetadata(update) {
    this.accessibility = setPresentationAccessibilityMetadata(this, this.accessibility, update, `Presentation image ${this.id}`);
  }
  get svgTextCapability() { return inspectSvgText(this.svgDataUrl || this.dataUrl); }
  get svgEditCapability() { return inspectSvgLeaves(this.svgDataUrl || this.dataUrl, presentationSvgLeafScope(this)); }
  getSvgTextNodes() {
    const capability = this.svgTextCapability;
    return capability.nodes ? capability.nodes.map((node) => ({ ...node })) : [];
  }
  editSvgText(nodeId, update = {}) {
    const capability = this.svgTextCapability;
    if (capability.supported !== true) {
      const error = new Error(`Presentation image ${this.id} does not expose editable SVG text: ${capability.reason}.`);
      error.code = "unsupported_presentation_svg_text";
      throw error;
    }
    const dataUrl = replaceSvgTextNode(this.svgDataUrl || this.dataUrl, nodeId, update);
    if (this.svgDataUrl) this.svgDataUrl = dataUrl;
    else this.replace({ dataUrl });
    const next = this.svgTextCapability.nodes?.find((node) => node.id === nodeId);
    return Object.freeze({
      kind: "svgTextEdit",
      imageId: this.id,
      nodeId,
      oldValue: capability.nodes.find((node) => node.id === nodeId)?.text,
      value: next?.text,
      expectedHash: capability.nodes.find((node) => node.id === nodeId)?.expectedHash,
      sourceSha256: capability.sourceSha256,
    });
  }
  getSvgEditLeaves() {
    const capability = this.svgEditCapability;
    return capability.leaves ? capability.leaves.map((leaf) => ({ ...leaf })) : [];
  }
  editSvgLeaf(leafId, update = {}) {
    const capability = this.svgEditCapability;
    if (capability.supported !== true) {
      const error = new Error(`Presentation image ${this.id} does not expose editable SVG leaves: ${capability.reason}.`);
      error.code = "unsupported_presentation_svg_leaf";
      throw error;
    }
    const leaf = capability.leaves.find((candidate) => candidate.id === leafId);
    const leafIndex = capability.leaves.findIndex((candidate) => candidate.id === leafId);
    const dataUrl = replaceSvgLeaf(this.svgDataUrl || this.dataUrl, leafId, update, presentationSvgLeafScope(this));
    if (this.svgDataUrl) this.svgDataUrl = dataUrl;
    else this.replace({ dataUrl });
    const next = leafIndex >= 0 ? this.svgEditCapability.leaves?.[leafIndex] : undefined;
    return Object.freeze({
      kind: "svgLeafEdit",
      imageId: this.id,
      leafId,
      leafKind: leaf?.leafKind,
      oldValue: leaf?.value,
      value: next?.value,
      expectedHash: leaf?.expectedHash,
      sourceSha256: capability.sourceSha256,
    });
  }
  replace(config = {}) {
    const { alt, accessibility, blob, ...rest } = config;
    let nextAccessibility = this.accessibility;
    if (accessibility !== undefined) {
      if (Object.hasOwn(config, "alt") && accessibility?.description != null && accessibility.description !== (alt == null ? "" : alt)) {
        throw new TypeError(`Presentation image ${this.id} alt and accessibility.description must match when both are provided.`);
      }
      nextAccessibility = setPresentationAccessibilityMetadata(this, nextAccessibility, accessibility, `Presentation image ${this.id}`);
    }
    if (Object.hasOwn(config, "alt")) {
      const nextAlt = alt == null ? "" : alt;
      if (nextAlt !== (nextAccessibility?.description || "")) {
        nextAccessibility = setPresentationAccessibilityMetadata(
          this,
          nextAccessibility,
          { description: nextAlt === "" ? null : nextAlt },
          `Presentation image ${this.id}`,
        );
      }
    }
    if (blob != null) {
      if (rest.dataUrl != null || rest.uri != null) {
        throw new TypeError(`Presentation image ${this.id} blob cannot be combined with dataUrl or uri.`);
      }
      const embedded = presentationImageDataUrlFromBlob(blob, rest.contentType, `Presentation image ${this.id}`);
      rest.dataUrl = embedded.dataUrl;
      rest.contentType = embedded.contentType;
      rest.uri = undefined;
    }
    Object.assign(this, rest);
    this.accessibility = nextAccessibility;
  }

  inspectRecord() {
    const p = this.position;
    return { kind: "image", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, contentType: this.contentType, alt: this.alt || undefined, accessibility: this.accessibility ? { ...this.accessibility } : undefined, accessibilityCapability: this.accessibilityCapability, deletionCapability: this.deletionCapability, svgFallback: Boolean(this.svgDataUrl), svgTextCapability: this.svgTextCapability, svgEditCapability: this.svgEditCapability, prompt: this.prompt || undefined, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", fit: this.fit, crop: this.crop, transform: this.transform };
  }

  layoutJson() { return { kind: "image", id: this.id, name: this.name, frame: this.position, alt: this.alt, accessibility: this.accessibility ? { ...this.accessibility } : undefined, accessibilityCapability: this.accessibilityCapability, prompt: this.prompt, uri: this.uri, contentType: this.contentType, dataUrl: this.dataUrl, svgDataUrl: this.svgDataUrl, fit: this.fit, crop: this.crop, geometry: this.geometry, borderRadius: this.borderRadius, transform: this.transform }; }

  toSvg() {
    const p = this.position;
    const label = this.alt || this.prompt || this.uri || "image";
    const cx = p.left + p.width / 2;
    const cy = p.top + p.height / 2;
    const rotation = Number(this.transform?.rotationDegrees || 0);
    const flipHorizontal = this.transform?.flipHorizontal === true ? -1 : 1;
    const flipVertical = this.transform?.flipVertical === true ? -1 : 1;
    const transform = this.transform ? ` transform="translate(${cx} ${cy}) rotate(${rotation}) scale(${flipHorizontal} ${flipVertical}) translate(${-cx} ${-cy})"` : "";
    if (this.dataUrl) {
      const viewport = presentationImageCropViewport({ crop: this.crop, fit: this.fit, dataUrl: this.dataUrl, frame: p });
      if (viewport) {
        const cropped = `<svg x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" viewBox="${viewport.x} ${viewport.y} ${viewport.width} ${viewport.height}" preserveAspectRatio="none" overflow="hidden"><image href="${attrEscape(this.dataUrl)}" x="0" y="0" width="${viewport.imageWidth}" height="${viewport.imageHeight}" preserveAspectRatio="none"/></svg>`;
        return transform ? `<g${transform}>${cropped}</g>` : cropped;
      }
      const aspect = this.fit === "cover" ? "xMidYMid slice" : this.fit === "stretch" ? "none" : "xMidYMid meet";
      return `<image href="${attrEscape(this.dataUrl)}" x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" preserveAspectRatio="${aspect}"${transform}/>`;
    }
    const rect = this.geometry === "ellipse"
      ? `<ellipse cx="${p.left + p.width / 2}" cy="${p.top + p.height / 2}" rx="${p.width / 2}" ry="${p.height / 2}" fill="#e0f2fe" stroke="#0284c7"/>`
      : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" rx="${this.borderRadius ? 12 : 0}" fill="#e0f2fe" stroke="#0284c7"/>`;
    const fallback = `${rect}<text x="${p.left + 12}" y="${p.top + 28}" font-family="Arial" font-size="14" fill="#075985">${xmlEscape(label)}</text>`;
    return transform ? `<g${transform}>${fallback}</g>` : fallback;
  }

}

export class PresentationFile {
  static async inspectPptx(blobOrBuffer, options = {}) {
    return inspectOoxmlPackage(blobOrBuffer, options, PPTX_PACKAGE_CONFIG);
  }

  static async patchPptx(blobOrBuffer, patches = [], options = {}) {
    const patched = await patchOoxmlPackage(blobOrBuffer, patches, options, PPTX_PACKAGE_CONFIG);
    return new FileBlob(patched.bytes, { type: PPTX_MIME, metadata: { artifactKind: "presentation", patchedParts: patched.patchedParts, recipesApplied: patched.recipesApplied, contentTypesUpdated: patched.contentTypesUpdated, relationshipsUpdated: patched.relationshipsUpdated, sourceReferencesUpdated: patched.sourceReferencesUpdated, validated: patched.validated, validationIssues: patched.validationIssues } });
  }

  static async exportPptx(presentation, options = {}) {
    const { exportPptxWithOfficeKit } = await import("../codecs/office-kit-presentation-codec.mjs");
    return exportPptxWithOfficeKit(presentation, options);
  }

  static async importPptx(blobOrBuffer, options = {}) {
    const { importPptxWithOfficeKit } = await import("../codecs/office-kit-presentation-codec.mjs");
    return importPptxWithOfficeKit(blobOrBuffer, options);
  }
}

function presentationElementKind(element) {
  if (element instanceof NativePresentationObject) return "nativeObject";
  if (element instanceof ConnectorElement) return "connector";
  if (element instanceof GroupShape) return "groupShape";
  if (element instanceof TableElement) return "table";
  if (element instanceof ChartElement) return "chart";
  if (element instanceof ImageElement) return "image";
  return "shape";
}

function presentationSlideElements(slide) {
  const direct = [...slide.elements.items];
  return direct.flatMap((element) => element instanceof GroupShape ? element.allElements() : [element]);
}

function directSlideModelElements(slide) {
  return [...slide.elements.items];
}

function removePendingCloneDirectElement(slide, element) {
  const collections = [
    slide.connectors,
    slide.shapes,
    slide.tables,
    slide.charts,
    slide.images,
    slide.groups,
    slide.nativeObjects,
  ];
  for (const collection of collections) {
    const index = collection.items.indexOf(element);
    if (index < 0) continue;
    collection.items.splice(index, 1);
    const sceneIndex = slide.elements.items.indexOf(element);
    if (sceneIndex >= 0) slide.elements.items.splice(sceneIndex, 1);
    return;
  }
  throw new Error(`Presentation clone element ${element?.id || "<unknown>"} does not belong to its slide.`);
}
