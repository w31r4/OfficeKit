import { create, fromBinary, toBinary, toJsonString } from "@bufbuild/protobuf";
import { createHash } from "node:crypto";
import { ChartElement, GroupShape, ImageElement, Presentation, Shape, Slide, TableElement } from "../presentation/index.mjs";
import {
  ArtifactFamily,
  PresentationCustomGeometryPath_FillMode,
  PresentationArtifactSchema,
  PresentationDiagramTextNodeSchema,
  PresentationElementSchema,
  PresentationElementSourceBindingSchema,
  PresentationModernCommentAnchor_Kind,
  PresentationSlideSchema,
  PresentationSlideGuide_Orientation,
  PresentationTextBodySchema,
  PresentationTextRunSchema,
} from "../generated/office_kit/artifact/v1/office_artifact_pb.js";
import { normalizePresentationRunLink } from "../presentation/ooxml-hyperlinks.mjs";
import { planPresentationCustomShows } from "../presentation/ooxml-custom-shows.mjs";
import { planPresentationSections } from "../presentation/ooxml-sections.mjs";
import { normalizePresentationTransition, PRESENTATION_TRANSITION_CAPABILITY } from "../presentation/ooxml-transitions.mjs";
import { normalizePresentationAnimation, normalizePresentationMorph, PRESENTATION_ANIMATIONS_CAPABILITY, PRESENTATION_MORPH_CAPABILITY } from "../presentation/ooxml-animations.mjs";
import { deterministicPresentationGuid } from "../presentation/ooxml-modern-comments.mjs";
import { normalizePresentationThemeConfig } from "../presentation/ooxml-theme.mjs";
import { normalizePresentationTextBodyProperties } from "../presentation/text-body-properties.mjs";
import { effectivePresentationImageCrop, presentationImageCropFromWire, presentationImageCropToWire } from "../presentation/image-crop.mjs";
import { normalizePresentationCustomAdjustmentHandles, normalizePresentationCustomConnectionSites, normalizePresentationCustomPaths, normalizePresentationCustomTextRectangle } from "../presentation/custom-geometry.mjs";
import { normalizePresentationCustomGeometryFormulaGraph } from "../presentation/custom-geometry-formulas.mjs";
import { isPresentationAutoNumberType, normalizePresentationParagraphs, normalizePresentationParagraphStyles } from "../presentation/text-paragraphs.mjs";
import { normalizePresentationLineStyle, presentationLineColor } from "../presentation/line-styles.mjs";
import { normalizePresentationAccessibility } from "../presentation/accessibility.mjs";
import { resolveColorToken } from "../shared/colors.mjs";
import { createPresentationAssetCatalog, validatePictureBulletUri } from "./office-kit-assets.mjs";
import { OfficeKitCodecError } from "./office-kit-error.mjs";
import { modelPresentationChartFromWire, presentationChartToWire } from "./office-kit-presentation-charts.mjs";
import { materializePresentationNativeGraphs, presentationNativeGraphSnapshot } from "./office-kit-presentation-native.mjs";

const EMU_PER_PIXEL = 9525;
const EMU_PER_POINT = 12700;
const POINTS_PER_PIXEL = 0.75;
const MAX_FONT_SIZE_PIXELS = 1024;
const MAX_PARAGRAPH_COORDINATE_EMU = 51_206_400;
const MAX_TEXT_BODY_INSET_EMU = 2_147_483_647;
const ROTATION_UNITS_PER_DEGREE = 60_000;
const MAX_PARAGRAPH_SPACING_POINTS = 1584;
const MAX_PARAGRAPH_SPACING_MULTIPLIER = 132;
const PRESENTATION_STATE = Symbol.for("office-kit.presentation-state");
const PRESENTATION_SLIDE_DUPLICATOR = Symbol.for("office-kit.presentation-duplicate");
const PRESENTATION_SPEAKER_NOTES_CAPABILITY = Symbol.for("office-kit.speaker-notes-capability");
const PRESENTATION_LEGACY_COMMENTS_CAPABILITY = Symbol.for("office-kit.legacy-comments-capability");
const PRESENTATION_SLIDE_VISIBILITY_CAPABILITY = Symbol.for("office-kit.slide-visibility-capability");
const PRESENTATION_SLIDE_DELETION_CAPABILITY = Symbol.for("office-kit.slide-deletion-capability");
const PRESENTATION_SLIDE_CLONE_CAPABILITY = Symbol.for("office-kit.slide-clone-capability");
const PRESENTATION_SLIDE_CONTINUATION_CAPABILITY = Symbol.for("office-kit.slide-continuation-capability");
const PRESENTATION_ELEMENT_DELETION_CAPABILITY = Symbol.for("office-kit.presentation-element-deletion-capability");
const PRESENTATION_ELEMENT_DELETED = Symbol.for("office-kit.presentation-element-deleted");
const PRESENTATION_ELEMENT_ORDER_CAPABILITY = Symbol.for("office-kit.presentation-element-order-capability");
const PRESENTATION_NATIVE_LEAF_CAPABILITY = Symbol.for("office-kit.presentation-native-leaf-capability");
const PRESENTATION_COMPONENT_CAPABILITY = Symbol.for("office-kit.presentation-component-capability");
const PRESENTATION_IMPORTED_GROUP_CHILD = Symbol.for("office-kit.presentation-imported-group-child");
const PRESENTATION_IMAGE_DATA_URL_SOURCE = Symbol.for("office-kit.presentation-image-data-url-source");
const PRESENTATION_IMAGE_SVG_DATA_URL_SOURCE = Symbol.for("office-kit.presentation-image-svg-data-url-source");
const PRESENTATION_SCHEME_COLORS = new Set([
  "dk1", "lt1", "dk2", "lt2", "tx1", "bg1", "tx2", "bg2",
  "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink",
]);
const NATIVE_SCHEME_COLOR_CANONICAL = Object.freeze(Object.fromEntries(
  [...PRESENTATION_SCHEME_COLORS].map((token) => [token.toLowerCase(), token]),
));
const SOURCE_FREE_LAYOUT_TYPES = new Map([
  ["blank", "blank"],
  ["title", "title"],
  ["titleOnly", "titleOnly"],
  ["title-only", "titleOnly"],
  ["obj", "obj"],
  ["object", "obj"],
  ["content", "obj"],
  ["titleAndContent", "obj"],
  ["title-and-content", "obj"],
]);
const SOURCE_FREE_TEXT_PLACEHOLDER_TYPES = new Set(["title", "body", "ctrTitle", "subTitle"]);
const DEFAULT_PRESENTATION_THEME = JSON.stringify(normalizePresentationThemeConfig({}));
const RUN_STYLE_KEYS = new Set(["bold", "italic", "fontSize", "fontFamily", "fontFamilyEastAsia", "color"]);
const TEXT_FRAME_PARAGRAPH_KEYS = new Set(["alignment", "tabStops", "marginLeft", "indent", "lineSpacing", "spaceBefore", "spaceBeforePercent", "spaceAfter", "spaceAfterPercent"]);
const CUSTOM_TEXT_RECTANGLE_FIELDS = Object.freeze([
  Object.freeze(["left", "leftEmu", "leftReference"]),
  Object.freeze(["top", "topEmu", "topReference"]),
  Object.freeze(["right", "rightEmu", "rightReference"]),
  Object.freeze(["bottom", "bottomEmu", "bottomReference"]),
]);
const PARAGRAPH_KEYS = new Set([
  "runs", "level", "alignment", "style", "bulletCharacter", "autoNumber", "bulletImage", "bulletNone",
  "bulletFont", "bulletFontFollowText", "bulletColor", "bulletColorFollowText",
  "bulletSize", "bulletSizePercent", "bulletSizeFollowText", "tabStops", "marginLeft", "indent",
  "lineSpacing", "spaceBefore", "spaceBeforePercent", "spaceAfter", "spaceAfterPercent",
]);
function modelPresentationSlideGuides(viewProperties) {
  return (viewProperties?.slideGuides || []).map((guide) => ({
    orientation: guide.orientation === PresentationSlideGuide_Orientation.VERTICAL ? "vertical" : "horizontal",
    position: Number(guide.position),
  }));
}

function modelPresentationViewSourceBinding(binding) {
  if (!binding) return undefined;
  return {
    partPath: binding.partPath,
    relationshipId: binding.relationshipId,
    viewXmlSha256: binding.viewXmlSha256,
    semanticSha256: binding.semanticSha256,
    residualSha256: binding.residualSha256,
    editable: binding.editable === true,
  };
}

function modelPresentationView(viewProperties) {
  if (!viewProperties) return undefined;
  return {
    ...(viewProperties.gridSpacingCxEmu === undefined ? {} : { gridSpacingCxEmu: Number(viewProperties.gridSpacingCxEmu) }),
    ...(viewProperties.gridSpacingCyEmu === undefined ? {} : { gridSpacingCyEmu: Number(viewProperties.gridSpacingCyEmu) }),
    ...(viewProperties.slideViewSnapToGrid === undefined ? {} : { slideViewSnapToGrid: viewProperties.slideViewSnapToGrid }),
    ...(viewProperties.slideViewSnapToObjects === undefined ? {} : { slideViewSnapToObjects: viewProperties.slideViewSnapToObjects }),
    slideViewShowGuides: false,
    slideGuides: modelPresentationSlideGuides(viewProperties),
    ...(viewProperties.source ? { source: modelPresentationViewSourceBinding(viewProperties.source) } : {}),
  };
}

function samePresentationViewSourceBinding(left, right) {
  if (!left || !right) return false;
  return String(left.partPath || "").toLowerCase() === String(right.partPath || "").toLowerCase() &&
    String(left.relationshipId || "") === String(right.relationshipId || "") &&
    String(left.viewXmlSha256 || "").toLowerCase() === String(right.viewXmlSha256 || "").toLowerCase() &&
    String(left.semanticSha256 || "").toLowerCase() === String(right.semanticSha256 || "").toLowerCase() &&
    String(left.residualSha256 || "").toLowerCase() === String(right.residualSha256 || "").toLowerCase() &&
    left.editable === right.editable;
}

function presentationViewPropertiesForEnvelope(presentation, state) {
  const source = state?.viewProperties;
  if (!source) return undefined;
  const model = presentation.view.toProto();
  const binding = presentation.view._sourceBindingForExport();
  if (!model || !binding || !samePresentationViewSourceBinding(binding, source.source)) {
    throw new OfficeKitCodecError("Presentation view properties no longer match their imported source binding.", [], { code: "presentation_view_source_binding_mismatch" });
  }
  const optionalFields = ["gridSpacingCxEmu", "gridSpacingCyEmu", "slideViewSnapToGrid", "slideViewSnapToObjects"];
  if (optionalFields.some((field) => Object.hasOwn(model, field) !== (source[field] !== undefined)) ||
      model.slideGuides.length !== (source.slideGuides || []).length ||
      model.slideGuides.some((guide, index) => guide.orientation !== modelPresentationSlideGuides(source)[index]?.orientation)) {
    throw new OfficeKitCodecError("Presentation view properties must retain their imported grid/snap field presence and guide topology.", [], { code: "presentation_view_topology_changed" });
  }
  return {
    ...(Object.hasOwn(model, "gridSpacingCxEmu") ? { gridSpacingCxEmu: BigInt(model.gridSpacingCxEmu) } : {}),
    ...(Object.hasOwn(model, "gridSpacingCyEmu") ? { gridSpacingCyEmu: BigInt(model.gridSpacingCyEmu) } : {}),
    ...(Object.hasOwn(model, "slideViewSnapToGrid") ? { slideViewSnapToGrid: model.slideViewSnapToGrid } : {}),
    ...(Object.hasOwn(model, "slideViewSnapToObjects") ? { slideViewSnapToObjects: model.slideViewSnapToObjects } : {}),
    slideGuides: model.slideGuides.map((guide) => ({
      orientation: guide.orientation === "vertical" ? PresentationSlideGuide_Orientation.VERTICAL : PresentationSlideGuide_Orientation.HORIZONTAL,
      position: guide.position,
    })),
    source: { ...source.source },
  };
}

function assertTrustedPresentationState(state) {
  if (!state) return;
  const sourceHash = String(state.source?.packageSha256 || "").toLowerCase();
  const snapshot = state.opaqueOpc?.sourcePackage;
  const snapshotHash = String(snapshot?.sha256 || "").toLowerCase();
  if (!sourceHash || !snapshotHash || sourceHash !== snapshotHash || !snapshot?.data?.length) {
    throw new OfficeKitCodecError("PPTX source-bound export requires its validated source package snapshot.", [], { code: "missing_source_package" });
  }
}

// A source-bound slide stays attached to its imported SlidePart by object
// identity, not by whichever array index it happens to occupy now. Clone
// instances deliberately live in their own map: their source points at an
// origin Part, but they must never masquerade as a second binding to it.
function presentationSourceSlideStateMap(presentation, state) {
  if (!state) return undefined;
  const sourceBySlide = new Map();
  for (const sourceState of state.slides || []) {
    if (!(sourceState?.slide instanceof Slide) || sourceState.slide.presentation !== presentation || sourceBySlide.has(sourceState.slide)) {
      throw new OfficeKitCodecError("Imported presentation source bindings are invalid or ambiguous.", [], { code: "presentation_topology_changed" });
    }
    sourceBySlide.set(sourceState.slide, sourceState);
  }
  const cloneBySlide = new Map();
  for (const cloneState of state.clones || []) {
    if (!(cloneState?.slide instanceof Slide) || cloneState.slide.presentation !== presentation ||
        !sourceBySlide.has(cloneState.source?.slide) || cloneBySlide.has(cloneState.slide) || sourceBySlide.has(cloneState.slide)) {
      throw new OfficeKitCodecError("Imported presentation clone bindings are invalid or ambiguous.", [], { code: "presentation_topology_changed" });
    }
    cloneBySlide.set(cloneState.slide, cloneState);
  }
  if (presentation.slides.items.some((slide) => !sourceBySlide.has(slide) && !cloneBySlide.has(slide))) {
    throw new OfficeKitCodecError("Source-preserving PPTX export does not accept newly added slides. Use a supported imported-slide clone operation or a source-free presentation.", [], { code: "presentation_topology_changed" });
  }
  return { sourceBySlide, cloneBySlide };
}

function clonedPresentationValue(value) {
  return value === undefined ? undefined : structuredClone(value);
}

function isPresentationConnectorElement(element) {
  return element?.kind === "connector" && typeof element.id === "string";
}

function createPresentationCloneContext() {
  return {
    cloneIdBySourceId: new Map(),
    sourceIdByCloneId: new Map(),
    pendingConnectors: [],
  };
}

// The pending clone has fresh public model IDs, whereas its first export must
// still prove equality with the origin's source-bound wire. Keep that identity
// translation private to the clone transaction instead of leaking native IDs
// through the Agent-facing model.
function registerPresentationCloneElement(context, source, clone) {
  const sourceId = String(source?.id || "");
  const cloneId = String(clone?.id || "");
  if (!sourceId || !cloneId || context.cloneIdBySourceId.has(sourceId) || context.sourceIdByCloneId.has(cloneId)) {
    throw new OfficeKitCodecError("Imported presentation clone element identities are invalid or ambiguous.", [], { code: "unsupported_presentation_slide_clone" });
  }
  context.cloneIdBySourceId.set(sourceId, cloneId);
  context.sourceIdByCloneId.set(cloneId, sourceId);
  return clone;
}

function cloneImportedPresentationShape(container, source, context) {
  const clone = container.shapes.add({
    name: source.name,
    geometry: source.geometry,
    ...(source.customAdjustments?.length ? { customAdjustments: clonedPresentationValue(source.customAdjustments) } : {}),
    ...(source.customGuides?.length ? { customGuides: clonedPresentationValue(source.customGuides) } : {}),
    ...(source.customConnectionSites?.length ? { customConnectionSites: clonedPresentationValue(source.customConnectionSites) } : {}),
    ...(source.customAdjustmentHandles?.length ? { customAdjustmentHandles: clonedPresentationValue(source.customAdjustmentHandles) } : {}),
    ...(source.customPaths?.length ? { customPaths: clonedPresentationValue(source.customPaths) } : {}),
    ...(source.textRectangle ? { textRectangle: clonedPresentationValue(source.textRectangle) } : {}),
    position: clonedPresentationValue(source.position),
    ...(source.transform ? { transform: clonedPresentationValue(source.transform) } : {}),
    fill: clonedPresentationValue(source.fill),
    line: clonedPresentationValue(source.line),
    ...(source.borderRadius === undefined ? {} : { borderRadius: source.borderRadius }),
    ...(source.shadow ? { shadow: clonedPresentationValue(source.shadow) } : {}),
    ...(source.placeholder ? { placeholder: clonedPresentationValue(source.placeholder) } : {}),
    ...(source.accessibility ? { accessibility: clonedPresentationValue(source.accessibility) } : {}),
    _officeKitAccessibilityEditable: source.accessibilityCapability.editable,
    ...(source.useBackgroundFill === undefined ? {} : { _officeKitUseBackgroundFill: source.useBackgroundFill }),
    text: clonedPresentationValue(source.text.paragraphs),
    textBodyProperties: clonedPresentationValue(source.text.bodyProperties),
    textStyle: clonedPresentationValue(source.text.style),
  });
  clone.text.inheritedParagraphStyles = clonedPresentationValue(source.text.inheritedParagraphStyles);
  return registerPresentationCloneElement(context, source, clone);
}

// A clone needs a fresh model object so it cannot share mutable JavaScript
// identity with its origin. Its embedded asset stays content-addressed, and
// the native exporter deliberately shares that immutable ImagePart through a
// new relationship on the clone SlidePart.
function cloneImportedPresentationImage(container, source, context) {
  const clone = container.images.add({
    name: source.name,
    position: clonedPresentationValue(source.position),
    ...(source.accessibility ? { accessibility: clonedPresentationValue(source.accessibility) } : {}),
    _officeKitAccessibilityEditable: source.accessibilityCapability.editable,
    dataUrl: source.dataUrl,
    ...(source.contentType ? { contentType: source.contentType } : {}),
    ...(source.svgDataUrl ? { svgDataUrl: source.svgDataUrl } : {}),
    fit: source.fit,
    ...(source.crop ? { crop: clonedPresentationValue(source.crop) } : {}),
    geometry: source.geometry,
    ...(source.transform ? { transform: clonedPresentationValue(source.transform) } : {}),
  });
  return registerPresentationCloneElement(context, source, clone);
}

// Canonical tables are an accepted GraphicFrame leaf whose bounded DrawingML
// payload is inline in the SlidePart, so this creates a fresh model without
// copying an OPC part or relationship.
function cloneImportedPresentationTable(container, source, context) {
  const clone = container.tables.add({
    name: source.name,
    position: clonedPresentationValue(source.position),
    rows: source.rows,
    columns: source.columns,
    values: clonedPresentationValue(source.values),
    ...(source.style === undefined ? {} : { style: clonedPresentationValue(source.style) }),
    ...(source.styleOptions === undefined ? {} : { styleOptions: clonedPresentationValue(source.styleOptions) }),
    mergeRanges: clonedPresentationValue(source.mergeRanges),
    ...(source.accessibility ? { accessibility: clonedPresentationValue(source.accessibility) } : {}),
    _officeKitAccessibilityEditable: source.accessibilityCapability.editable,
  });
  if (source.border !== undefined) clone.border = clonedPresentationValue(source.border);
  return registerPresentationCloneElement(context, source, clone);
}

// A recognized literal-data chart is the one accepted relationship-owning
// GraphicFrame leaf. The JavaScript model must be independent immediately;
// OfficeKit then copies the verified closed ChartPart into a distinct OPC
// part so origin and clone can be edited independently after export/reimport.
function cloneImportedPresentationChart(container, source, context) {
  if (source.externalData) {
    throw new OfficeKitCodecError("The bounded imported-slide clone profile does not accept charts with embedded or external workbook data.", [], { code: "unsupported_presentation_slide_clone" });
  }
  const clone = container.charts.add(source.chartType, {
    name: source.name,
    position: clonedPresentationValue(source.position),
    title: source.title,
    categories: clonedPresentationValue(source.categories),
    series: clonedPresentationValue(source.series),
    axes: clonedPresentationValue(source.axes),
    legend: clonedPresentationValue(source.legend),
    dataLabels: clonedPresentationValue(source.dataLabels),
    ...(source.accessibility ? { accessibility: clonedPresentationValue(source.accessibility) } : {}),
    _officeKitAccessibilityEditable: source.accessibilityCapability.editable,
    ...(source.styleId === undefined ? {} : { styleId: source.styleId }),
    varyColors: source.varyColors,
    barOptions: clonedPresentationValue(source.barOptions),
    lineOptions: clonedPresentationValue(source.lineOptions),
  });
  return registerPresentationCloneElement(context, source, clone);
}

const CLONE_DIAGRAM_RELATIONSHIPS = new Map([
  ["dm", "/diagramData"],
  ["lo", "/diagramLayout"],
  ["qs", "/diagramQuickStyle"],
  ["cs", "/diagramColors"],
]);
const CLONE_DIAGRAM_CONTENT_TYPES = new Set([
  "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
  "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
  "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
  "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
]);
const CLONE_INK_CONTENT_TYPE = "application/inkml+xml";
const CLONE_MP4_CONTENT_TYPE = "video/mp4";
const CLONE_VIDEO_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
const CLONE_MEDIA_RELATIONSHIP = "http://schemas.microsoft.com/office/2007/relationships/media";
const CLONE_IMAGE_RELATIONSHIPS = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/image",
]);
const CLONE_CUSTOM_XML_RELATIONSHIPS = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml",
]);
const CLONE_RELATIONSHIP_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);

function isCloneDiagramGraphicFrame(source) {
  return /^<(?:[A-Za-z_][\w.-]*:)?graphicFrame(?:\s|>)/.test(String(source?.rawXml || "").trimStart());
}

function cloneDiagramReferenceIds(source) {
  const references = source?.relationshipReferences;
  if (!Array.isArray(references) || references.length !== 4) return undefined;
  const ids = new Map();
  for (const reference of references) {
    const localName = String(reference.attribute || "").split(":").at(-1);
    const id = String(reference.id ?? reference.relationshipId ?? "");
    if (!CLONE_DIAGRAM_RELATIONSHIPS.has(localName) || !CLONE_RELATIONSHIP_NAMESPACES.has(String(reference.namespaceUri || "")) ||
        !id || [...ids.values()].includes(id) || ids.has(localName)) return undefined;
    ids.set(localName, id);
  }
  return ids.size === 4 ? ids : undefined;
}

function cloneablePresentationDiagramWire(source) {
  const ids = source?.nativeKind === "diagram" ? cloneDiagramReferenceIds(source) : undefined;
  const paths = source?.preservedPartPaths;
  return Boolean(ids && isCloneDiagramGraphicFrame(source) && Array.isArray(paths) && paths.length === 4 && new Set(paths.map(String)).size === 4);
}

function cloneablePresentationDiagramModel(source) {
  const ids = source?.kind === "nativeObject" && source.nativeKind === "diagram" ? cloneDiagramReferenceIds(source) : undefined;
  if (!ids || !isCloneDiagramGraphicFrame(source) || source.oleWorkbook || source.oleOfficePackage || !Array.isArray(source.rootRelationships) || source.rootRelationships.length !== 4 ||
      !Array.isArray(source.parts) || source.parts.length !== 4) return false;
  const roots = new Map(source.rootRelationships.map((relationship) => [relationship.id, relationship]));
  if (roots.size !== 4) return false;
  for (const [localName, id] of ids) {
    const relationship = roots.get(id);
    if (!relationship || String(relationship.targetMode || "").toLowerCase() === "external" ||
        !String(relationship.type || "").endsWith(CLONE_DIAGRAM_RELATIONSHIPS.get(localName))) return false;
  }
  const contentTypes = new Set();
  const paths = new Set();
  for (const part of source.parts) {
    if (!CLONE_DIAGRAM_CONTENT_TYPES.has(part.contentType) || contentTypes.has(part.contentType) ||
        !part.path || paths.has(part.path) || !part.bytes?.length || !/^[0-9a-f]{64}$/i.test(String(part.sourceSha256 || "")) ||
        !Array.isArray(part.relationships) || part.relationships.length !== 0) return false;
    contentTypes.add(part.contentType);
    paths.add(part.path);
  }
  return contentTypes.size === 4;
}

function isCloneInkContentPart(source) {
  return /^<(?:[A-Za-z_][\w.-]*:)?contentPart(?:\s|>)/.test(String(source?.rawXml || "").trimStart());
}

function cloneInkContentReference(source) {
  const references = source?.relationshipReferences;
  if (!Array.isArray(references) || references.length !== 1) return undefined;
  const reference = references[0];
  return String(reference.attribute || "").split(":").at(-1) === "id" &&
    CLONE_RELATIONSHIP_NAMESPACES.has(String(reference.namespaceUri || "")) &&
    String(reference.id ?? reference.relationshipId ?? "")
    ? String(reference.id ?? reference.relationshipId)
    : undefined;
}

function cloneablePresentationInkContentWire(source) {
  const relationshipId = source?.nativeKind === "contentPart" ? cloneInkContentReference(source) : undefined;
  return Boolean(relationshipId && isCloneInkContentPart(source) && Array.isArray(source.preservedPartPaths) &&
    source.preservedPartPaths.length === 1 && source.preservedPartPaths[0]);
}

function cloneablePresentationInkContentModel(source) {
  const relationshipId = source?.kind === "nativeObject" && source.nativeKind === "contentPart"
    ? cloneInkContentReference(source)
    : undefined;
  if (!relationshipId || !isCloneInkContentPart(source) || source.oleWorkbook || source.oleOfficePackage ||
      !Array.isArray(source.rootRelationships) || source.rootRelationships.length !== 1 ||
      !Array.isArray(source.parts) || source.parts.length !== 1) return false;
  const relationship = source.rootRelationships[0];
  const part = source.parts[0];
  return relationship.id === relationshipId &&
    String(relationship.targetMode || "").toLowerCase() !== "external" &&
    CLONE_CUSTOM_XML_RELATIONSHIPS.has(String(relationship.type || "")) &&
    part.contentType === CLONE_INK_CONTENT_TYPE && Boolean(part.path) && Boolean(part.bytes?.length) &&
    /^[0-9a-f]{64}$/i.test(String(part.sourceSha256 || "")) &&
    Array.isArray(part.relationships) && part.relationships.length === 0;
}

function isCloneMediaPicture(source) {
  return /^<(?:[A-Za-z_][\w.-]*:)?pic(?:\s|>)/.test(String(source?.rawXml || "").trimStart());
}

function cloneMediaXmlTags(rawXml, localName) {
  return [...String(rawXml || "").matchAll(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>`, "gi"))]
    .map((match) => match[0]);
}

function cloneMediaXmlTag(rawXml, localName) {
  const matches = cloneMediaXmlTags(rawXml, localName);
  return matches.length === 1 ? matches[0] : undefined;
}

function cloneMediaXmlAttribute(tag, localName) {
  return new RegExp(`\\s(?:[A-Za-z_][\\w.-]*:)?${localName}="([^"]*)"`, "i").exec(String(tag || ""))?.[1];
}

function cloneablePresentationMediaMarkup(source, ids) {
  const rawXml = String(source?.rawXml || "");
  const click = cloneMediaXmlTag(rawXml, "hlinkClick");
  const video = cloneMediaXmlTag(rawXml, "videoFile");
  const media = cloneMediaXmlTag(rawXml, "media");
  const extensions = cloneMediaXmlTags(rawXml, "ext")
    .filter((tag) => cloneMediaXmlAttribute(tag, "uri") !== undefined);
  const extension = extensions.length === 1 &&
    cloneMediaXmlAttribute(extensions[0], "uri") === "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"
    ? extensions[0]
    : undefined;
  const blip = cloneMediaXmlTag(rawXml, "blip");
  if (!click || !video || !media || !extension || !blip || /<(?:[A-Za-z_][\w.-]*:)?audioFile\b/i.test(rawXml) ||
      cloneMediaXmlAttribute(click, "id") !== "" || cloneMediaXmlAttribute(click, "action") !== "ppaction://media" ||
      cloneMediaXmlAttribute(extension, "uri") !== "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" ||
      cloneMediaXmlAttribute(video, "link") !== ids.link) return false;
  const mediaId = cloneMediaXmlAttribute(media, "embed");
  const posterId = cloneMediaXmlAttribute(blip, "embed");
  return Boolean(mediaId && posterId && mediaId !== posterId && ids.embeds.includes(mediaId) && ids.embeds.includes(posterId));
}

function cloneMediaReferenceIds(source) {
  const references = source?.relationshipReferences;
  if (!Array.isArray(references) || references.length !== 3) return undefined;
  const byAttribute = new Map();
  const seenIds = new Set();
  for (const reference of references) {
    const attribute = String(reference.attribute || "").split(":").at(-1);
    const id = String(reference.id ?? reference.relationshipId ?? "");
    if (!CLONE_RELATIONSHIP_NAMESPACES.has(String(reference.namespaceUri || "")) || !id || seenIds.has(id) || !new Set(["link", "embed"]).has(attribute)) return undefined;
    seenIds.add(id);
    const values = byAttribute.get(attribute) || [];
    values.push(id);
    byAttribute.set(attribute, values);
  }
  const links = byAttribute.get("link") || [];
  const embeds = byAttribute.get("embed") || [];
  return links.length === 1 && embeds.length === 2 ? { link: links[0], embeds } : undefined;
}

function cloneablePresentationMediaWire(source) {
  const ids = source?.nativeKind === "media" ? cloneMediaReferenceIds(source) : undefined;
  return Boolean(ids && isCloneMediaPicture(source) && cloneablePresentationMediaMarkup(source, ids) && Array.isArray(source.preservedPartPaths) &&
    source.preservedPartPaths.length === 2 && new Set(source.preservedPartPaths.map(String)).size === 2);
}

function cloneablePresentationMediaModel(source) {
  const ids = source?.kind === "nativeObject" && source.nativeKind === "media" ? cloneMediaReferenceIds(source) : undefined;
  if (!ids || !isCloneMediaPicture(source) || !cloneablePresentationMediaMarkup(source, ids) || source.oleWorkbook || source.oleOfficePackage || !Array.isArray(source.rootRelationships) ||
      source.rootRelationships.length !== 3 || !Array.isArray(source.parts) || source.parts.length !== 2) return false;
  const relationships = new Map(source.rootRelationships.map((relationship) => [relationship.id, relationship]));
  if (relationships.size !== 3) return false;
  const video = relationships.get(ids.link);
  const mediaId = ids.embeds.find((id) => relationships.get(id)?.type === CLONE_MEDIA_RELATIONSHIP);
  const imageId = ids.embeds.find((id) => CLONE_IMAGE_RELATIONSHIPS.has(relationships.get(id)?.type));
  const media = relationships.get(mediaId);
  const image = relationships.get(imageId);
  if (!video || video.type !== CLONE_VIDEO_RELATIONSHIP || !media || !image || mediaId === imageId ||
      String(video.targetMode || "").toLowerCase() === "external" || String(media.targetMode || "").toLowerCase() === "external" ||
      String(image.targetMode || "").toLowerCase() === "external" || !video.target || video.target !== media.target || video.target === image.target) return false;
  const mp4Parts = source.parts.filter((part) => part.contentType === CLONE_MP4_CONTENT_TYPE && /^(?:ppt\/)?media\/[^/]+\.mp4$/i.test(String(part.path || "")));
  const posterParts = source.parts.filter((part) => /^image\/(?:png|jpeg)$/i.test(String(part.contentType || "")) && /^ppt\/media\/[^/]+$/i.test(String(part.path || "")));
  return mp4Parts.length === 1 && posterParts.length === 1 && source.parts.every((part) =>
    Boolean(part.bytes?.length) && /^[0-9a-f]{64}$/i.test(String(part.sourceSha256 || "")) &&
    Array.isArray(part.relationships) && part.relationships.length === 0);
}

// Eligible OLE, SmartArt, InkML, and embedded MP4 objects remain opaque PresentationML, but their
// package graphs have already been proved closed. Give the pending slide clone
// a fresh JavaScript object while retaining the exact source graph snapshot;
// C# allocates independent mutable parts during the first export.
function cloneImportedPresentationNativeObject(container, source, context) {
  if (source?.kind !== "nativeObject" || source._embeddedWorkbookReplacementBytes?.() ||
      source._embeddedOfficePackageReplacementBytes?.() || source._diagramTextReplacement?.()) {
    throw new OfficeKitCodecError("Imported-slide graph clone requires an unchanged source-bound native object.", [], { code: "unsupported_presentation_slide_clone" });
  }
  const clone = container.nativeObjects.add({
    name: source.name,
    nativeId: source.nativeId,
    creationId: source.creationId,
    nativeKind: source.nativeKind,
    text: source.text,
    position: clonedPresentationValue(source.position),
    rawXml: source.rawXml,
    sourcePart: source.sourcePart,
    relationshipReferences: clonedPresentationValue(source.relationshipReferences),
    rootRelationships: clonedPresentationValue(source.rootRelationships),
    parts: clonedPresentationValue(source.parts),
    placementCapability: clonedPresentationValue(source.placementCapability),
    oleWorkbook: clonedPresentationValue(source.oleWorkbook),
    oleOfficePackage: clonedPresentationValue(source.oleOfficePackage),
    diagramText: clonedPresentationValue(source._diagramTextSourceBinding?.()),
    nativeChart: clonedPresentationValue(source._nativeChartSourceBinding?.()),
  });
  return registerPresentationCloneElement(context, source, clone);
}

function cloneImportedPresentationConnector(container, source, context) {
  const clone = container.connectors.add({
    name: source.name,
    connectorType: source.connectorType,
    start: clonedPresentationValue(source.start),
    end: clonedPresentationValue(source.end),
    startSiteIndex: source.startSiteIndex,
    endSiteIndex: source.endSiteIndex,
    line: clonedPresentationValue(source.line),
    head: clonedPresentationValue(source.head),
    tail: clonedPresentationValue(source.tail),
    cap: source.cap,
    join: source.join,
    ...(source.accessibility ? { accessibility: clonedPresentationValue(source.accessibility) } : {}),
    _officeKitAccessibilityEditable: source.accessibilityCapability.editable,
    _officeKitSourceBound: true,
  });
  registerPresentationCloneElement(context, source, clone);
  context.pendingConnectors.push({ source, clone });
  return clone;
}

// A group is not automatically a clone leaf: it can contain connectors,
// native objects, or external edges. This recursive helper is called only
// after the source wire tree has proved every descendant is one of the same
// bounded clone-safe element kinds. A chart descendant is accepted only when
// the native preflight proves its ChartPart is closed. Each new model object
// has a fresh JS identity; the source bindings still make the pending clone
// immutable until its export/reimport boundary.
function cloneImportedPresentationGroup(container, source, context) {
  const clone = container.groups.add({
    name: source.name,
    position: clonedPresentationValue(source.position),
    childFrame: clonedPresentationValue(source.childFrame),
    ...(source.accessibility ? { accessibility: clonedPresentationValue(source.accessibility) } : {}),
    _officeKitAccessibilityEditable: source.accessibilityCapability.editable,
  });
  registerPresentationCloneElement(context, source, clone);
  for (const child of source.children) cloneImportedPresentationElement(clone, child, context);
  return clone;
}

function cloneImportedPresentationElement(container, source, context) {
  if (source instanceof Shape) return cloneImportedPresentationShape(container, source, context);
  if (source instanceof TableElement) return cloneImportedPresentationTable(container, source, context);
  if (source instanceof ChartElement) return cloneImportedPresentationChart(container, source, context);
  if (source instanceof ImageElement) return cloneImportedPresentationImage(container, source, context);
  if (isPresentationConnectorElement(source)) return cloneImportedPresentationConnector(container, source, context);
  if (source instanceof GroupShape) return cloneImportedPresentationGroup(container, source, context);
  if (source?.kind === "nativeObject") return cloneImportedPresentationNativeObject(container, source, context);
  throw new OfficeKitCodecError("The bounded imported-slide clone profile encountered an unsupported group descendant.", [], { code: "unsupported_presentation_slide_clone" });
}

function cloneSupportedPresentationContent(content, allowNativeGraphLeaf = true) {
  if (content?.case === "shape" || content?.case === "table" || content?.case === "chart" || content?.case === "image" || content?.case === "connector") return true;
  if (content?.case === "opaque") {
    return allowNativeGraphLeaf && Boolean(content.value?.rawXml);
  }
  if (content?.case !== "group") return false;
  const children = content.value?.children;
  return Array.isArray(children) && children.length > 0 && children.every((child) => cloneSupportedPresentationContent(child?.content, false));
}

function collectPresentationCloneSourceIds(source, ids, allowNativeGraphLeaf = true) {
  const cloneableNative = allowNativeGraphLeaf && source?.kind === "nativeObject" &&
    !source._embeddedWorkbookReplacementBytes?.() && !source._embeddedOfficePackageReplacementBytes?.() && !source._diagramTextReplacement?.();
  if (!(source instanceof Shape) && !(source instanceof TableElement) && !(source instanceof ChartElement) && !(source instanceof ImageElement) && !isPresentationConnectorElement(source) && !(source instanceof GroupShape) && !cloneableNative) {
    throw new OfficeKitCodecError("The bounded imported-slide clone profile encountered an unsupported source element.", [], { code: "unsupported_presentation_slide_clone" });
  }
  const id = String(source.id || "");
  if (!id || ids.has(id)) {
    throw new OfficeKitCodecError("Imported presentation clone source element identities are invalid or ambiguous.", [], { code: "unsupported_presentation_slide_clone" });
  }
  ids.add(id);
  if (source instanceof GroupShape) {
    for (const child of source.children) collectPresentationCloneSourceIds(child, ids, false);
  }
}

function assertPresentationCloneConnectorTargets(source, sourceIds) {
  if (isPresentationConnectorElement(source)) {
    for (const targetId of [source.startTargetId, source.endTargetId]) {
      if (targetId && !sourceIds.has(targetId)) {
        throw new OfficeKitCodecError("A bounded imported-slide connector may target only an element cloned in the same slide tree.", [], { code: "unsupported_presentation_slide_clone" });
      }
    }
  }
  if (source instanceof GroupShape) {
    for (const child of source.children) assertPresentationCloneConnectorTargets(child, sourceIds);
  }
}

function bindPresentationCloneConnectorTargets(context) {
  for (const { source, clone } of context.pendingConnectors) {
    const targetId = (value, side) => {
      if (!value) return undefined;
      const cloneTargetId = context.cloneIdBySourceId.get(value);
      if (!cloneTargetId) {
        throw new OfficeKitCodecError(`Imported presentation clone connector ${source.id} has an unresolved ${side} target.`, [], { code: "unsupported_presentation_slide_clone" });
      }
      return cloneTargetId;
    };
    clone.startTargetId = targetId(source.startTargetId, "start");
    clone.endTargetId = targetId(source.endTargetId, "end");
    clone.captureAttachedEndpointState?.();
  }
}

function capturePresentationConnectorEndpointState(element) {
  if (isPresentationConnectorElement(element)) element.captureAttachedEndpointState?.();
  if (element instanceof GroupShape) for (const child of element.children) capturePresentationConnectorEndpointState(child);
}

// A legacy comment has no JavaScript object identity that may be shared with
// its origin. Copy the imported thread record into a fresh slide model while
// retaining its native author/index evidence; the C# clone preflight then
// proves the clone-local comments XML and shared immutable author catalog are
// unchanged before writing any OPC graph.
function cloneImportedPresentationLegacyComments(slide, source) {
  for (const thread of source.comments.items) {
    const snapshot = clonedPresentationValue(thread.toJSON());
    slide.comments.addThread(undefined, snapshot.comments?.[0]?.text || "", snapshot);
  }
}

function duplicateImportedPresentationSlide(presentation, state, slide) {
  const source = (state.slides || []).find((entry) => entry.slide === slide);
  if (!source) {
    throw new OfficeKitCodecError("Only an original imported PPTX slide can be duplicated in this bounded clone profile.", [], { code: "unsupported_presentation_slide_clone" });
  }
  if ((state.clones || []).some((entry) => entry.source === source)) {
    throw new OfficeKitCodecError("The bounded imported-slide clone profile permits only one pending clone per origin; export and reimport it before cloning that source again.", [], { code: "unsupported_presentation_slide_clone" });
  }
  const capability = slide.cloneCapability;
  if (!capability.known || !capability.supported) {
    const detail = capability.blockedReason ? `: ${capability.blockedReason}` : ".";
    throw new OfficeKitCodecError(`Imported presentation slide cannot be safely cloned${detail}`, [], { code: "unsupported_presentation_slide_clone" });
  }
  if (source.entries.some((entry) => !cloneSupportedPresentationContent(entry.wire.content))) {
    throw new OfficeKitCodecError("Imported-slide graph clone encountered an unsupported semantic element tree.", [], { code: "unsupported_presentation_slide_clone" });
  }
  for (const entry of source.entries) {
    if (entry.wire.content.case === "opaque" && entry.snapshot !== opaquePresentationSnapshot(entry.model)) {
      throw new OfficeKitCodecError("Imported-slide graph clone requires every source-bound native object to remain unchanged before its first export.", [], { code: "unsupported_presentation_slide_clone" });
    }
  }
  const sourceIds = new Set();
  for (const entry of source.entries) collectPresentationCloneSourceIds(entry.model, sourceIds);
  for (const entry of source.entries) assertPresentationCloneConnectorTargets(entry.model, sourceIds);
  const clone = presentation.slides.insert({
    after: slide,
    name: slide.name,
    hidden: slide.hidden,
    ...(slide.background?.fill || slide.background?.image ? { background: clonedPresentationValue(slide.background) } : {}),
    ...(slide.transition?.configured ? { transition: slide.transition.toJSON() } : {}),
    ...(source.wire.speakerNotes
      ? { notes: source.wire.speakerNotes.textBody ? slide.speakerNotes?.textFrame?.paragraphs || [] : slide.speakerNotes?.text || "" }
      : {}),
  });
  Object.defineProperty(clone, PRESENTATION_SLIDE_VISIBILITY_CAPABILITY, {
    value: Object.freeze({ ...slide.visibilityCapability }),
  });
  Object.defineProperty(clone, PRESENTATION_SLIDE_CONTINUATION_CAPABILITY, {
    value: Object.freeze({
      sourceBound: true,
      ready: false,
      profile: "pending-clone",
      requiresExportReopen: true,
      sourceRevisionSha256: slide.cloneCapability.sourceRevisionSha256,
    }),
  });
  clone.layoutId = slide.layoutId;
  cloneImportedPresentationLegacyComments(clone, slide);
  const cloneContext = createPresentationCloneContext();
  const entries = source.entries.map((entry) => {
    const model = cloneImportedPresentationElement(clone, entry.model, cloneContext);
    return {
      wire: entry.wire,
      model,
      placeholderSnapshot: entry.wire.content.case === "shape" && entry.wire.content.value.placeholder
        ? slidePlaceholderState(model)
        : undefined,
      snapshot: entry.wire.content.case === "image"
        ? presentationImageReadOnlySnapshot(model)
        : entry.wire.content.case === "table"
          ? presentationTableReadOnlySnapshot(model)
          : entry.wire.content.case === "opaque"
            ? opaquePresentationSnapshot(model)
            : undefined,
      cloneModelSnapshot: presentationCloneElementSnapshot(model),
    };
  });
  bindPresentationCloneConnectorTargets(cloneContext);
  for (const entry of entries) entry.cloneModelSnapshot = presentationCloneElementSnapshot(entry.model);
  const cloneState = {
    source,
    slide: clone,
    name: clone.name,
    commentSnapshot: presentationSlideCommentSnapshot(clone),
    entries,
    sourceIdByCloneId: cloneContext.sourceIdByCloneId,
    allowedDeletedIds: undefined,
    componentReuse: undefined,
  };
  state.clones.push(cloneState);
  return clone;
}

function presentationCloneMessage(slide, omittedElementIds = undefined) {
  const { id: _id, source: _source, cloneSource: _cloneSource, $typeName: _typeName, elementDeletions: _elementDeletions, ...comparable } = slide;
  if (omittedElementIds?.size) {
    comparable.elements = (comparable.elements || []).filter((element) => !omittedElementIds.has(element.id));
  }
  return create(PresentationSlideSchema, comparable);
}

function presentationCloneMatches(requested, source, omittedElementIds = undefined) {
  // Compare the full typed protobuf projection, not its byte encoding. Buf
  // preserves unknown fields in their original wire order, while a bounded
  // clone reconstructs the same known fields in canonical order; byte equality
  // would therefore reject semantically identical clones after schema growth.
  return toJsonString(PresentationSlideSchema, presentationCloneMessage(requested, omittedElementIds)) ===
    toJsonString(PresentationSlideSchema, presentationCloneMessage(source, omittedElementIds));
}

function emuFromPixels(value, name, { allowNegative = false } = {}) {
  const number = Number(value);
  if (!Number.isFinite(number) || (!allowNegative && number < 0)) {
    throw new OfficeKitCodecError(`${name} must be a ${allowNegative ? "finite" : "non-negative finite"} number.`, [], { code: "invalid_presentation_frame" });
  }
  return BigInt(Math.round(number * EMU_PER_PIXEL));
}

// DrawingML permits a negative offset for source-bound bleed images. The
// authoring profile remains non-negative, while an imported image keeps its
// valid edge-bleed geometry so the semantic projection does not downgrade it
// to opaque solely because it crosses the slide boundary.
function sourceBoundFrameEmuFromPixels(value, name, original) {
  return emuFromPixels(value, name, { allowNegative: Boolean(original?.source) });
}

function signedEmuFromPixels(value, name) {
  const number = Number(value);
  if (!Number.isFinite(number) || Math.abs(number) > 10_000_000) throw new OfficeKitCodecError(`${name} must be a bounded finite number.`, [], { code: "invalid_presentation_group" });
  return BigInt(Math.round(number * EMU_PER_PIXEL));
}

function paragraphEmuFromPixels(value, name, { allowNegative = false } = {}) {
  const number = Number(value);
  const emu = Math.round(number * EMU_PER_PIXEL);
  if (!Number.isFinite(number) || (!allowNegative && number < 0) || emu < (allowNegative ? -MAX_PARAGRAPH_COORDINATE_EMU : 0) || emu > MAX_PARAGRAPH_COORDINATE_EMU) {
    throw new OfficeKitCodecError(`${name} is outside the supported DrawingML coordinate range.`, [], { code: "invalid_presentation_text" });
  }
  return BigInt(emu);
}

function presentationRgb(value, name) {
  const source = typeof value === "string" ? value : value?.color || value?.fill;
  const raw = resolveColorToken(source, source);
  if (raw == null || raw === "transparent" || raw === "none") return "";
  const match = /^#([0-9a-f]{6})$/i.exec(String(raw));
  if (!match) throw new OfficeKitCodecError(`${name} must be transparent or a six-digit RGB color.`, [], { code: "unsupported_presentation_features" });
  return match[1].toUpperCase();
}

function presentationFillOpacityThousandthPercent(value, name, fillRgb) {
  if (typeof value === "string" || value?.opacity == null) return undefined;
  const opacity = Number(value.opacity);
  if (!Number.isFinite(opacity) || opacity < 0 || opacity > 1) {
    throw new OfficeKitCodecError(`${name}.opacity must be a finite number from 0 through 1.`, [], { code: "invalid_presentation_fill" });
  }
  if (!fillRgb) {
    throw new OfficeKitCodecError(`${name}.opacity requires a solid RGB fill.`, [], { code: "invalid_presentation_fill" });
  }
  return Math.round(opacity * 100_000);
}

function modelPresentationShapeFill(shape) {
  const color = shape.fillScheme || (shape.fillRgb ? `#${shape.fillRgb}` : "transparent");
  return shape.fillOpacityThousandthPercent === undefined
    ? color
    : { color, opacity: Number(shape.fillOpacityThousandthPercent) / 100_000 };
}

function presentationChart(chart, original) {
  const result = presentationChartToWire(chart, original, {
    emuFromPixels,
    rgb: presentationRgb,
    sourceBoundFrameEmuFromPixels,
  });
  const accessibility = normalizePresentationAccessibility(chart.accessibility, `Presentation chart ${chart.id}`);
  if (accessibility) result.content.value.accessibility = accessibility;
  return result;
}

function modelPresentationChart(source, accessibilityEditable) {
  return {
    ...modelPresentationChartFromWire(source, EMU_PER_PIXEL),
    ...modelPresentationAccessibility(source.accessibility, "Imported Presentation chart"),
    ...(accessibilityEditable === undefined ? {} : { _officeKitAccessibilityEditable: accessibilityEditable === true }),
  };
}

function unsupportedStyleFields(style = {}) {
  return Object.keys(style).filter((key) => !RUN_STYLE_KEYS.has(key));
}

function presentationFontFamily(value, label) {
  if (value == null) return undefined;
  const family = String(value);
  if (!family.trim() || family.length > 255) {
    throw new OfficeKitCodecError(`${label} uses an invalid font family.`, [], { code: "invalid_presentation_text" });
  }
  return family;
}

function containsEastAsianText(value) {
  return /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}]/u.test(String(value ?? ""));
}

function wireTextStyle(style = {}, shapeId) {
  const unsupported = unsupportedStyleFields(style);
  if (unsupported.length) throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses unsupported paragraph text style fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const fontSize = style.fontSize == null ? undefined : Number(style.fontSize);
  if (fontSize !== undefined && (!Number.isFinite(fontSize) || fontSize <= 0 || fontSize > MAX_FONT_SIZE_PIXELS)) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses a paragraph font size outside the supported 0-${MAX_FONT_SIZE_PIXELS} pixel range.`, [], { code: "invalid_presentation_text" });
  }
  const fontFamily = presentationFontFamily(style.fontFamily, `Presentation shape ${shapeId} paragraph`);
  const fontFamilyEastAsia = presentationFontFamily(style.fontFamilyEastAsia, `Presentation shape ${shapeId} paragraph East Asian`);
  let color;
  if (style.color != null) {
    const token = String(style.color).trim();
    color = PRESENTATION_SCHEME_COLORS.has(token)
      ? { case: "colorScheme", value: token }
      : { case: "colorRgb", value: presentationRgb(style.color, `${shapeId}.text.paragraphStyle.color`) };
    if (!color.value) throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses a transparent paragraph color outside the PPTX NativeAOT text slice.`, [], { code: "unsupported_presentation_features" });
  }
  return {
    ...(style.bold == null ? {} : { bold: Boolean(style.bold) }),
    ...(style.italic == null ? {} : { italic: Boolean(style.italic) }),
    ...(fontSize === undefined ? {} : { fontSizePoints: fontSize * POINTS_PER_PIXEL }),
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(fontFamilyEastAsia === undefined ? {} : { fontFamilyEastAsia }),
    ...(color ? { color } : {}),
  };
}

function wireDefaultRunStyle(paragraph, original, shapeId) {
  if (Object.keys(paragraph.style || {}).length) {
    return { case: "defaultRunProperties", value: wireTextStyle(paragraph.style, shapeId) };
  }
  if (new Set(["defaultRunProperties", "noDefaultRunProperties"]).has(original?.defaultRunStyle?.case)) {
    return { case: "noDefaultRunProperties", value: true };
  }
  return undefined;
}

function wireHyperlink(value, original, shapeId, customShowLinks) {
  if (value == null) {
    if (new Set(["runHyperlink", "noHyperlink"]).has(original?.hyperlink?.case)) return { case: "noHyperlink", value: true };
    return undefined;
  }
  let link;
  try {
    link = normalizePresentationRunLink(value);
  } catch (error) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses an invalid run hyperlink: ${error.message}`, [], { code: "invalid_presentation_hyperlink" });
  }
  const target = link.uri
    ? { case: "uri", value: link.uri }
    : link.slideId
      ? { case: "slideId", value: link.slideId }
      : link.action
        ? { case: "action", value: link.action }
        : link.customShow
          ? { case: "customShowId", value: resolvePresentationCustomShowLinkId(link.customShow, original, shapeId, customShowLinks) }
        : undefined;
  if (!target) throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses an unsupported run hyperlink target.`, [], { code: "unsupported_presentation_features" });
  return {
    case: "runHyperlink",
    value: {
      target,
      ...(link.tooltip == null ? {} : { tooltip: link.tooltip }),
      ...(link.targetFrame == null ? {} : { targetFrame: link.targetFrame }),
      ...(link.history == null ? {} : { history: link.history }),
      ...(link.highlightClick == null ? {} : { highlightClick: link.highlightClick }),
      ...(link.returnToSlide == null ? {} : { returnToSlide: link.returnToSlide }),
    },
  };
}

function wireRun(run, inheritedStyle, shapeId, original, customShowLinks) {
  const unsupported = unsupportedStyleFields(run.style);
  if (unsupported.length) throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses unsupported run style fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const style = { ...inheritedStyle, ...(run.style || {}) };
  const fontSize = style.fontSize == null ? undefined : Number(style.fontSize);
  if (fontSize !== undefined && (!Number.isFinite(fontSize) || fontSize <= 0 || fontSize > MAX_FONT_SIZE_PIXELS)) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses a font size outside the supported 0-${MAX_FONT_SIZE_PIXELS} pixel range.`, [], { code: "invalid_presentation_text" });
  }
  const fontFamily = presentationFontFamily(style.fontFamily, `Presentation shape ${shapeId}`);
  const runText = run.field?.text ?? run.text ?? "";
  const explicitEastAsianFamily = presentationFontFamily(style.fontFamilyEastAsia, `Presentation shape ${shapeId} East Asian`);
  const fontFamilyEastAsia = explicitEastAsianFamily ?? (fontFamily && containsEastAsianText(runText) ? fontFamily : undefined);
  const colorRgb = style.color == null ? undefined : presentationRgb(style.color, `${shapeId}.text.color`);
  if (colorRgb === "") {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses a transparent run color outside the PPTX NativeAOT text slice.`, [], { code: "unsupported_presentation_features" });
  }
  const hyperlink = wireHyperlink(run.link, original, shapeId, customShowLinks);
  return {
    content: run.break
      ? { case: "lineBreak", value: true }
      : run.field
        ? { case: "field", value: { id: run.field.id, type: run.field.type, text: run.field.text } }
        : { case: "text", value: String(run.text ?? "") },
    ...(style.bold == null ? {} : { bold: Boolean(style.bold) }),
    ...(style.italic == null ? {} : { italic: Boolean(style.italic) }),
    ...(fontSize === undefined ? {} : { fontSizePoints: fontSize * POINTS_PER_PIXEL }),
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(fontFamilyEastAsia === undefined ? {} : { fontFamilyEastAsia }),
    ...(colorRgb === undefined ? {} : { colorRgb }),
    ...(hyperlink ? { hyperlink } : {}),
  };
}

function wireBullet(paragraph, original, shapeId, assetCatalog) {
  const choices = [paragraph.bulletCharacter != null, Boolean(paragraph.autoNumber), Boolean(paragraph.bulletImage), paragraph.bulletNone === true];
  if (choices.filter(Boolean).length > 1) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} paragraph selects more than one list marker.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletCharacter != null) {
    const character = String(paragraph.bulletCharacter);
    if ([...character].length !== 1) throw new OfficeKitCodecError(`Presentation shape ${shapeId} bullet character must contain one Unicode scalar value.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletCharacter", value: character };
  }
  if (paragraph.autoNumber) {
    const scheme = String(paragraph.autoNumber.type || paragraph.autoNumber.scheme || "");
    if (!isPresentationAutoNumberType(scheme)) throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses unsupported auto-number scheme ${scheme || "(missing)"}.`, [], { code: "invalid_presentation_text" });
    const rawStart = paragraph.autoNumber.startAt ?? paragraph.autoNumber.start;
    const startAt = rawStart == null ? undefined : Number(rawStart);
    if (startAt !== undefined && (!Number.isInteger(startAt) || startAt < 1 || startAt > 32767)) {
      throw new OfficeKitCodecError(`Presentation shape ${shapeId} auto-number start must be from 1 through 32767.`, [], { code: "invalid_presentation_text" });
    }
    return { case: "autoNumber", value: { scheme, ...(startAt === undefined ? {} : { startAt }) } };
  }
  if (paragraph.bulletImage) {
    if (paragraph.bulletImage.relationshipId) {
      throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses an unresolved picture-bullet relationship ID.`, [], { code: "invalid_presentation_asset" });
    }
    const source = paragraph.bulletImage.dataUrl
      ? { case: "assetId", value: assetCatalog.addDataUrl(paragraph.bulletImage.dataUrl) }
      : { case: "uri", value: validatePictureBulletUri(paragraph.bulletImage.uri) };
    return { case: "pictureBullet", value: { source } };
  }
  if (paragraph.bulletNone === true || new Set(["noBullet", "bulletCharacter", "autoNumber", "pictureBullet"]).has(original?.bullet?.case)) {
    return { case: "noBullet", value: true };
  }
  return undefined;
}

function wireBulletFont(paragraph, original, shapeId) {
  if (paragraph.bulletFont != null && paragraph.bulletFontFollowText === true) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} paragraph selects both a bullet font and follow-text font.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletFont != null) {
    const family = String(paragraph.bulletFont).trim();
    if (!family || family.length > 255) throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses an invalid bullet font family.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletFontFamily", value: family };
  }
  if (paragraph.bulletFontFollowText === true || new Set(["bulletFontFamily", "bulletFontFollowText"]).has(original?.bulletFont?.case)) {
    return { case: "bulletFontFollowText", value: true };
  }
  return undefined;
}

function wireBulletColor(paragraph, original, shapeId) {
  if (paragraph.bulletColor != null && paragraph.bulletColorFollowText === true) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} paragraph selects both a bullet color and follow-text color.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletColor != null) {
    const scheme = String(paragraph.bulletColor).trim();
    if (PRESENTATION_SCHEME_COLORS.has(scheme)) return { case: "bulletColorScheme", value: scheme };
    const rgb = presentationRgb(paragraph.bulletColor, `${shapeId}.text.bulletColor`);
    if (!rgb) throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses a transparent bullet color outside the PPTX NativeAOT text slice.`, [], { code: "unsupported_presentation_features" });
    return { case: "bulletColorRgb", value: rgb };
  }
  if (paragraph.bulletColorFollowText === true || new Set(["bulletColorRgb", "bulletColorScheme", "bulletColorFollowText"]).has(original?.bulletColor?.case)) {
    return { case: "bulletColorFollowText", value: true };
  }
  return undefined;
}

function wireBulletSize(paragraph, original, shapeId) {
  const choices = [paragraph.bulletSize != null, paragraph.bulletSizePercent != null, paragraph.bulletSizeFollowText === true];
  if (choices.filter(Boolean).length > 1) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} paragraph selects more than one bullet size.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletSize != null) {
    const pixels = Number(paragraph.bulletSize);
    if (!Number.isFinite(pixels) || pixels < 4 / 3 || pixels > MAX_FONT_SIZE_PIXELS) throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses an invalid fixed bullet size.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletSizePoints", value: pixels * POINTS_PER_PIXEL };
  }
  if (paragraph.bulletSizePercent != null) {
    const percent = Number(paragraph.bulletSizePercent);
    if (!Number.isFinite(percent) || percent < 0.25 || percent > 4) throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses an invalid percentage bullet size.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletSizePercent", value: percent };
  }
  if (paragraph.bulletSizeFollowText === true || new Set(["bulletSizePoints", "bulletSizePercent", "bulletSizeFollowText"]).has(original?.bulletSize?.case)) {
    return { case: "bulletSizeFollowText", value: true };
  }
  return undefined;
}

function wireTabStops(paragraph, original, shapeId) {
  if (paragraph.tabStops?.length) {
    return { tabStops: paragraph.tabStops.map((tab) => ({ positionEmu: emuFromPixels(tab.position, `${shapeId}.text.tabStops.position`), alignment: tab.alignment })) };
  }
  if (original?.tabStops?.length || original?.noTabStops === true) return { noTabStops: true };
  return {};
}

function wireParagraphLayout(paragraph, original, shapeId) {
  const leftMargin = paragraph.marginLeft != null
    ? { case: "marginLeftEmu", value: paragraphEmuFromPixels(paragraph.marginLeft, `${shapeId}.text.marginLeft`) }
    : new Set(["marginLeftEmu", "noMarginLeft"]).has(original?.leftMargin?.case)
      ? { case: "noMarginLeft", value: true }
      : undefined;
  const indentation = paragraph.indent != null
    ? { case: "indentEmu", value: paragraphEmuFromPixels(paragraph.indent, `${shapeId}.text.indent`, { allowNegative: true }) }
    : new Set(["indentEmu", "noIndent"]).has(original?.indentation?.case)
      ? { case: "noIndent", value: true }
      : undefined;
  return {
    ...(leftMargin ? { leftMargin } : {}),
    ...(indentation ? { indentation } : {}),
  };
}

function paragraphSpacingValue(value, name, { allowZero = true, maximum }) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < (allowZero ? 0 : Number.EPSILON) || number > maximum) {
    throw new OfficeKitCodecError(`${name} is outside the supported DrawingML spacing range.`, [], { code: "invalid_presentation_text" });
  }
  return number;
}

function paragraphSpacingPointsFromPixels(value, name, { allowZero = true } = {}) {
  const pixels = paragraphSpacingValue(value, name, { allowZero, maximum: MAX_PARAGRAPH_SPACING_POINTS / POINTS_PER_PIXEL });
  return pixels * POINTS_PER_PIXEL;
}

function wireParagraphSpacing(paragraph, original, shapeId) {
  let lineSpacing;
  if (paragraph.lineSpacing != null) {
    const value = Number(paragraph.lineSpacing);
    lineSpacing = value <= 10
      ? { case: "lineSpacingMultiplier", value: paragraphSpacingValue(value, `${shapeId}.text.lineSpacing`, { allowZero: false, maximum: MAX_PARAGRAPH_SPACING_MULTIPLIER }) }
      : { case: "lineSpacingPoints", value: paragraphSpacingPointsFromPixels(value, `${shapeId}.text.lineSpacing`, { allowZero: false }) };
  } else if (new Set(["lineSpacingPoints", "lineSpacingMultiplier", "noLineSpacing"]).has(original?.lineSpacing?.case)) {
    lineSpacing = { case: "noLineSpacing", value: true };
  }

  if (paragraph.spaceBefore != null && paragraph.spaceBeforePercent != null) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} paragraph must use either point or percentage space-before, not both.`, [], { code: "invalid_presentation_text" });
  }
  let spaceBefore;
  if (paragraph.spaceBefore != null) {
    spaceBefore = { case: "spaceBeforePoints", value: paragraphSpacingPointsFromPixels(paragraph.spaceBefore, `${shapeId}.text.spaceBefore`) };
  } else if (paragraph.spaceBeforePercent != null) {
    spaceBefore = { case: "spaceBeforeMultiplier", value: paragraphSpacingValue(paragraph.spaceBeforePercent, `${shapeId}.text.spaceBeforePercent`, { maximum: MAX_PARAGRAPH_SPACING_MULTIPLIER }) };
  } else if (new Set(["spaceBeforePoints", "spaceBeforeMultiplier", "noSpaceBefore"]).has(original?.spaceBefore?.case)) {
    spaceBefore = { case: "noSpaceBefore", value: true };
  }

  if (paragraph.spaceAfter != null && paragraph.spaceAfterPercent != null) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} paragraph must use either point or percentage space-after, not both.`, [], { code: "invalid_presentation_text" });
  }
  let spaceAfter;
  if (paragraph.spaceAfter != null) {
    spaceAfter = { case: "spaceAfterPoints", value: paragraphSpacingPointsFromPixels(paragraph.spaceAfter, `${shapeId}.text.spaceAfter`) };
  } else if (paragraph.spaceAfterPercent != null) {
    spaceAfter = { case: "spaceAfterMultiplier", value: paragraphSpacingValue(paragraph.spaceAfterPercent, `${shapeId}.text.spaceAfterPercent`, { maximum: MAX_PARAGRAPH_SPACING_MULTIPLIER }) };
  } else if (new Set(["spaceAfterPoints", "spaceAfterMultiplier", "noSpaceAfter"]).has(original?.spaceAfter?.case)) {
    spaceAfter = { case: "noSpaceAfter", value: true };
  }

  return {
    ...(lineSpacing ? { lineSpacing } : {}),
    ...(spaceBefore ? { spaceBefore } : {}),
    ...(spaceAfter ? { spaceAfter } : {}),
  };
}

function wireParagraph(paragraph, textStyle, original, shapeId, assetCatalog, { forceLevel = false, customShowLinks } = {}) {
  const unsupported = Object.keys(paragraph).filter((key) => !PARAGRAPH_KEYS.has(key));
  if (unsupported.length) throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses unsupported paragraph fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const paragraphStyleUnsupported = unsupportedStyleFields(paragraph.style);
  if (paragraphStyleUnsupported.length) throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses unsupported paragraph text style fields: ${paragraphStyleUnsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const level = Number(paragraph.level || 0);
  if (!Number.isInteger(level) || level < 0 || level > 8) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses a paragraph level outside the supported 0-8 range.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.alignment && !new Set(["left", "center", "right", "justify"]).has(paragraph.alignment)) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses unsupported paragraph alignment ${paragraph.alignment}.`, [], { code: "invalid_presentation_text" });
  }
  const originalLevel = original?.level;
  const includeLevel = forceLevel || level !== 0 || originalLevel !== undefined;
  const bullet = wireBullet(paragraph, original, shapeId, assetCatalog);
  const bulletFont = wireBulletFont(paragraph, original, shapeId);
  const bulletColor = wireBulletColor(paragraph, original, shapeId);
  const bulletSize = wireBulletSize(paragraph, original, shapeId);
  const tabs = wireTabStops(paragraph, original, shapeId);
  const layout = wireParagraphLayout(paragraph, original, shapeId);
  const spacing = wireParagraphSpacing(paragraph, original, shapeId);
  const defaultRunStyle = wireDefaultRunStyle(paragraph, original, shapeId);
  const directInheritedStyle = Object.fromEntries(Object.entries(textStyle).filter(([key]) => !Object.hasOwn(paragraph.style || {}, key)));
  return {
    ...(includeLevel ? { level } : {}),
    ...(paragraph.alignment ? { alignment: paragraph.alignment } : {}),
    runs: (paragraph.runs || []).map((run, index) => wireRun(run, directInheritedStyle, shapeId, original?.runs?.[index], customShowLinks)),
    ...(bullet ? { bullet } : {}),
    ...(bulletFont ? { bulletFont } : {}),
    ...(bulletColor ? { bulletColor } : {}),
    ...(bulletSize ? { bulletSize } : {}),
    ...tabs,
    ...layout,
    ...spacing,
    ...(defaultRunStyle ? { defaultRunStyle } : {}),
  };
}

function modelRunCase(run) {
  if (run.break) return "lineBreak";
  if (run.field) return "field";
  return "text";
}

function wireTextBodyProperties(value, original, shapeId) {
  let properties;
  try {
    properties = normalizePresentationTextBodyProperties(value);
  } catch (error) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses invalid text body properties: ${error.message}`, [], { code: "invalid_presentation_text" });
  }
  const originalProperties = original?.bodyProperties;
  const insetChoice = (key, wireName, noWireName) => {
    if (properties.insets?.[key] != null) {
      const emu = Math.round(properties.insets[key] * EMU_PER_PIXEL);
      if (emu < 0 || emu > MAX_TEXT_BODY_INSET_EMU) throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses an out-of-range ${key} text inset.`, [], { code: "invalid_presentation_text" });
      return { case: wireName, value: BigInt(emu) };
    }
    const originalCase = originalProperties?.[`${key}Inset`]?.case;
    return new Set([wireName, noWireName]).has(originalCase) ? { case: noWireName, value: true } : undefined;
  };
  const leftInset = insetChoice("left", "leftInsetEmu", "noLeftInset");
  const topInset = insetChoice("top", "topInsetEmu", "noTopInset");
  const rightInset = insetChoice("right", "rightInsetEmu", "noRightInset");
  const bottomInset = insetChoice("bottom", "bottomInsetEmu", "noBottomInset");
  const anchor = properties.anchor != null
    ? { case: "verticalAnchor", value: properties.anchor }
    : new Set(["verticalAnchor", "noVerticalAnchor"]).has(originalProperties?.anchor?.case)
      ? { case: "noVerticalAnchor", value: true }
      : undefined;
  const wrapping = properties.wrap != null
    ? { case: "wrap", value: properties.wrap }
    : new Set(["wrap", "noWrap"]).has(originalProperties?.wrapping?.case)
      ? { case: "noWrap", value: true }
      : undefined;
  const autoFit = properties.autoFit != null
    ? { case: "autoFitMode", value: properties.autoFit }
    : new Set(["autoFitMode", "noAutoFitMode"]).has(originalProperties?.autoFit?.case)
      ? { case: "noAutoFitMode", value: true }
      : undefined;
  const normalAutoFitChoice = (key, wireName, noWireName) => {
    if (properties.normalAutoFit?.[key] != null) return { case: wireName, value: Math.round(properties.normalAutoFit[key] * 1000) };
    const originalCase = originalProperties?.normalAutoFit?.[key]?.case;
    return new Set([wireName, noWireName]).has(originalCase) ? { case: noWireName, value: true } : undefined;
  };
  const normalAutoFitFontScale = properties.autoFit === "shrinkText" ? normalAutoFitChoice("fontScale", "fontScale1000", "noFontScale") : undefined;
  const normalAutoFitLineSpacingReduction = properties.autoFit === "shrinkText" ? normalAutoFitChoice("lineSpacingReduction", "lineSpacingReduction1000", "noLineSpacingReduction") : undefined;
  const normalAutoFit = normalAutoFitFontScale || normalAutoFitLineSpacingReduction
    ? {
        ...(normalAutoFitFontScale ? { fontScale: normalAutoFitFontScale } : {}),
        ...(normalAutoFitLineSpacingReduction ? { lineSpacingReduction: normalAutoFitLineSpacingReduction } : {}),
      }
    : undefined;
  const rotation = properties.rotation != null
    ? { case: "rotationAngle60000", value: Math.round(properties.rotation * ROTATION_UNITS_PER_DEGREE) }
    : new Set(["rotationAngle60000", "noRotation"]).has(originalProperties?.rotation?.case)
      ? { case: "noRotation", value: true }
      : undefined;
  const verticalText = properties.verticalText != null
    ? { case: "verticalTextMode", value: properties.verticalText }
    : new Set(["verticalTextMode", "noVerticalTextMode"]).has(originalProperties?.verticalText?.case)
      ? { case: "noVerticalTextMode", value: true }
      : undefined;
  const verticalOverflow = properties.verticalOverflow != null
    ? { case: "verticalOverflowMode", value: properties.verticalOverflow }
    : new Set(["verticalOverflowMode", "noVerticalOverflowMode"]).has(originalProperties?.verticalOverflow?.case)
      ? { case: "noVerticalOverflowMode", value: true }
      : undefined;
  const horizontalOverflow = properties.horizontalOverflow != null
    ? { case: "horizontalOverflowMode", value: properties.horizontalOverflow }
    : new Set(["horizontalOverflowMode", "noHorizontalOverflowMode"]).has(originalProperties?.horizontalOverflow?.case)
      ? { case: "noHorizontalOverflowMode", value: true }
      : undefined;
  const columnCount = properties.columns?.count != null
    ? { case: "columns", value: properties.columns.count }
    : new Set(["columns", "noColumns"]).has(originalProperties?.columnCount?.case)
      ? { case: "noColumns", value: true }
      : undefined;
  const columnSpacing = properties.columns?.spacing != null
    ? { case: "columnSpacingEmu", value: BigInt(Math.round(properties.columns.spacing * EMU_PER_PIXEL)) }
    : new Set(["columnSpacingEmu", "noColumnSpacing"]).has(originalProperties?.columnSpacing?.case)
      ? { case: "noColumnSpacing", value: true }
      : undefined;
  const columnDirection = properties.columns?.rightToLeft != null
    ? { case: "rightToLeftColumns", value: properties.columns.rightToLeft }
    : new Set(["rightToLeftColumns", "noColumnDirection"]).has(originalProperties?.columnDirection?.case)
      ? { case: "noColumnDirection", value: true }
      : undefined;
  const uprightText = properties.upright != null
    ? { case: "upright", value: properties.upright }
    : new Set(["upright", "noUpright"]).has(originalProperties?.uprightText?.case)
      ? { case: "noUpright", value: true }
      : undefined;
  if (![leftInset, topInset, rightInset, bottomInset, anchor, wrapping, autoFit, normalAutoFit, rotation, verticalText, verticalOverflow, horizontalOverflow, columnCount, columnSpacing, columnDirection, uprightText].some(Boolean)) return undefined;
  return {
    ...(leftInset ? { leftInset } : {}),
    ...(topInset ? { topInset } : {}),
    ...(rightInset ? { rightInset } : {}),
    ...(bottomInset ? { bottomInset } : {}),
    ...(anchor ? { anchor } : {}),
    ...(wrapping ? { wrapping } : {}),
    ...(autoFit ? { autoFit } : {}),
    ...(normalAutoFit ? { normalAutoFit } : {}),
    ...(rotation ? { rotation } : {}),
    ...(verticalText ? { verticalText } : {}),
    ...(verticalOverflow ? { verticalOverflow } : {}),
    ...(horizontalOverflow ? { horizontalOverflow } : {}),
    ...(columnCount ? { columnCount } : {}),
    ...(columnSpacing ? { columnSpacing } : {}),
    ...(columnDirection ? { columnDirection } : {}),
    ...(uprightText ? { uprightText } : {}),
  };
}

function presentationTextBody(shape, original, assetCatalog, customShowLinks) {
  const textStyle = shape.text?.style || {};
  const textStyleUnsupported = Object.keys(textStyle).filter((key) => !RUN_STYLE_KEYS.has(key) && !TEXT_FRAME_PARAGRAPH_KEYS.has(key));
  if (textStyleUnsupported.length) throw new OfficeKitCodecError(`Presentation shape ${shape.id} uses unsupported text-frame style fields: ${textStyleUnsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const inheritedRunStyle = Object.fromEntries(Object.entries(textStyle).filter(([key]) => RUN_STYLE_KEYS.has(key)));
  const inheritedParagraph = Object.fromEntries(Object.entries(textStyle).filter(([key]) => TEXT_FRAME_PARAGRAPH_KEYS.has(key)));
  const paragraphs = shape.text?.paragraphs || [];
  if (original?.textBody && (original.textBody.paragraphs.length !== paragraphs.length || original.textBody.paragraphs.some((paragraph, index) => paragraph.runs.length !== (paragraphs[index]?.runs || []).length || paragraph.runs.some((run, runIndex) => run.content?.case !== modelRunCase(paragraphs[index].runs[runIndex]))))) {
    throw new OfficeKitCodecError(`Presentation shape ${shape.id} changed its source-bound paragraph/inline topology.`, [], { code: "presentation_text_topology_changed" });
  }
  const originalListStyles = new Map((original?.textBody?.listStyles || []).map((style) => [Number(style.level), style]));
  const inheritedParagraphStyles = Object.entries(shape.text?.inheritedParagraphStyles || {}).sort(([left], [right]) => Number(left) - Number(right));
  const listStyles = inheritedParagraphStyles.map(([level, style]) => wireParagraph(
    { ...style, level: Number(level), runs: [] },
    {},
    originalListStyles.get(Number(level)),
    shape.id,
    assetCatalog,
    { forceLevel: true, customShowLinks },
  ));
  const noListStyles = listStyles.length === 0 && (originalListStyles.size > 0 || original?.textBody?.noListStyles === true);
  const bodyProperties = wireTextBodyProperties(shape.text?.bodyProperties, original?.textBody, shape.id);
  return {
    paragraphs: paragraphs.map((paragraph, index) => wireParagraph({ ...inheritedParagraph, ...paragraph }, inheritedRunStyle, original?.textBody?.paragraphs?.[index], shape.id, assetCatalog, { customShowLinks })),
    ...(listStyles.length ? { listStyles } : {}),
    ...(noListStyles ? { noListStyles: true } : {}),
    ...(bodyProperties ? { bodyProperties } : {}),
  };
}

// Speaker notes reuse the public paragraph/run wire rather than creating a
// second rich-text format. Notes-local text is intentionally narrower than a
// slide shape: the native codec rejects relationships, fields, picture
// bullets, list styles, and body properties. An imported notes part without a
// projected textBody remains text-only, so an unchanged round trip can never
// silently turn an opaque source into a lossy rich edit request.
function presentationSpeakerNotes(slide, original, assetCatalog, customShowLinks) {
  const notes = slide.speakerNotes;
  if (original) {
    const result = { text: notes?.text || "", source: original.source };
    if (original.textBody) {
      result.textBody = presentationTextBody(
        { id: `${slide.id}/notes`, text: notes?.textFrame },
        original,
        assetCatalog,
        customShowLinks,
      );
    }
    return result;
  }
  if (!notes?.text) return undefined;
  return {
    text: notes.text,
    textBody: presentationTextBody(
      { id: `${slide.id}/notes`, text: notes.textFrame },
      undefined,
      assetCatalog,
      customShowLinks,
    ),
  };
}

const MASTER_STYLE_KINDS = [
  ["title", "titleLevels", "deletedTitleLevels"],
  ["body", "bodyLevels", "deletedBodyLevels"],
  ["other", "otherLevels", "deletedOtherLevels"],
];

function wireMasterTextStyles(master, original, assetCatalog) {
  const result = {};
  for (const [kind, levelsField, deletedField] of MASTER_STYLE_KINDS) {
    const originalLevels = new Map((original?.textStyles?.[levelsField] || []).map((style) => [Number(style.level), style]));
    const current = master.textParagraphStyles?.[kind] || {};
    const levels = Object.entries(current)
      .sort(([left], [right]) => Number(left) - Number(right))
      .map(([level, style]) => wireParagraph(
        { ...style, level: Number(level), runs: [] },
        {},
        originalLevels.get(Number(level)),
        `master ${master.id} ${kind} level ${Number(level) + 1}`,
        assetCatalog,
        { forceLevel: true },
      ));
    const currentLevels = new Set(Object.keys(current).map(Number));
    const deleted = [...originalLevels.keys()].filter((level) => !currentLevels.has(level)).sort((left, right) => left - right);
    result[levelsField] = levels;
    if (deleted.length) result[deletedField] = deleted;
  }
  return result;
}

function hasPresentationBackground(background) {
  return Boolean(background && (background.fill || background.image));
}

function wireBackground(background, ownerId, assetCatalog) {
  if (!background) return undefined;
  if (background.image) {
    const image = background.image;
    const assetId = image.dataUrl ? assetCatalog.addDataUrl(image.dataUrl) : String(image.assetId || "").trim();
    if (!assetId) throw new OfficeKitCodecError(`Presentation ${ownerId} image background requires an embedded asset.`, [], { code: "invalid_presentation_background" });
    if (image.dataUrl && image.assetId && String(image.assetId).trim() !== assetId) {
      throw new OfficeKitCodecError(`Presentation ${ownerId} image background assetId does not match its dataUrl.`, [], { code: "invalid_presentation_background" });
    }
    return { imageAssetId: assetId, ...(image.alphaModulationFixed ? { imageAlphaModulationFixed: true } : {}) };
  }
  const fill = String(background.fill || "").trim();
  const color = PRESENTATION_SCHEME_COLORS.has(fill)
    ? { case: "colorScheme", value: fill }
    : { case: "colorRgb", value: presentationRgb(fill, `${ownerId}.background.fill`) };
  if (!color.value) throw new OfficeKitCodecError(`Presentation ${ownerId} uses an unsupported transparent background.`, [], { code: "unsupported_presentation_features" });
  if (background.mode === "reference") {
    const index = Number(background.index);
    if (!Number.isInteger(index) || index < 0 || index > 4_294_967_295) {
      throw new OfficeKitCodecError(`Presentation ${ownerId} background reference index must be an unsigned 32-bit integer.`, [], { code: "invalid_presentation_background" });
    }
    return { color, kind: { case: "styleReferenceIndex", value: index } };
  }
  if (background.mode !== "solid") throw new OfficeKitCodecError(`Presentation ${ownerId} background mode must be solid or reference.`, [], { code: "invalid_presentation_background" });
  return { color, kind: { case: "solid", value: true } };
}

function modelBackground(background, assetCatalog) {
  if (background?.imageAssetId) {
    const assetId = String(background.imageAssetId);
    return { image: { assetId, dataUrl: assetCatalog.dataUrl(assetId), fit: "stretch", ...(background.imageAlphaModulationFixed ? { alphaModulationFixed: true } : {}) } };
  }
  if (!background?.color?.case || !background?.kind?.case) return undefined;
  const fill = background.color.case === "colorScheme" ? background.color.value : `#${String(background.color.value).toLowerCase()}`;
  return background.kind.case === "styleReferenceIndex"
    ? { fill, mode: "reference", index: Number(background.kind.value) }
    : { fill, mode: "solid" };
}

function wirePresentationTransition(transition) {
  const value = transition?.toJSON?.();
  if (!value) return undefined;
  return {
    effect: value.effect,
    ...(value.direction ? { direction: value.direction } : {}),
    ...(value.orientation ? { orientation: value.orientation } : {}),
    ...(value.throughBlack === undefined ? {} : { throughBlack: value.throughBlack }),
    ...(value.spokes === undefined ? {} : { spokes: value.spokes }),
    speed: value.speed,
    ...(value.durationMs === undefined ? {} : { durationMs: value.durationMs }),
    advanceOnClick: value.advanceOnClick,
    ...(value.advanceAfterMs === undefined ? {} : { advanceAfterMs: value.advanceAfterMs }),
  };
}

function modelPresentationTransition(source, slideIndex) {
  if (!source) return undefined;
  if (typeof source.advanceOnClick !== "boolean") {
    throw new OfficeKitCodecError(`OfficeKit returned slide ${slideIndex + 1} transition without an explicit advanceOnClick value.`, [], { code: "invalid_presentation_artifact" });
  }
  try {
    return normalizePresentationTransition({
      effect: source.effect,
      ...(source.direction ? { direction: source.direction } : {}),
      ...(source.orientation ? { orientation: source.orientation } : {}),
      ...(source.throughBlack === undefined ? {} : { throughBlack: source.throughBlack }),
      ...(source.spokes === undefined ? {} : { spokes: Number(source.spokes) }),
      speed: source.speed,
      ...(source.durationMs === undefined ? {} : { durationMs: Number(source.durationMs) }),
      advanceOnClick: source.advanceOnClick,
      ...(source.advanceAfterMs === undefined ? {} : { advanceAfterMs: Number(source.advanceAfterMs) }),
    });
  } catch (error) {
    throw new OfficeKitCodecError(`OfficeKit returned invalid slide ${slideIndex + 1} transition semantics: ${error.message}`, [], { code: "invalid_presentation_artifact" });
  }
}

function wirePresentationAnimation(animation, ownerLabel) {
  const value = animation && typeof animation.toJSON === "function" ? animation.toJSON() : animation;
  if (!value) return undefined;
  try {
    return {
      id: value.id,
      targetId: value.targetId,
      targetKind: value.targetKind,
      effect: value.effect,
      phase: value.phase,
      start: value.start,
      ...(value.direction ? { direction: value.direction } : {}),
      durationMs: value.durationMs,
      ...(value.delayMs === undefined ? {} : { delayMs: value.delayMs }),
      ...(value.textBuild ? { textBuild: value.textBuild } : {}),
      ...(value.chartBuild ? { chartBuild: value.chartBuild } : {}),
      ...(value.staggerMs === undefined ? {} : { staggerMs: value.staggerMs }),
      ...(value.animateChartBackground === undefined ? {} : { animateChartBackground: value.animateChartBackground }),
    };
  } catch (error) {
    throw new OfficeKitCodecError(`Invalid ${ownerLabel} animation: ${error.message}`, [], { code: "invalid_presentation_animation" });
  }
}

function modelPresentationAnimation(source, ownerLabel) {
  try {
    return normalizePresentationAnimation({
      id: source.id,
      targetId: source.targetId,
      targetKind: source.targetKind,
      effect: source.effect,
      phase: source.phase,
      start: source.start,
      ...(source.direction ? { direction: source.direction } : {}),
      durationMs: Number(source.durationMs),
      ...(source.delayMs === undefined ? {} : { delayMs: Number(source.delayMs) }),
      ...(source.textBuild ? { textBuild: source.textBuild } : {}),
      ...(source.chartBuild ? { chartBuild: source.chartBuild } : {}),
      ...(source.staggerMs === undefined ? {} : { staggerMs: Number(source.staggerMs) }),
      ...(source.animateChartBackground === undefined ? {} : { animateChartBackground: Boolean(source.animateChartBackground) }),
    });
  } catch (error) {
    throw new OfficeKitCodecError(`Invalid ${ownerLabel} animation: ${error.message}`, [], { code: "invalid_presentation_artifact" });
  }
}

function wirePresentationMorph(morph, ownerLabel) {
  if (!morph) return undefined;
  try {
    const value = normalizePresentationMorph(morph);
    return { durationMs: value.durationMs, pairs: value.pairs.map((pair) => ({ key: pair.key, fromId: pair.fromId, toId: pair.toId })), fromSlideId: value.fromSlideId };
  } catch (error) {
    throw new OfficeKitCodecError(`Invalid ${ownerLabel} Morph: ${error.message}`, [], { code: "invalid_presentation_morph" });
  }
}

function modelPresentationMorph(source, ownerLabel) {
  if (!source) return undefined;
  try {
    return normalizePresentationMorph({
      durationMs: Number(source.durationMs),
      pairs: (source.pairs || []).map((pair) => ({ key: pair.key, fromId: pair.fromId, toId: pair.toId })),
      fromSlideId: source.fromSlideId,
    });
  } catch (error) {
    throw new OfficeKitCodecError(`Invalid ${ownerLabel} Morph: ${error.message}`, [], { code: "invalid_presentation_artifact" });
  }
}

function wirePresentationTransform(transform, ownerLabel) {
  if (transform == null) return {};
  if (typeof transform !== "object" || Array.isArray(transform)) {
    throw new OfficeKitCodecError(`Presentation ${ownerLabel} transform must be an object.`, [], { code: "invalid_presentation_transform" });
  }
  const output = {};
  if (Object.hasOwn(transform, "rotationDegrees") && transform.rotationDegrees != null) {
    const degrees = Number(transform.rotationDegrees);
    if (!Number.isFinite(degrees) || degrees < -360 || degrees > 360) {
      throw new OfficeKitCodecError(`Presentation ${ownerLabel} rotation must be between -360 and 360 degrees.`, [], { code: "invalid_presentation_transform" });
    }
    output.rotationAngle60000 = Math.round(degrees * ROTATION_UNITS_PER_DEGREE);
  }
  for (const key of ["flipHorizontal", "flipVertical"]) {
    if (!Object.hasOwn(transform, key) || transform[key] == null) continue;
    if (typeof transform[key] !== "boolean") {
      throw new OfficeKitCodecError(`Presentation ${ownerLabel} ${key} must be a boolean.`, [], { code: "invalid_presentation_transform" });
    }
    output[key] = transform[key];
  }
  if (Object.keys(output).length === 0) {
    throw new OfficeKitCodecError(`Presentation ${ownerLabel} transform must define rotationDegrees, flipHorizontal, or flipVertical.`, [], { code: "invalid_presentation_transform" });
  }
  return output;
}

function masterReadOnlySnapshot(master) {
  return JSON.stringify(master.toJSON());
}

function layoutReadOnlySnapshot(layout) {
  return JSON.stringify(layout.toJSON());
}

function sourceFreeLayoutType(type, layoutId) {
  const requested = String(type || "blank").trim();
  const normalized = SOURCE_FREE_LAYOUT_TYPES.get(requested);
  if (!normalized) {
    throw new OfficeKitCodecError(
      `Presentation layout ${layoutId} uses unsupported source-free type ${requested || "(empty)"}. Use blank, title, titleOnly, or obj/titleAndContent.`,
      [],
      { code: "unsupported_presentation_features" },
    );
  }
  return normalized;
}

function sourceFreePlaceholder(placeholder, ownerId, assetCatalog, customShowLinks) {
  const type = String(placeholder.type || "");
  if (!SOURCE_FREE_TEXT_PLACEHOLDER_TYPES.has(type)) {
    throw new OfficeKitCodecError(
      `Presentation placeholder ${placeholder.id || ownerId} uses ${type || "(empty)"}; source-free layouts currently author only title, body, ctrTitle, and subTitle text placeholders.`,
      [],
      { code: "unsupported_presentation_features" },
    );
  }
  if (!placeholder.position) {
    throw new OfficeKitCodecError(
      `Presentation placeholder ${placeholder.id || ownerId} requires a direct position for source-free PPTX export.`,
      [],
      { code: "invalid_presentation_placeholder" },
    );
  }
  const index = Number(placeholder.idx);
  if (!Number.isInteger(index) || index < 0 || index > 4_294_967_295) {
    throw new OfficeKitCodecError(`Presentation placeholder ${placeholder.id || ownerId} has an invalid idx.`, [], { code: "invalid_presentation_placeholder" });
  }
  const shape = {
    id: placeholder.id || `${ownerId}/placeholder/${index}`,
    text: {
      style: { ...(placeholder.style || {}) },
      paragraphs: normalizePresentationParagraphs(placeholder.text ?? ""),
      inheritedParagraphStyles: { ...(placeholder.paragraphStyles || {}) },
      bodyProperties: placeholder.textBodyProperties,
    },
  };
  const position = placeholder.position;
  return {
    id: shape.id,
    name: String(placeholder.name || `${type} placeholder`),
    type,
    index,
    textBody: presentationTextBody(shape, undefined, assetCatalog, customShowLinks),
    directFrame: {
      leftEmu: emuFromPixels(position.left, `${shape.id}.position.left`),
      topEmu: emuFromPixels(position.top, `${shape.id}.position.top`),
      widthEmu: emuFromPixels(position.width, `${shape.id}.position.width`),
      heightEmu: emuFromPixels(position.height, `${shape.id}.position.height`),
      ...wirePresentationTransform(placeholder.transform, `placeholder ${shape.id}`),
    },
  };
}

function sourceFreeSlidePlaceholder(shape) {
  if (!shape.placeholder) return undefined;
  const type = String(shape.placeholder.type || "");
  if (!SOURCE_FREE_TEXT_PLACEHOLDER_TYPES.has(type)) {
    throw new OfficeKitCodecError(
      `Presentation slide placeholder ${shape.id} uses ${type || "(empty)"}; source-free layouts currently author only title, body, ctrTitle, and subTitle text placeholders.`,
      [],
      { code: "unsupported_presentation_features" },
    );
  }
  const index = Number(shape.placeholder.idx ?? shape.placeholder.index);
  if (!Number.isInteger(index) || index < 0 || index > 4_294_967_295) {
    throw new OfficeKitCodecError(`Presentation slide placeholder ${shape.id} has an invalid idx.`, [], { code: "invalid_presentation_placeholder" });
  }
  const position = shape.position || {};
  return {
    placeholder: { type, index, inheritsGeometry: false },
    directFrame: {
      leftEmu: emuFromPixels(position.left, `${shape.id}.position.left`),
      topEmu: emuFromPixels(position.top, `${shape.id}.position.top`),
      widthEmu: emuFromPixels(position.width, `${shape.id}.position.width`),
      heightEmu: emuFromPixels(position.height, `${shape.id}.position.height`),
      ...wirePresentationTransform(shape.transform, `placeholder ${shape.id}`),
    },
  };
}

function presentationMasters(presentation, state, assetCatalog, customShowLinks) {
  if (state) {
    if (presentation.masters.items.length !== state.masters.length || state.masters.some((entry, index) => presentation.masters.items[index] !== entry.model)) {
      throw new OfficeKitCodecError(`Source-preserving PPTX export requires the original ${state.masters.length}-master topology.`, [], { code: "presentation_master_topology_changed" });
    }
    return state.masters.map((entry) => {
      if (masterReadOnlySnapshot(entry.model) !== entry.snapshot) {
        throw new OfficeKitCodecError(`Presentation master ${entry.model.id} is source-bound and read-only in OfficeKit 0.2.`, [], { code: "unsupported_presentation_edit" });
      }
      return entry.wire;
    });
  }
  const master = presentation.master;
  return master ? [{
    id: master.id,
    name: master.name,
    textStyles: wireMasterTextStyles(master, undefined, assetCatalog),
    background: wireBackground(master.background, `master ${master.id}`, assetCatalog),
    placeholders: master.placeholders.map((placeholder) => sourceFreePlaceholder(placeholder, master.id, assetCatalog, customShowLinks)),
  }] : [];
}

function presentationLayouts(presentation, state, assetCatalog, customShowLinks) {
  if (!state) {
    return presentation.layouts.items.map((layout) => ({
      id: layout.id,
      name: layout.name,
      masterId: layout.masterId,
      type: sourceFreeLayoutType(layout.type, layout.id),
      ...(layout.background ? { background: wireBackground(layout.background, `layout ${layout.id}`, assetCatalog) } : {}),
      placeholders: layout.placeholders.map((placeholder) => sourceFreePlaceholder(placeholder, layout.id, assetCatalog, customShowLinks)),
    }));
  }
  if (presentation.layouts.items.length !== state.layouts.length || state.layouts.some((entry, index) => presentation.layouts.items[index] !== entry.model)) {
    throw new OfficeKitCodecError(`Source-preserving PPTX export requires the original ${state.layouts.length}-layout topology.`, [], { code: "presentation_layout_topology_changed" });
  }
  return state.layouts.map((entry) => {
    if (layoutReadOnlySnapshot(entry.model) !== entry.snapshot) {
      throw new OfficeKitCodecError(`Presentation layout ${entry.model.id} is source-bound and read-only in OfficeKit 0.2.`, [], { code: "unsupported_presentation_edit" });
    }
    return entry.wire;
  });
}

function presentationShadow(shadow, shapeId) {
  if (shadow == null || shadow === false || shadow === "shadow-none") return undefined;
  const presets = {
    "shadow-sm": { color: "#000000", blurRadius: 4, distance: 2, direction: 45, opacity: 0.15 },
    shadow: { color: "#000000", blurRadius: 6, distance: 3, direction: 45, opacity: 0.18 },
    "shadow-md": { color: "#000000", blurRadius: 10, distance: 4, direction: 45, opacity: 0.2 },
    "shadow-lg": { color: "#000000", blurRadius: 15, distance: 6, direction: 45, opacity: 0.22 },
    "shadow-xl": { color: "#000000", blurRadius: 22, distance: 9, direction: 45, opacity: 0.24 },
    "shadow-2xl": { color: "#000000", blurRadius: 32, distance: 14, direction: 45, opacity: 0.25 },
  };
  let source = typeof shadow === "string" ? presets[shadow] : shadow;
  if (!source && typeof shadow === "string") {
    const match = /^(-?\d+(?:\.\d+)?)px\s+(-?\d+(?:\.\d+)?)px\s+(\d+(?:\.\d+)?)px\s+(#[0-9a-f]{6})(?:\/(\d+(?:\.\d+)?))?$/i.exec(shadow.trim());
    if (match) {
      const offsetX = Number(match[1]);
      const offsetY = Number(match[2]);
      source = {
        color: match[4],
        blurRadius: Number(match[3]),
        distance: Math.hypot(offsetX, offsetY),
        direction: (Math.atan2(offsetY, offsetX) * 180 / Math.PI + 360) % 360,
        opacity: match[5] == null ? 1 : Number(match[5]) / 100,
      };
    }
  }
  if (!source || typeof source !== "object") {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} uses an unsupported shadow.`, [], { code: "unsupported_presentation_features" });
  }
  const blurRadius = Number(source.blurRadius ?? source.blur ?? 0);
  const distance = Number(source.distance ?? 0);
  const direction = Number(source.direction ?? source.angle ?? 0);
  const opacity = Number(source.opacity ?? 0.2);
  if (![blurRadius, distance, direction, opacity].every(Number.isFinite) || blurRadius < 0 || distance < 0 || opacity < 0 || opacity > 1) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} has an invalid shadow.`, [], { code: "invalid_presentation_shadow" });
  }
  const normalizedDirection = ((direction % 360) + 360) % 360;
  return {
    colorRgb: presentationRgb(source.color || source.fill || "#000000", `${shapeId}.shadow.color`),
    blurRadiusEmu: emuFromPixels(blurRadius, `${shapeId}.shadow.blurRadius`),
    distanceEmu: emuFromPixels(distance, `${shapeId}.shadow.distance`),
    directionAngle60000: Math.round(normalizedDirection * ROTATION_UNITS_PER_DEGREE),
    opacityThousandthPercent: Math.round(opacity * 100_000),
  };
}

function modelPresentationShadow(shadow) {
  if (!shadow) return undefined;
  return {
    color: shadow.colorRgb ? `#${shadow.colorRgb}` : "#000000",
    blurRadius: Number(shadow.blurRadiusEmu) / EMU_PER_PIXEL,
    distance: Number(shadow.distanceEmu) / EMU_PER_PIXEL,
    direction: Number(shadow.directionAngle60000) / ROTATION_UNITS_PER_DEGREE,
    opacity: Number(shadow.opacityThousandthPercent) / 100_000,
  };
}

function modelPresentationAccessibility(value, owner = "Imported Presentation shape") {
  if (!value) return {};
  const accessibility = normalizePresentationAccessibility({
    ...(value.title === undefined ? {} : { title: value.title }),
    ...(value.description === undefined ? {} : { description: value.description }),
    ...(value.decorative === undefined ? {} : { decorative: value.decorative }),
  }, owner);
  return accessibility ? { accessibility } : {};
}

function modelPresentationImageAccessibility(image) {
  return modelPresentationAccessibility({
    ...(image.accessibilityTitle ? { title: image.accessibilityTitle } : {}),
    ...(image.altText ? { description: image.altText } : {}),
    ...(image.accessibilityDecorative === undefined ? {} : { decorative: image.accessibilityDecorative }),
  }, "Imported Presentation image");
}

function sourceBoundCloneConnectorTargetId(value, sourceIdByCloneId, connector, side) {
  const targetId = String(value || "");
  if (!targetId || !sourceIdByCloneId) return targetId;
  const sourceId = sourceIdByCloneId.get(targetId);
  if (!sourceId) {
    throw new OfficeKitCodecError(`Imported presentation clone connector ${connector.id} has an unresolved ${side} target.`, [], { code: "unsupported_presentation_slide_clone" });
  }
  return sourceId;
}

function presentationConnector(connector, original, sourceIdByCloneId) {
  const type = String(connector.connectorType || "straight");
  if (!new Set(["straight", "elbow", "curved"]).has(type)) {
    throw new OfficeKitCodecError(`Presentation connector ${connector.id} uses unsupported type ${type}.`, [], { code: "unsupported_presentation_features" });
  }
  const sourceLine = connector.line || {};
  const line = normalizePresentationLineStyle({
    ...sourceLine,
    ...(connector.head == null ? {} : { head: connector.head }),
    ...(connector.tail == null ? {} : { tail: connector.tail }),
    ...(connector.cap == null ? {} : { cap: connector.cap }),
    ...(connector.join == null ? {} : { join: connector.join }),
  }, {
    name: `Presentation connector ${connector.id} line`,
    defaultWidth: 2,
  });
  const width = line.width;
  const head = line.head || {};
  const tail = line.tail || {};
  const startSiteIndex = Number(connector.startSiteIndex ?? 0);
  const endSiteIndex = Number(connector.endSiteIndex ?? 0);
  if (![startSiteIndex, endSiteIndex].every((value) => Number.isInteger(value) && value >= 0 && value <= 0xffff_ffff)) {
    throw new OfficeKitCodecError(`Presentation connector ${connector.id} has an invalid connection-site index.`, [], { code: "invalid_presentation_connector" });
  }
  const endpoints = typeof connector.resolvedEndpoints === "function"
    ? connector.resolvedEndpoints({ strict: true })
    : { start: connector.start, end: connector.end };
  const startTargetId = sourceBoundCloneConnectorTargetId(connector.startTargetId, sourceIdByCloneId, connector, "start");
  const endTargetId = sourceBoundCloneConnectorTargetId(connector.endTargetId, sourceIdByCloneId, connector, "end");
  if ((!startTargetId && startSiteIndex !== 0) || (!endTargetId && endSiteIndex !== 0)) {
    throw new OfficeKitCodecError(`Presentation connector ${connector.id} cannot define a connection-site index without its target.`, [], { code: "invalid_presentation_connector" });
  }
  const lineRgb = line.style === "none"
    ? ""
    : presentationRgb(presentationLineColor(line, width > 0 ? "#334155" : "transparent"), `${connector.id}.line.fill`);
  const accessibility = normalizePresentationAccessibility(connector.accessibility, `Presentation connector ${connector.id}`);
  return {
    id: original?.id || connector.id,
    name: connector.name || original?.name || "",
    source: original?.source,
    content: {
      case: "connector",
      value: {
        connectorType: type,
        startXEmu: sourceBoundFrameEmuFromPixels(endpoints.start?.x, `${connector.id}.start.x`, original),
        startYEmu: sourceBoundFrameEmuFromPixels(endpoints.start?.y, `${connector.id}.start.y`, original),
        endXEmu: sourceBoundFrameEmuFromPixels(endpoints.end?.x, `${connector.id}.end.x`, original),
        endYEmu: sourceBoundFrameEmuFromPixels(endpoints.end?.y, `${connector.id}.end.y`, original),
        lineRgb,
        lineWidthEmu: BigInt(Math.round(width * EMU_PER_POINT)),
        startArrow: head.type || "",
        endArrow: tail.type || "",
        startTargetId,
        endTargetId,
        startConnectionSiteIndex: startSiteIndex,
        endConnectionSiteIndex: endSiteIndex,
        lineStyle: lineRgb ? line.style : "none",
        startArrowWidth: head.width || "",
        startArrowLength: head.length || "",
        endArrowWidth: tail.width || "",
        endArrowLength: tail.length || "",
        lineCap: line.cap || "",
        lineJoin: line.join || "",
        ...(accessibility ? { accessibility } : {}),
      },
    },
  };
}

function presentationCustomGeometryPointToWire(point) {
  return {
    ...(typeof point.x === "string" ? { xReference: point.x } : { x: BigInt(point.x) }),
    ...(typeof point.y === "string" ? { yReference: point.y } : { y: BigInt(point.y) }),
  };
}

function presentationCustomGeometryArcToWire(arc) {
  return {
    ...(typeof arc.widthRadius === "string" ? { widthRadiusReference: arc.widthRadius } : { widthRadius: BigInt(arc.widthRadius) }),
    ...(typeof arc.heightRadius === "string" ? { heightRadiusReference: arc.heightRadius } : { heightRadius: BigInt(arc.heightRadius) }),
    ...(typeof arc.startAngle === "string" ? { startAngleReference: arc.startAngle } : { startAngle: arc.startAngle }),
    ...(typeof arc.sweepAngle === "string" ? { sweepAngleReference: arc.sweepAngle } : { sweepAngle: arc.sweepAngle }),
  };
}

function presentationCustomGeometryConnectionSiteToWire(site) {
  return {
    ...(typeof site.angle === "string"
      ? { angleReference: site.angle }
      : { angle60000: Math.round(site.angle * ROTATION_UNITS_PER_DEGREE) }),
    ...(typeof site.x === "string"
      ? { xReference: site.x }
      : { xEmu: BigInt(Math.round(site.x * EMU_PER_PIXEL)) }),
    ...(typeof site.y === "string"
      ? { yReference: site.y }
      : { yEmu: BigInt(Math.round(site.y * EMU_PER_PIXEL)) }),
  };
}

function presentationCustomGeometryHandleBoundToWire(value, literalField, referenceField, literal) {
  if (value === undefined) return {};
  return typeof value === "string"
    ? { [referenceField]: value }
    : { [literalField]: literal(value) };
}

function presentationCustomGeometryAdjustmentHandleToWire(handle) {
  const position = {
    ...(typeof handle.x === "string" ? { xReference: handle.x } : { x: BigInt(Math.round(handle.x * EMU_PER_PIXEL)) }),
    ...(typeof handle.y === "string" ? { yReference: handle.y } : { y: BigInt(Math.round(handle.y * EMU_PER_PIXEL)) }),
  };
  if (handle.kind === "xy") return {
    handle: {
      case: "xy",
      value: {
        xAdjustment: handle.xAdjustment || "",
        ...presentationCustomGeometryHandleBoundToWire(handle.minX, "minX", "minXReference", BigInt),
        ...presentationCustomGeometryHandleBoundToWire(handle.maxX, "maxX", "maxXReference", BigInt),
        yAdjustment: handle.yAdjustment || "",
        ...presentationCustomGeometryHandleBoundToWire(handle.minY, "minY", "minYReference", BigInt),
        ...presentationCustomGeometryHandleBoundToWire(handle.maxY, "maxY", "maxYReference", BigInt),
        position,
      },
    },
  };
  return {
    handle: {
      case: "polar",
      value: {
        radialAdjustment: handle.radialAdjustment || "",
        ...presentationCustomGeometryHandleBoundToWire(handle.minRadius, "minRadius", "minRadiusReference", BigInt),
        ...presentationCustomGeometryHandleBoundToWire(handle.maxRadius, "maxRadius", "maxRadiusReference", BigInt),
        angleAdjustment: handle.angleAdjustment || "",
        ...presentationCustomGeometryHandleBoundToWire(handle.minAngle, "minAngle60000", "minAngleReference", (value) => Math.round(value * ROTATION_UNITS_PER_DEGREE)),
        ...presentationCustomGeometryHandleBoundToWire(handle.maxAngle, "maxAngle60000", "maxAngleReference", (value) => Math.round(value * ROTATION_UNITS_PER_DEGREE)),
        position,
      },
    },
  };
}

function presentationCustomGeometryAdjustmentHandleIdentity(handle) {
  if (handle?.handle?.case === "xy") {
    return ["xy", handle.handle.value.xAdjustment || "", handle.handle.value.yAdjustment || ""].join("\0");
  }
  if (handle?.handle?.case === "polar") {
    return ["polar", handle.handle.value.radialAdjustment || "", handle.handle.value.angleAdjustment || ""].join("\0");
  }
  return "invalid";
}

function presentationCustomGeometryTextRectangleToWire(rectangle) {
  return Object.fromEntries(CUSTOM_TEXT_RECTANGLE_FIELDS.map(([field, literalField, referenceField]) => (
    typeof rectangle[field] === "string"
      ? [referenceField, rectangle[field]]
      : [literalField, BigInt(Math.round(rectangle[field] * EMU_PER_PIXEL))]
  )));
}

function presentationShape(shape, original, assetCatalog, customShowLinks) {
  const originalShape = original?.content?.case === "shape" ? original.content.value : original;
  if (!new Set(["rect", "ellipse", "roundRect", "textbox", "line", "custom"]).has(shape.geometry)) {
    throw new OfficeKitCodecError(`Presentation shape ${shape.id} uses unsupported geometry ${shape.geometry}.`, [], { code: "unsupported_presentation_features" });
  }
  const formulaGraph = normalizePresentationCustomGeometryFormulaGraph({ adjustments: shape.customAdjustments, guides: shape.customGuides });
  const position = shape.position || {};
  const normalizedCustomPaths = shape.customPaths?.length ? normalizePresentationCustomPaths(shape.customPaths, {
    ...formulaGraph,
    widthEmu: Math.round(Number(position.width) * EMU_PER_PIXEL),
    heightEmu: Math.round(Number(position.height) * EMU_PER_PIXEL),
  }) : [];
  const customConnectionSites = normalizePresentationCustomConnectionSites(shape.customConnectionSites, {
    ...formulaGraph,
    widthEmu: Math.round(Number(position.width) * EMU_PER_PIXEL),
    heightEmu: Math.round(Number(position.height) * EMU_PER_PIXEL),
  });
  const customAdjustmentHandles = normalizePresentationCustomAdjustmentHandles(shape.customAdjustmentHandles, {
    ...formulaGraph,
    widthEmu: Math.round(Number(position.width) * EMU_PER_PIXEL),
    heightEmu: Math.round(Number(position.height) * EMU_PER_PIXEL),
  });
  const wireCustomAdjustmentHandles = customAdjustmentHandles.map(presentationCustomGeometryAdjustmentHandleToWire);
  const textRectangle = normalizePresentationCustomTextRectangle(shape.textRectangle, {
    ...formulaGraph,
    widthEmu: Math.round(Number(position.width) * EMU_PER_PIXEL),
    heightEmu: Math.round(Number(position.height) * EMU_PER_PIXEL),
  });
  if (shape.geometry !== "custom" && (normalizedCustomPaths.length || customConnectionSites.length || customAdjustmentHandles.length || formulaGraph.adjustments.length || formulaGraph.guides.length || textRectangle)) {
    throw new OfficeKitCodecError(`Presentation shape ${shape.id} has custom geometry data without custom geometry.`, [], { code: "invalid_presentation_geometry" });
  }
  if (originalShape?.geometry === "custom" && (originalShape.customConnectionSites?.length || 0) !== customConnectionSites.length) {
    throw new OfficeKitCodecError(`Source-preserving PPTX export requires custom shape ${shape.id}'s original connection-site list length; each existing index is the native identity.`, [], { code: "unsupported_presentation_edit" });
  }
  if (originalShape?.geometry === "custom") {
    const originalHandles = originalShape.customAdjustmentHandles || [];
    const changedHandleTopology = originalHandles.length !== wireCustomAdjustmentHandles.length || originalHandles.some(
      (handle, index) => presentationCustomGeometryAdjustmentHandleIdentity(handle) !== presentationCustomGeometryAdjustmentHandleIdentity(wireCustomAdjustmentHandles[index]),
    );
    if (changedHandleTopology) {
      throw new OfficeKitCodecError(`Source-preserving PPTX export requires custom shape ${shape.id}'s original adjustment-handle order, kind, and controlled adjustment identity.`, [], { code: "unsupported_presentation_edit" });
    }
  }
  const customPaths = normalizedCustomPaths.map((path) => ({
    width: BigInt(path.width ?? 0),
    height: BigInt(path.height ?? 0),
    fillMode: path.fillMode === "normal"
      ? PresentationCustomGeometryPath_FillMode.NORMAL
      : path.fillMode === "none"
        ? PresentationCustomGeometryPath_FillMode.NONE
        : PresentationCustomGeometryPath_FillMode.UNSPECIFIED,
    ...(Object.hasOwn(path, "stroke") ? { stroke: path.stroke } : {}),
    ...(Object.hasOwn(path, "extrusionAllowed") ? { extrusionAllowed: path.extrusionAllowed } : {}),
    commands: path.commands.map((command) => {
      if (command.moveTo) return { command: { case: "moveTo", value: presentationCustomGeometryPointToWire(command.moveTo) } };
      if (command.lineTo) return { command: { case: "lineTo", value: presentationCustomGeometryPointToWire(command.lineTo) } };
      if (command.quadraticBezTo) return {
        command: {
          case: "quadraticBezierTo",
          value: {
            control: presentationCustomGeometryPointToWire({ x: command.quadraticBezTo.x1, y: command.quadraticBezTo.y1 }),
            end: presentationCustomGeometryPointToWire(command.quadraticBezTo),
          },
        },
      };
      if (command.arcTo) return {
        command: {
          case: "arcTo",
          value: presentationCustomGeometryArcToWire(command.arcTo),
        },
      };
      if (command.cubicBezTo) return {
        command: {
          case: "cubicBezierTo",
          value: {
            control1: presentationCustomGeometryPointToWire({ x: command.cubicBezTo.x1, y: command.cubicBezTo.y1 }),
            control2: presentationCustomGeometryPointToWire({ x: command.cubicBezTo.x2, y: command.cubicBezTo.y2 }),
            end: presentationCustomGeometryPointToWire(command.cubicBezTo),
          },
        },
      };
      return { command: { case: "close", value: true } };
    }),
  }));
  // The model deliberately withholds an unrecognized custom-path grammar.
  // An unchanged source-bound, non-editable shape can still be carried by the
  // C# codec, which rechecks its source binding and rejects every mutation.
  // Source-free or editable shapes must continue to provide the full grammar.
  const opaqueSourceBoundCustomGeometry = original?.source?.editable === false;
  if (shape.geometry === "custom" && customPaths.length === 0 && !opaqueSourceBoundCustomGeometry) {
    throw new OfficeKitCodecError(`Presentation shape ${shape.id} requires custom paths.`, [], { code: "invalid_presentation_geometry" });
  }
  const line = normalizePresentationLineStyle(shape.line, { name: `Presentation shape ${shape.id} line` });
  if (shape.geometry !== "line" && (line.head || line.tail)) {
    throw new OfficeKitCodecError(`Presentation shape ${shape.id} arrowheads require geometry line.`, [], { code: "unsupported_presentation_line" });
  }
  const lineWidth = line.width;
  const widthEmu = emuFromPixels(position.width, `${shape.id}.position.width`);
  const heightEmu = emuFromPixels(position.height, `${shape.id}.position.height`);
  if (shape.geometry === "line" && widthEmu === 0n && heightEmu === 0n) {
    throw new OfficeKitCodecError(`Presentation free line ${shape.id} requires at least one positive extent.`, [], { code: "invalid_presentation_frame" });
  }
  if (shape.geometry === "line" && shape.placeholder) {
    throw new OfficeKitCodecError(`Presentation free line ${shape.id} cannot be a placeholder.`, [], { code: "unsupported_presentation_features" });
  }
  const requestedLineRgb = line.style === "none"
    ? ""
    : presentationRgb(presentationLineColor(line, lineWidth > 0 ? "#334155" : "transparent"), `${shape.id}.line.fill`);
  const lineStyle = requestedLineRgb ? line.style : "none";
  const placeholder = !original && shape.placeholder ? sourceFreeSlidePlaceholder(shape) : undefined;
  const textBody = presentationTextBody(shape, originalShape, assetCatalog, customShowLinks);
  const shadow = presentationShadow(shape.shadow, shape.id);
  const accessibility = normalizePresentationAccessibility(shape.accessibility, `Presentation shape ${shape.id}`);
  const sourceFillScheme = String(originalShape?.fillScheme || "");
  const preserveSourceFillScheme = sourceFillScheme && typeof shape.fill === "string" &&
    shape.fill.toLowerCase() === sourceFillScheme.toLowerCase();
  // Imported theme fills such as dk1/dk2 are valid DrawingML but are not
  // authoring color tokens. Keep an unchanged source-bound scheme token in the
  // wire instead of forcing presentationRgb() to interpret it as RGB. This is
  // needed when an unrelated source-bound leaf (for example a run font size)
  // causes the owner shape to be serialized for an edit plan.
  const fillRgb = preserveSourceFillScheme ? "" : presentationRgb(shape.fill, `${shape.id}.fill`);
  const fillOpacityThousandthPercent = presentationFillOpacityThousandthPercent(shape.fill, `${shape.id}.fill`, fillRgb);
  return {
    id: original?.id || shape.id,
    name: shape.name || original?.name || "",
    source: original?.source,
    content: {
      case: "shape",
      value: {
        geometry: shape.geometry,
        leftEmu: sourceBoundFrameEmuFromPixels(position.left, `${shape.id}.position.left`, original),
        topEmu: sourceBoundFrameEmuFromPixels(position.top, `${shape.id}.position.top`, original),
        widthEmu,
        heightEmu,
        text: shape.text?.value || "",
        textBody,
        fillRgb,
        ...(preserveSourceFillScheme ? { fillScheme: sourceFillScheme } : {}),
        ...(fillOpacityThousandthPercent === undefined ? {} : { fillOpacityThousandthPercent }),
        lineRgb: requestedLineRgb,
        lineWidthEmu: BigInt(Math.round(lineWidth * EMU_PER_POINT)),
        lineStyle,
        startArrow: line.head?.type || "",
        endArrow: line.tail?.type || "",
        startArrowWidth: line.head?.width || "",
        startArrowLength: line.head?.length || "",
        endArrowWidth: line.tail?.width || "",
        endArrowLength: line.tail?.length || "",
        lineCap: line.cap || "",
        lineJoin: line.join || "",
        ...(placeholder || {}),
        ...(placeholder || shape.transform == null ? {} : { transform: wirePresentationTransform(shape.transform, `shape ${shape.id}`) }),
        ...(shadow ? { shadow } : {}),
        ...(formulaGraph.adjustments.length ? { customAdjustments: formulaGraph.adjustments } : {}),
        ...(formulaGraph.guides.length ? { customGuides: formulaGraph.guides } : {}),
        ...(customConnectionSites.length ? { customConnectionSites: customConnectionSites.map(presentationCustomGeometryConnectionSiteToWire) } : {}),
        ...(wireCustomAdjustmentHandles.length ? { customAdjustmentHandles: wireCustomAdjustmentHandles } : {}),
        ...(customPaths.length ? { customPaths } : {}),
        ...(textRectangle ? { textRectangle: presentationCustomGeometryTextRectangleToWire(textRectangle) } : {}),
        ...(shape.useBackgroundFill === undefined ? {} : { useBackgroundFill: shape.useBackgroundFill }),
        ...(accessibility ? { accessibility } : {}),
      },
    },
  };
}

function presentationImage(image, original, assetCatalog) {
  const position = image.position || {};
  const importedDataUrl = image[PRESENTATION_IMAGE_DATA_URL_SOURCE];
  const dataUrlDescriptor = Object.getOwnPropertyDescriptor(image, "dataUrl");
  const unchangedImportedAsset = importedDataUrl && importedDataUrl.modified !== true &&
    dataUrlDescriptor?.get === importedDataUrl.get && dataUrlDescriptor?.set === importedDataUrl.set;
  const importedSvgDataUrl = image[PRESENTATION_IMAGE_SVG_DATA_URL_SOURCE];
  const svgDataUrlDescriptor = Object.getOwnPropertyDescriptor(image, "svgDataUrl");
  const unchangedImportedSvgAsset = importedSvgDataUrl && importedSvgDataUrl.modified !== true &&
    svgDataUrlDescriptor?.get === importedSvgDataUrl.get && svgDataUrlDescriptor?.set === importedSvgDataUrl.set;
  const dataUrl = unchangedImportedAsset ? undefined : image.dataUrl;
  const svgDataUrl = unchangedImportedSvgAsset ? undefined : image.svgDataUrl;
  if (!unchangedImportedAsset && !dataUrl) {
    throw new OfficeKitCodecError(`Presentation image ${image.id} requires an embedded dataUrl.`, [], { code: "invalid_presentation_image" });
  }
  if (image.uri || image.geometry !== "rect" || image.borderRadius != null) {
    throw new OfficeKitCodecError(`Presentation image ${image.id} uses external, geometry, or mask semantics outside the bounded PPTX image slice.`, [], { code: "unsupported_presentation_features" });
  }
  const accessibility = normalizePresentationAccessibility(image.accessibility, `Presentation image ${image.id}`);
  const crop = effectivePresentationImageCrop({
    crop: image.crop,
    fit: image.fit,
    dataUrl: image.fit === "stretch" ? dataUrl : image.dataUrl,
    frame: position,
  });
  return {
    id: original?.id || image.id,
    name: image.name || original?.name || "",
    source: original?.source,
    content: {
      case: "image",
      value: {
        assetId: unchangedImportedAsset
          ? assetCatalog.addAsset(importedDataUrl.source.asset)
          : assetCatalog.addDataUrl(dataUrl),
        ...(image.svgDataUrl ? {
          svgAssetId: unchangedImportedSvgAsset
            ? assetCatalog.addAsset(importedSvgDataUrl.source.asset)
            : assetCatalog.addDataUrl(svgDataUrl),
        } : {}),
        altText: accessibility?.description ?? (accessibility ? "" : image.prompt || ""),
        leftEmu: sourceBoundFrameEmuFromPixels(position.left, `${image.id}.position.left`, original),
        topEmu: sourceBoundFrameEmuFromPixels(position.top, `${image.id}.position.top`, original),
        widthEmu: emuFromPixels(position.width, `${image.id}.position.width`),
        heightEmu: emuFromPixels(position.height, `${image.id}.position.height`),
        ...(crop ? { crop: presentationImageCropToWire(crop) } : {}),
        ...(image.transform == null ? {} : { transform: wirePresentationTransform(image.transform, `image ${image.id}`) }),
        ...(accessibility?.title ? { accessibilityTitle: accessibility.title } : {}),
        ...(accessibility?.decorative === undefined ? {} : { accessibilityDecorative: accessibility.decorative }),
      },
    },
  };
}

function presentationImageReadOnlySnapshot(image) {
  return JSON.stringify({
    uri: image.uri,
    contentType: image.contentType,
    geometry: image.geometry,
    borderRadius: image.borderRadius,
  });
}

function distributePresentationTableSize(total, count, ownerLabel) {
  const slots = Number(count);
  if (!Number.isInteger(slots) || slots < 1) {
    throw new OfficeKitCodecError(`Presentation table ${ownerLabel} must contain at least one row and column.`, [], { code: "invalid_presentation_table" });
  }
  const base = total / BigInt(slots);
  const remainder = Number(total % BigInt(slots));
  if (base < 1n) throw new OfficeKitCodecError(`Presentation table ${ownerLabel} is too small for its grid.`, [], { code: "invalid_presentation_table" });
  return Array.from({ length: slots }, (_, index) => base + (index < remainder ? 1n : 0n));
}

function scalePresentationTableSize(values, total, ownerLabel) {
  const source = values.map((value) => BigInt(value));
  const sourceTotal = source.reduce((sum, value) => sum + value, 0n);
  if (!source.length || sourceTotal < 1n) return distributePresentationTableSize(total, source.length, ownerLabel);
  const scaled = source.map((value) => ({ value: (value * total) / sourceTotal, remainder: (value * total) % sourceTotal }));
  let missing = total - scaled.reduce((sum, item) => sum + item.value, 0n);
  for (const index of scaled.map((item, index) => ({ index, remainder: item.remainder })).sort((left, right) => left.remainder === right.remainder ? left.index - right.index : left.remainder > right.remainder ? -1 : 1).map((item) => item.index)) {
    if (missing <= 0n) break;
    scaled[index].value += 1n;
    missing -= 1n;
  }
  if (scaled.some((item) => item.value < 1n)) {
    throw new OfficeKitCodecError(`Presentation table ${ownerLabel} is too small for its imported grid.`, [], { code: "invalid_presentation_table" });
  }
  return scaled.map((item) => item.value);
}

function presentationTable(table, original) {
  const originalTable = original?.content?.case === "table" ? original.content.value : undefined;
  const rows = Number(table.rows);
  const columns = Number(table.columns);
  if (!Number.isInteger(rows) || rows < 1 || rows > 2048 || !Number.isInteger(columns) || columns < 1 || columns > 256 ||
      table.values.length !== rows || table.values.some((row) => !Array.isArray(row) || row.length !== columns)) {
    throw new OfficeKitCodecError(`Presentation table ${table.id} requires a rectangular 1-2048 by 1-256 value matrix.`, [], { code: "invalid_presentation_table" });
  }
  const position = table.position || {};
  const leftEmu = sourceBoundFrameEmuFromPixels(position.left, `${table.id}.position.left`, original);
  const topEmu = sourceBoundFrameEmuFromPixels(position.top, `${table.id}.position.top`, original);
  const widthEmu = emuFromPixels(position.width, `${table.id}.position.width`);
  const heightEmu = emuFromPixels(position.height, `${table.id}.position.height`);
  if (widthEmu < 1n || heightEmu < 1n) throw new OfficeKitCodecError(`Presentation table ${table.id} requires positive width and height.`, [], { code: "invalid_presentation_table" });
  if (originalTable && (originalTable.rows.length !== rows || originalTable.columnWidthsEmu.length !== columns)) {
    throw new OfficeKitCodecError(`Source-preserving PPTX export requires presentation table ${table.id}'s original fixed topology.`, [], { code: "presentation_table_topology_changed" });
  }
  const columnWidthsEmu = originalTable
    ? widthEmu === BigInt(originalTable.widthEmu)
      ? originalTable.columnWidthsEmu.map((value) => BigInt(value))
      : scalePresentationTableSize(originalTable.columnWidthsEmu, widthEmu, `${table.id} columns`)
    : distributePresentationTableSize(widthEmu, columns, `${table.id} columns`);
  const rowHeightsEmu = originalTable
    ? heightEmu === BigInt(originalTable.heightEmu)
      ? originalTable.rows.map((row) => BigInt(row.heightEmu))
      : scalePresentationTableSize(originalTable.rows.map((row) => row.heightEmu), heightEmu, `${table.id} rows`)
    : distributePresentationTableSize(heightEmu, rows, `${table.id} rows`);
  const accessibility = normalizePresentationAccessibility(table.accessibility, `Presentation table ${table.id}`);
  return {
    id: original?.id || table.id,
    name: String(table.name || original?.name || ""),
    source: original?.source,
    content: {
      case: "table",
      value: {
        leftEmu,
        topEmu,
        widthEmu,
        heightEmu,
        columnWidthsEmu,
        rows: table.values.map((row, rowIndex) => ({
          heightEmu: rowHeightsEmu[rowIndex],
          cells: row.map((value) => ({ text: String(value ?? "") })),
        })),
        mergeRanges: table.mergeRanges.map((range) => ({
          startRow: range.startRow,
          endRow: range.endRow,
          startColumn: range.startColumn,
          endColumn: range.endColumn,
        })),
        // Optional protobuf booleans are presence-sensitive.  Keep the
        // canonical representation sparse for false so an imported table
        // whose native flag is absent does not become a semantic edit merely
        // because the JS facade exposes `styleOptions` defaults.
        ...(originalTable?.firstRow === true || (originalTable?.firstRow === undefined && table.styleOptions?.headerRow === true) ? { firstRow: true } : {}),
        ...(originalTable?.bandedRows === true || (originalTable?.bandedRows === undefined && table.styleOptions?.bandedRows === true) ? { bandedRows: true } : {}),
        ...(accessibility ? { accessibility } : {}),
      },
    },
  };
}

function presentationTableReadOnlySnapshot(table) {
  return JSON.stringify({
    id: table.id,
    nativeId: table.nativeId,
    creationId: table.creationId,
    rows: table.rows,
    columns: table.columns,
    style: table.style,
    styleOptions: table.styleOptions,
    border: table.border,
    mergeRanges: table.mergeRanges,
  });
}

// Imported source-bound shapes keep their original wire projection until the
// public model actually changes. This is more than an optimization: rebuilding
// every imported shape forces unrelated native geometry through the authored
// shape subset and can make a safe leaf edit fail before the target is reached.
// The snapshot deliberately covers the complete public Shape projection plus
// its native identity. When it differs, export still goes through the existing
// typed validation and source-binding checks; unchanged shapes alone bypass
// semantic re-projection and retain the codec's original source hash.
function presentationImportedShapeSnapshot(shape) {
  return JSON.stringify({
    nativeId: shape.nativeId,
    creationId: shape.creationId,
    layout: shape.layoutJson(),
  });
}

function presentationImportedGroupSnapshot(group) {
  return JSON.stringify({
    nativeId: group.nativeId,
    creationId: group.creationId,
    layout: group.layoutJson(),
  });
}

function presentationImportedSlideShellSnapshot(slide) {
  return JSON.stringify({
    id: slide.id,
    name: slide.name,
    layoutId: slide.layoutId,
    hidden: slide.hidden,
    background: slide.background,
    transition: slide.transition.toJSON(),
    speakerNotes: slide.speakerNotes.text,
    comments: slide.comments.items.map((comment) => comment.toJSON()),
    elementIds: directSlideElements(slide).map((element) => element.id),
  });
}

function presentationImportedSlideShellWithoutElementsSnapshot(value) {
  const snapshot = JSON.parse(typeof value === "string" ? value : presentationImportedSlideShellSnapshot(value));
  delete snapshot.elementIds;
  return JSON.stringify(snapshot);
}

function presentationElement(element, original, assetCatalog, sourceIdByCloneId, customShowLinks) {
  if (element instanceof GroupShape) return presentationGroup(element, original, assetCatalog, sourceIdByCloneId, customShowLinks);
  if (element instanceof ImageElement) return presentationImage(element, original, assetCatalog);
  if (element instanceof TableElement) return presentationTable(element, original);
  if (element instanceof ChartElement) return presentationChart(element, original);
  if (element?.kind === "connector") return presentationConnector(element, original, sourceIdByCloneId);
  if (element instanceof Shape) return presentationShape(element, original, assetCatalog, customShowLinks);
  if (element?.kind === "nativeObject") return presentationNestedOpaque(element, original);
  throw new OfficeKitCodecError(`Presentation element ${element?.id || "<unknown>"} has no supported OfficeKit wire projection.`, [], { code: "unsupported_presentation_element" });
}

function markPresentationImportedGroupSnapshots(group, source, sourceRevisionSha256) {
  const children = source?.content?.case === "group" ? source.content.value.children || [] : [];
  for (let index = 0; index < group.children.length; index += 1) {
    const child = group.children[index];
    const wire = children[index];
    if (!wire) continue;
    Object.defineProperty(child, PRESENTATION_IMPORTED_GROUP_CHILD, {
      configurable: false,
      enumerable: false,
      writable: false,
      value: Object.freeze({ wire, snapshot: presentationCloneElementSnapshot(child) }),
    });
    const capability = wire.source?.zOrderCapability;
    Object.defineProperty(child, PRESENTATION_ELEMENT_ORDER_CAPABILITY, {
      value: Object.freeze({
        sourceBound: true,
        known: Boolean(capability),
        editable: capability?.supported === true,
        blockedReason: capability
          ? capability.blockedReason || ""
          : "Imported group-child order capability is unavailable.",
        ...(sourceRevisionSha256 ? { sourceRevisionSha256 } : {}),
      }),
    });
    if (child instanceof GroupShape && wire.content?.case === "group") markPresentationImportedGroupSnapshots(child, wire, sourceRevisionSha256);
  }
}

function presentationNestedOpaque(element, original) {
  if (original?.content?.case !== "opaque") {
    throw new OfficeKitCodecError(`Presentation native child ${element.id} has no source-bound opaque payload.`, [], { code: "unsupported_presentation_edit" });
  }
  const placementChanged = element._nativePlacementChanged?.() === true;
  const replacementPending = presentationCloneHasPendingNativeReplacement(element);
  if (!placementChanged && !replacementPending) return original;
  if (!placementChanged || replacementPending || original.source?.editable !== true || element.placementCapability?.supported !== true) {
    throw new OfficeKitCodecError(`Presentation native child ${element.id} changed outside its bounded placement-only profile.`, [], { code: "unsupported_presentation_edit" });
  }
  const frame = element.position;
  const updated = clonePresentationWire(PresentationElementSchema, original);
  updated.content.value.leftEmu = presentationNativePlacementEmu(frame.left, `${element.id}.left`);
  updated.content.value.topEmu = presentationNativePlacementEmu(frame.top, `${element.id}.top`);
  updated.content.value.widthEmu = presentationNativePlacementEmu(frame.width, `${element.id}.width`);
  updated.content.value.heightEmu = presentationNativePlacementEmu(frame.height, `${element.id}.height`);
  return updated;
}

function presentationGroup(group, original, assetCatalog, sourceIdByCloneId, customShowLinks) {
  // A source-bound group is normally emitted from its original wire.  When a
  // sibling edit makes the group require semantic projection, nested picture
  // bullets and image children must still be registered in the fresh request
  // asset catalog; otherwise the native validator sees a valid source asset
  // ID with no corresponding request bytes.
  registerPresentationCloneAssets(group, assetCatalog);
  const originalGroup = original?.content?.case === "group" ? original.content.value : undefined;
  if (!group.children.length) throw new OfficeKitCodecError(`Presentation group ${group.id} requires at least one child.`, [], { code: "invalid_presentation_group" });
  if (originalGroup && originalGroup.children.length !== group.children.length) {
    throw new OfficeKitCodecError(`Source-preserving PPTX export requires presentation group ${group.id}'s original ${originalGroup.children.length}-child topology.`, [], { code: "presentation_group_topology_changed" });
  }
  const frame = group.position || {};
  const childFrame = group.childFrame || {};
  const widthEmu = emuFromPixels(frame.width, `${group.id}.position.width`);
  const heightEmu = emuFromPixels(frame.height, `${group.id}.position.height`);
  const childWidthEmu = emuFromPixels(childFrame.width, `${group.id}.childFrame.width`);
  const childHeightEmu = emuFromPixels(childFrame.height, `${group.id}.childFrame.height`);
  if (widthEmu < 1n || heightEmu < 1n || childWidthEmu < 1n || childHeightEmu < 1n) {
    throw new OfficeKitCodecError(`Presentation group ${group.id} requires positive outer and child extents.`, [], { code: "invalid_presentation_group" });
  }
  const accessibility = normalizePresentationAccessibility(group.accessibility, `Presentation group ${group.id}`);
  return {
    id: original?.id || group.id,
    name: String(group.name || original?.name || ""),
    source: original?.source,
    content: {
      case: "group",
      value: {
        leftEmu: sourceBoundFrameEmuFromPixels(frame.left, `${group.id}.position.left`, original),
        topEmu: sourceBoundFrameEmuFromPixels(frame.top, `${group.id}.position.top`, original),
        widthEmu,
        heightEmu,
        childLeftEmu: signedEmuFromPixels(childFrame.left, `${group.id}.childFrame.left`),
        childTopEmu: signedEmuFromPixels(childFrame.top, `${group.id}.childFrame.top`),
        childWidthEmu,
        childHeightEmu,
        children: group.children.map((child, index) => {
          const imported = child[PRESENTATION_IMPORTED_GROUP_CHILD];
          if (imported && presentationCloneElementSnapshot(child) === imported.snapshot) return imported.wire;
          return presentationElement(child, imported?.wire || originalGroup?.children[index], assetCatalog, sourceIdByCloneId, customShowLinks);
        }),
        ...(accessibility ? { accessibility } : {}),
      },
    },
  };
}

function directSlideElements(slide) {
  return [...slide.elements.items];
}

const SOURCE_BOUND_AUTHORED_OVERLAY_GEOMETRIES = new Set(["textbox", "rect", "roundRect", "ellipse"]);

function assertSourceBoundAuthoredOverlayElement(element, slideIndex) {
  if (element instanceof ImageElement) return;
  if (!(element instanceof Shape)) {
    throw new OfficeKitCodecError(`Presentation slide ${slideIndex + 1} source-bound authored overlays must be canonical textboxes, basic shapes, or embedded rectangular images.`, [], { code: "unsupported_presentation_authored_overlay" });
  }
  if (!SOURCE_BOUND_AUTHORED_OVERLAY_GEOMETRIES.has(element.geometry) || element.placeholder || element.useBackgroundFill !== undefined ||
      element.customPaths?.length || element.customAdjustments?.length || element.customGuides?.length ||
      element.customConnectionSites?.length || element.customAdjustmentHandles?.length || element.textRectangle) {
    throw new OfficeKitCodecError(`Presentation shape ${element.id} uses geometry or layout identity outside the bounded source overlay profile.`, [], { code: "unsupported_presentation_authored_overlay" });
  }
  const paragraphs = [
    ...(element.text?.paragraphs || []),
    ...Object.values(element.text?.inheritedParagraphStyles || {}),
  ];
  if (paragraphs.some((paragraph) => paragraph?.bulletImage || paragraph?.runs?.some((run) => run?.link != null))) {
    throw new OfficeKitCodecError(`Presentation shape ${element.id} cannot add picture or hyperlink relationships through a source-bound authored overlay.`, [], { code: "unsupported_presentation_authored_overlay" });
  }
}

function legacyCommentCoordinate(value, unit, name) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    throw new OfficeKitCodecError(`${name} must be a finite coordinate.`, [], { code: "invalid_presentation_legacy_comment" });
  }
  if (unit === "emu") return Math.round(number);
  if (unit === undefined || unit === "px") return emuFromPixels(number, name);
  throw new OfficeKitCodecError(`${name}.unit must be "px" or "emu".`, [], { code: "invalid_presentation_legacy_comment" });
}

function legacyCommentTimestamp(value, name) {
  const text = String(value ?? "");
  if (!text || Number.isNaN(Date.parse(text))) {
    throw new OfficeKitCodecError(`${name} must be an ISO-8601 timestamp.`, [], { code: "invalid_presentation_legacy_comment" });
  }
  return text;
}

function legacyCommentInteger(value) {
  const number = Number(value);
  return Number.isSafeInteger(number) ? number : undefined;
}

// Keep native legacy-comment evidence only on the imported model. It is used
// to prove an unchanged source-bound export, never exposed as cross-file
// identity and never used to turn an element/thread API into a fake native
// anchor. New legacy comments are slide-level annotations at a fixed position.
function presentationLegacyComments(slide, slideIndex) {
  return slide.comments.items.map((thread, index) => {
    const label = `slide ${slideIndex + 1} legacy comment ${index + 1}`;
    if (thread.nativeFormat && thread.nativeFormat !== "legacy") {
      throw new OfficeKitCodecError(`${label} uses ${thread.nativeFormat} comments, which are outside the legacy PPTX profile.`, [], { code: "unsupported_presentation_comment" });
    }
    if (thread.targetId) {
      throw new OfficeKitCodecError(`${label} targets an element or text range. Legacy PPTX comments are slide-level only.`, [], { code: "unsupported_presentation_comment" });
    }
    if (thread.resolved) {
      throw new OfficeKitCodecError(`${label} is resolved. Legacy PPTX comments do not encode thread state.`, [], { code: "unsupported_presentation_comment" });
    }
    if (!Array.isArray(thread.comments) || thread.comments.length !== 1) {
      throw new OfficeKitCodecError(`${label} must contain exactly one root comment and no replies.`, [], { code: "unsupported_presentation_comment" });
    }
    const comment = thread.comments[0];
    const author = String(comment.author ?? thread.author ?? "").trim();
    if (!author) {
      throw new OfficeKitCodecError(`${label} requires a non-empty author.`, [], { code: "invalid_presentation_legacy_comment" });
    }
    const position = thread.position;
    if (!position || typeof position !== "object") {
      throw new OfficeKitCodecError(`${label} requires an explicit { x, y, unit? } position.`, [], { code: "invalid_presentation_legacy_comment" });
    }
    const anchor = thread.nativeFormat === "legacy" && thread.nativeAnchor && typeof thread.nativeAnchor === "object"
      ? thread.nativeAnchor
      : undefined;
    const anchorPositionXEmu = legacyCommentInteger(anchor?.positionXEmu);
    const anchorPositionYEmu = legacyCommentInteger(anchor?.positionYEmu);
    const positionXEmu = anchorPositionXEmu !== undefined
      ? anchorPositionXEmu
      : legacyCommentCoordinate(position.x, position.unit, `${label}.position.x`);
    const positionYEmu = anchorPositionYEmu !== undefined
      ? anchorPositionYEmu
      : legacyCommentCoordinate(position.y, position.unit, `${label}.position.y`);
    const result = {
      id: thread.id,
      author,
      text: String(comment.text ?? ""),
      createdAt: legacyCommentTimestamp(comment.created ?? thread.created, `${label}.created`),
      positionXEmu,
      positionYEmu,
    };
    const nativeAuthorId = legacyCommentInteger(anchor?.nativeAuthorId);
    const nativeIndex = legacyCommentInteger(anchor?.nativeIndex);
    if (nativeAuthorId !== undefined && nativeAuthorId >= 0) result.nativeAuthorId = nativeAuthorId;
    if (nativeIndex !== undefined && nativeIndex >= 0) result.nativeIndex = nativeIndex;
    return result;
  });
}

// Imported legacy comments do not become a general-purpose thread editor.
// The model snapshot catches every public-field mutation while the wire
// comparison proves that the package-local author/index/position identity is
// still the one OfficeKit imported. The one deliberately mutable leaf is
// the root comment text.
function sourceBoundLegacyCommentTextOnlyEdit(bindingState, slide, slideIndex) {
  if (!bindingState?.wire?.source?.legacyCommentsEditable ||
      !Array.isArray(bindingState.wire.legacyComments) ||
      bindingState.wire.legacyComments.length === 0) return false;
  let original;
  try {
    original = JSON.parse(bindingState.commentSnapshot);
  } catch {
    return false;
  }
  const current = slide.comments.items.map((thread) => thread.toJSON());
  if (!Array.isArray(original) || original.length !== current.length || !original.length) return false;
  const modelIsTextOnly = original.every((sourceThread, index) => {
    const requestedThread = current[index];
    if (!Array.isArray(sourceThread?.comments) || sourceThread.comments.length !== 1 ||
        !Array.isArray(requestedThread?.comments) || requestedThread.comments.length !== 1) return false;
    const sourceText = sourceThread.comments[0]?.text;
    const requestedText = requestedThread.comments[0]?.text;
    if (typeof sourceText !== "string" || typeof requestedText !== "string") return false;
    const sourceWithoutText = structuredClone(sourceThread);
    const requestedWithoutText = structuredClone(requestedThread);
    delete sourceWithoutText.comments[0].text;
    delete requestedWithoutText.comments[0].text;
    return JSON.stringify(sourceWithoutText) === JSON.stringify(requestedWithoutText);
  });
  if (!modelIsTextOnly) return false;
  try {
    const requested = presentationLegacyComments(
      slide,
      Number(bindingState.wire.source?.slideIndex ?? slideIndex),
    );
    return requested.length === bindingState.wire.legacyComments.length &&
      requested.every((comment, index) => {
        const source = bindingState.wire.legacyComments[index];
        return source.id === comment.id &&
          source.author === comment.author &&
          source.createdAt === comment.createdAt &&
          Number(source.positionXEmu || 0) === Number(comment.positionXEmu || 0) &&
          Number(source.positionYEmu || 0) === Number(comment.positionYEmu || 0) &&
          Number(source.nativeAuthorId || 0) === Number(comment.nativeAuthorId || 0) &&
          Number(source.nativeIndex || 0) === Number(comment.nativeIndex || 0);
      });
  } catch {
    return false;
  }
}

const PRESENTATION_MODERN_COMMENT_STATUSES = new Set(["active", "resolved", "closed"]);
const PRESENTATION_MODERN_COMMENT_GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/;

function modernCommentGuid(value, seed, label) {
  const guid = String(value || deterministicPresentationGuid(seed)).toUpperCase();
  if (!PRESENTATION_MODERN_COMMENT_GUID.test(guid)) {
    throw new OfficeKitCodecError(`${label} must be a brace-delimited GUID.`, [], { code: "invalid_presentation_modern_comment" });
  }
  return guid;
}

function modernCommentTimestamp(value, label) {
  const timestamp = String(value ?? "");
  if (!timestamp || Number.isNaN(Date.parse(timestamp))) {
    throw new OfficeKitCodecError(`${label} must be an ISO-8601 timestamp.`, [], { code: "invalid_presentation_modern_comment" });
  }
  return timestamp;
}

function modernCommentInitials(name) {
  const words = String(name || "User").trim().split(/\s+/).filter(Boolean);
  return (words.length > 1 ? words.slice(0, 2).map((word) => [...word][0]) : [...(words[0] || "U")].slice(0, 2)).join("").toUpperCase();
}

function modernCommentAuthor(comment, thread) {
  const person = comment.person || {};
  const author = String(comment.author || person.name || person.displayName || thread.author || "").trim();
  if (!author) throw new OfficeKitCodecError(`Modern comment ${thread.id} requires a non-empty author.`, [], { code: "invalid_presentation_modern_comment" });
  return {
    authorId: modernCommentGuid(comment.authorId || person.id, `author:${author}`, `Modern comment author ${author}`),
    author,
    initials: String(person.initials || comment.initials || modernCommentInitials(author)),
    userId: String(person.userId ?? comment.userId ?? author),
    providerId: String(person.providerId ?? comment.providerId ?? "None"),
  };
}

function flattenedPresentationWireElements(elements) {
  const output = [];
  const visit = (element) => {
    output.push(element);
    if (element.content?.case === "group") for (const child of element.content.value.children || []) visit(child);
  };
  for (const element of elements) visit(element);
  return output;
}

function modernCommentMoniker(wireElement) {
  return {
    shape: "spMk",
    image: "picMk",
    table: "graphicFrameMk",
    chart: "graphicFrameMk",
    connector: "cxnSpMk",
    group: "grpSpMk",
  }[wireElement?.content?.case];
}

function modernCommentCoordinate(value, unit, label) {
  const number = Number(value ?? 0);
  if (!Number.isFinite(number)) throw new OfficeKitCodecError(`${label} must be finite.`, [], { code: "invalid_presentation_modern_comment" });
  if (unit === "px") return emuFromPixels(number, label);
  if (unit === undefined || unit === "emu") return Math.round(number);
  throw new OfficeKitCodecError(`${label}.unit must be "emu" or "px".`, [], { code: "invalid_presentation_modern_comment" });
}

function presentationModernComments(slide, slideIndex, wireElements, originalThreads = []) {
  const flattened = flattenedPresentationWireElements(wireElements);
  const wireById = new Map(flattened.map((element, index) => [element.id, { element, nativeId: index + 2 }]));
  const directIds = new Set(wireElements.map((element) => element.id));
  return slide.comments.items.map((thread, threadIndex) => {
    const label = `slide ${slideIndex + 1} modern comment thread ${threadIndex + 1}`;
    if (thread.nativeFormat && thread.nativeFormat !== "modern") {
      throw new OfficeKitCodecError(`${label} uses ${thread.nativeFormat} comments.`, [], { code: "unsupported_presentation_comment" });
    }
    if (!Array.isArray(thread.comments) || thread.comments.length === 0) {
      throw new OfficeKitCodecError(`${label} requires one root comment.`, [], { code: "invalid_presentation_modern_comment" });
    }
    const target = slide.resolve(thread.targetId);
    const textRange = target?.kind === "textRange";
    const targetElementId = textRange ? target.parentId : target?.id;
    const targetWire = wireById.get(targetElementId);
    if (!targetWire || !directIds.has(targetElementId)) {
      throw new OfficeKitCodecError(`${label} must target a supported top-level slide element or its text range.`, [], { code: "unsupported_presentation_comment" });
    }
    const sourceAnchor = thread.nativeAnchor?.format === "modern" || thread.nativeAnchor?.type ? thread.nativeAnchor : undefined;
    const monikerType = sourceAnchor?.moniker || modernCommentMoniker(targetWire.element);
    if (!monikerType || (textRange && monikerType !== "spMk")) {
      throw new OfficeKitCodecError(`${label} has an unsupported target moniker.`, [], { code: "unsupported_presentation_comment" });
    }
    const nativeSlideId = Number(sourceAnchor?.nativeSlideId ?? sourceAnchor?.slideId ?? 256 + slideIndex);
    const nativeId = Number(sourceAnchor?.nativeId ?? targetWire.nativeId);
    const anchor = {
      kind: textRange ? PresentationModernCommentAnchor_Kind.TEXT_RANGE : PresentationModernCommentAnchor_Kind.ELEMENT,
      nativeSlideId,
      monikers: [{
        type: monikerType,
        nativeId,
        ...(sourceAnchor?.creationId ? { creationId: String(sourceAnchor.creationId).toUpperCase() } : {}),
      }],
      ...(textRange ? {
        textStart: Number(sourceAnchor?.textStart ?? sourceAnchor?.cp ?? 0),
        textLength: Number(sourceAnchor?.textLength ?? sourceAnchor?.length ?? String(target.text ?? "").length),
        ...(sourceAnchor?.contextLength === undefined ? {} : { contextLength: Number(sourceAnchor.contextLength) }),
        ...(sourceAnchor?.contextHash === undefined ? {} : { contextHash: Number(sourceAnchor.contextHash) }),
      } : {}),
    };
    const comments = thread.comments.map((comment, commentIndex) => {
      const author = modernCommentAuthor(comment, thread);
      const status = String(comment.status || (commentIndex === 0 && thread.resolved ? "resolved" : "active")).toLowerCase();
      if (!PRESENTATION_MODERN_COMMENT_STATUSES.has(status)) {
        throw new OfficeKitCodecError(`${label} comment ${commentIndex + 1} has invalid status ${status}.`, [], { code: "invalid_presentation_modern_comment" });
      }
      return {
        id: modernCommentGuid(comment.nativeId || comment.id, `comment:${thread.id}:${commentIndex}`, `${label} comment ${commentIndex + 1}`),
        ...author,
        text: String(comment.text ?? ""),
        createdAt: modernCommentTimestamp(comment.created || thread.created, `${label} comment ${commentIndex + 1}.created`),
        status,
      };
    });
    const original = originalThreads[threadIndex];
    const position = thread.position;
    if (!position || typeof position !== "object") {
      throw new OfficeKitCodecError(`${label} requires an explicit { x, y, unit? } position.`, [], { code: "invalid_presentation_modern_comment" });
    }
    return {
      id: comments[0].id,
      targetId: thread.targetId,
      anchor,
      positionXEmu: modernCommentCoordinate(position.x, position.unit, `${label}.position.x`),
      positionYEmu: modernCommentCoordinate(position.y, position.unit, `${label}.position.y`),
      root: comments[0],
      replies: comments.slice(1),
      ...(original?.source ? { source: original.source } : {}),
    };
  });
}

function presentationThemeSnapshot(theme) {
  return JSON.stringify({
    name: theme.name,
    colors: theme.colors,
    fonts: theme.fonts,
    textStyles: theme.textStyles,
    colorMap: theme.colorMap,
  });
}

function presentationAdvancedSnapshot(presentation) {
  return JSON.stringify({
    theme: JSON.parse(presentationThemeSnapshot(presentation.theme)),
    commentFormat: presentation.commentFormat,
  });
}

function presentationCustomShowLinkContext(currentShows, state) {
  const byId = new Map();
  const byName = new Map();
  for (const show of currentShows) {
    if (byId.has(show.id) || byName.has(show.name)) {
      throw new OfficeKitCodecError("Presentation custom-show hyperlink targets require unique show IDs and names.", [], { code: "invalid_presentation_custom_show" });
    }
    byId.set(show.id, show);
    byName.set(show.name, show);
  }
  return {
    byId,
    byName,
    originalNameById: new Map((state?.customShows || []).map((entry) => [entry.wire.id, entry.wire.name])),
  };
}

function resolvePresentationCustomShowLinkId(name, originalRun, shapeId, context) {
  const originalTarget = originalRun?.hyperlink?.case === "runHyperlink"
    ? originalRun.hyperlink.value?.target
    : undefined;
  let show;
  if (originalTarget?.case === "customShowId" && context?.originalNameById.get(originalTarget.value) === name) {
    // A show rename does not implicitly retarget every referring run. The
    // public model still carries the imported display name, while this stable
    // wire identity follows the same native show across that rename.
    show = context.byId.get(originalTarget.value);
  }
  show ||= context?.byName.get(name);
  if (!show) {
    throw new OfficeKitCodecError(`Presentation shape ${shapeId} references missing custom show ${name}.`, [], { code: "invalid_presentation_hyperlink" });
  }
  return show.id;
}

function presentationCustomShows(presentation, state) {
  const entries = planPresentationCustomShows(presentation).entries;
  if (!state) return entries.map((show) => ({
    id: show.id,
    name: show.name,
    nativeId: show.nativeId,
    slideIds: [...show.slideIds],
  }));
  if (state.customShowsOpaque) {
    if (entries.length) {
      throw new OfficeKitCodecError("The imported PPTX contains an opaque custom-show graph; it can only be preserved unchanged.", [], { code: "unsupported_presentation_custom_show_edit" });
    }
    return [];
  }
  const sourceEntries = state.customShows || [];
  if (entries.length !== sourceEntries.length || entries.some((show, index) => show !== sourceEntries[index].model)) {
    throw new OfficeKitCodecError("Imported PPTX custom shows keep their original count and order; adding, removing, or reordering shows is unsupported.", [], { code: "presentation_custom_show_topology_changed" });
  }
  return entries.map((show, index) => {
    const sourceEntry = sourceEntries[index];
    if (show.id !== sourceEntry.wire.id || show.nativeId !== sourceEntry.wire.nativeId) {
      throw new OfficeKitCodecError(`Imported PPTX custom show ${index + 1} cannot change its facade or native identity.`, [], { code: "presentation_custom_show_topology_changed" });
    }
    return {
      id: sourceEntry.wire.id,
      name: show.name,
      nativeId: show.nativeId,
      slideIds: [...show.slideIds],
      source: sourceEntry.wire.source,
    };
  });
}

function presentationSections(presentation, state) {
  if (state?.sectionsOpaque) {
    if (presentation.sections.items.length) {
      throw new OfficeKitCodecError("The imported PPTX contains an opaque PowerPoint section graph; it can only be preserved unchanged.", [], { code: "unsupported_presentation_section_edit" });
    }
    return [];
  }
  const sourceEntries = state?.sections || [];
  if (state && (presentation.sections.items.length !== sourceEntries.length || presentation.sections.items.some((section, index) => section !== sourceEntries[index].model))) {
    throw new OfficeKitCodecError("Imported PPTX sections keep their original count and order; adding, removing, or reordering sections is unsupported.", [], { code: "presentation_section_topology_changed" });
  }
  const entries = planPresentationSections(presentation, { allowPendingClone: Boolean(state?.clones?.length) }).entries;
  if (!state) return entries.map((section) => ({
    id: section.id,
    name: section.name,
    nativeId: section.nativeId,
    slideIds: [...section.slideIds],
  }));
  return entries.map((section, index) => {
    const sourceEntry = sourceEntries[index];
    if (section.id !== sourceEntry.wire.id || section.nativeId !== sourceEntry.wire.nativeId) {
      throw new OfficeKitCodecError(`Imported PPTX section ${index + 1} cannot change its facade or native GUID identity.`, [], { code: "presentation_section_topology_changed" });
    }
    return {
      id: sourceEntry.wire.id,
      name: section.name,
      nativeId: section.nativeId,
      slideIds: [...section.slideIds],
      source: sourceEntry.wire.source,
    };
  });
}

// Imported comment state belongs to its source SlidePart, not to its current
// display index. Keeping the snapshot per source-state lets a valid deletion
// omit that state while every surviving slide remains strictly read-only.
function presentationSlideCommentSnapshot(slide) {
  return JSON.stringify(slide.comments.items.map((comment) => comment.toJSON()));
}

function unsupportedPresentationFeatures(presentation) {
  const unsupported = [];
  if (presentationThemeSnapshot(presentation.theme) !== DEFAULT_PRESENTATION_THEME) unsupported.push("presentation theme customization");
  if (presentation.masters?.items?.length !== 1) unsupported.push("multiple slide masters");
  const master = presentation.master;
  if (master?.theme) unsupported.push("master theme override");
  if (!["legacy", "modern"].includes(presentation.commentFormat)) unsupported.push(`unknown comment format ${presentation.commentFormat}`);
  for (const slide of presentation.slides?.items || []) {
    const prefix = `slide ${slide.index + 1}`;
    if (presentation.commentFormat === "legacy" && slide.comments?.items?.length) {
      try { presentationLegacyComments(slide, slide.index); }
      catch (error) { unsupported.push(`${prefix} comments (${error.message})`); }
    }
    if (slide.nativeObjects?.items?.length) unsupported.push(`${prefix} native objects`);
  }
  return unsupported;
}

function opaquePresentationSnapshot(object) {
  const oleWorkbook = object.oleWorkbook ? {
    partPath: object.oleWorkbook.partPath,
    contentType: object.oleWorkbook.contentType,
    sourceSha256: object.oleWorkbook.sourceSha256,
    relationshipId: object.oleWorkbook.relationshipId,
  } : undefined;
  const oleOfficePackage = object.oleOfficePackage ? {
    partPath: object.oleOfficePackage.partPath,
    contentType: object.oleOfficePackage.contentType,
    sourceSha256: object.oleOfficePackage.sourceSha256,
    relationshipId: object.oleOfficePackage.relationshipId,
    kind: object.oleOfficePackage.kind,
  } : undefined;
  return JSON.stringify({
    id: object.id,
    name: object.name,
    position: object.position,
    nativeKind: object.nativeKind,
    rawXml: object.rawXml,
    oleWorkbook,
    oleOfficePackage,
    diagramText: object._diagramTextSourceBinding?.(),
    nativeChart: object._nativeChartSourceBinding?.(),
    ...presentationNativeGraphSnapshot(object),
  });
}

function presentationCloneElementSnapshot(element) {
  if (element?.nativeKind || typeof element?._embeddedWorkbookReplacementBytes === "function") {
    return opaquePresentationSnapshot(element);
  }
  if (typeof element?.layoutJson === "function") {
    const layout = element.layoutJson();
    // Image data can be large; the clone guard needs identity, not a second
    // copy of the payload in every pending-clone entry.
    if (typeof layout?.dataUrl === "string") {
      layout.dataUrl = createHash("sha256").update(layout.dataUrl).digest("hex");
    }
    if (typeof layout?.svgDataUrl === "string") {
      layout.svgDataUrl = createHash("sha256").update(layout.svgDataUrl).digest("hex");
    }
    return JSON.stringify(layout);
  }
  return JSON.stringify(element);
}

function presentationCloneHasPendingNativeReplacement(element) {
  return Boolean(
    element?._embeddedWorkbookReplacementBytes?.() ||
    element?._embeddedOfficePackageReplacementBytes?.() ||
    element?._diagramTextReplacement?.(),
  );
}

function registerPresentationCloneAssets(element, assetCatalog) {
  if (element instanceof ImageElement && element.dataUrl) assetCatalog.addDataUrl(element.dataUrl);
  if (element instanceof ImageElement && element.svgDataUrl) assetCatalog.addDataUrl(element.svgDataUrl);
  if (element instanceof Shape) {
    for (const paragraph of element.text?.paragraphs || []) {
      if (paragraph.bulletImage?.dataUrl) assetCatalog.addDataUrl(paragraph.bulletImage.dataUrl);
    }
  }
  if (element instanceof GroupShape) {
    for (const child of element.children) registerPresentationCloneAssets(child, assetCatalog);
  }
}

function presentationOpaque(object, original, snapshot, assetCatalog) {
  if (opaquePresentationSnapshot(object) !== snapshot) {
    // A semantically opaque native object may still expose only its direct
    // frame as a source-issued leaf. Keep the payload byte-bound while
    // allowing a proven placement-only edit; every other model change stays
    // fail-closed. The wire clone retains the original source graph and only
    // replaces the four frame scalars consumed by the native codec.
    if (original?.content?.case === "opaque" &&
        original.source?.editable === true &&
        object.placementCapability?.supported === true &&
        (object._nativePlacementMutationIssued === true || hasPendingPresentationNativeLeafEdit(object)) &&
        opaquePresentationSnapshotWithoutPosition(object) === opaquePresentationSnapshotWithoutPosition(snapshot)) {
      const frame = object.position;
      const leftEmu = presentationNativePlacementEmu(frame.left, "left");
      const topEmu = presentationNativePlacementEmu(frame.top, "top");
      const widthEmu = presentationNativePlacementEmu(frame.width, "width");
      const heightEmu = presentationNativePlacementEmu(frame.height, "height");
      if (leftEmu < 0n || topEmu < 0n || widthEmu <= 0n || heightEmu <= 0n) {
        throw new OfficeKitCodecError(`Presentation native element ${object.id} requires a non-negative position and positive size.`, [], { code: "invalid_presentation_frame" });
      }
      const updated = clonePresentationWire(PresentationElementSchema, original);
      updated.content.value.leftEmu = leftEmu;
      updated.content.value.topEmu = topEmu;
      updated.content.value.widthEmu = widthEmu;
      updated.content.value.heightEmu = heightEmu;
      return updated;
    }
    const message = object.oleWorkbook || object.oleOfficePackage
      ? `Presentation native element ${object.id} changed outside its bounded embedded Office package replacement boundary.`
      : `Presentation native element ${object.id} is source-bound and read-only in OfficeKit 0.2.`;
    throw new OfficeKitCodecError(message, [], { code: "unsupported_presentation_edit" });
  }
  const replacement = object._embeddedWorkbookReplacementBytes?.();
  const officePackageReplacement = object._embeddedOfficePackageReplacementBytes?.();
  const diagramTextReplacement = object._diagramTextReplacement?.();
  if (!replacement && !officePackageReplacement && !diagramTextReplacement) return original;
  const originalOpaque = original?.content?.case === "opaque" ? original.content.value : undefined;
  if (!originalOpaque) {
    throw new OfficeKitCodecError(`Presentation native element ${object.id} has no source-bound opaque payload.`, [], { code: "unsupported_presentation_edit" });
  }
  if (replacement && (!object.oleWorkbook || !originalOpaque.oleWorkbook)) {
    throw new OfficeKitCodecError(`Presentation native element ${object.id} has no source-bound embedded XLSX workbook.`, [], { code: "unsupported_presentation_edit" });
  }
  if (officePackageReplacement && (!object.oleOfficePackage || !originalOpaque.oleOfficePackage)) {
    throw new OfficeKitCodecError(`Presentation native element ${object.id} has no source-bound embedded Office package.`, [], { code: "unsupported_presentation_edit" });
  }
  if (diagramTextReplacement && !originalOpaque.diagramText) {
    throw new OfficeKitCodecError(`Presentation native element ${object.id} has no source-bound SmartArt diagram-text binding.`, [], { code: "unsupported_presentation_edit" });
  }
  return {
    ...original,
    content: {
      case: "opaque",
      value: {
        ...originalOpaque,
        ...(replacement ? { oleWorkbook: {
          ...originalOpaque.oleWorkbook,
          replacementAssetId: assetCatalog.addOleWorkbook(replacement),
        } } : {}),
        ...(officePackageReplacement ? { oleOfficePackage: {
          ...originalOpaque.oleOfficePackage,
          replacementAssetId: assetCatalog.addOleOfficePackage(officePackageReplacement, object.oleOfficePackage),
        } } : {}),
        ...(diagramTextReplacement ? { diagramText: {
          ...originalOpaque.diagramText,
          nodes: diagramTextReplacement.nodes.map((node) => create(PresentationDiagramTextNodeSchema, {
            modelId: node.id,
            text: node.text,
            runTexts: node.runs,
          })),
        } } : {}),
      },
    },
  };
}

function hasPendingPresentationNativeLeafEdit(object) {
  const pending = object?.slide?.presentation?.[PRESENTATION_STATE]?.pendingNativeLeafEdits;
  return Boolean([...pending?.values?.() || []].some((entry) => entry?.leaf?.rootEntry?.model === object &&
    ["leftEmu", "topEmu", "widthEmu", "heightEmu"].includes(entry?.leaf?.leafKind)));
}

function presentationNativePlacementEmu(value, field) {
  const number = Number(value);
  const emu = Math.round(number * EMU_PER_PIXEL);
  if (!Number.isFinite(number) || !Number.isSafeInteger(emu)) {
    throw new OfficeKitCodecError(`Presentation native placement ${field} must be a finite safe coordinate.`, [], { code: "invalid_presentation_frame" });
  }
  return BigInt(emu);
}

function opaquePresentationSnapshotWithoutPosition(value) {
  const snapshot = JSON.parse(typeof value === "string" ? value : opaquePresentationSnapshot(value));
  const copy = { ...snapshot };
  delete copy.position;
  return JSON.stringify(copy);
}

export function presentationEnvelope(presentation, protocolVersion) {
  if (!(presentation instanceof Presentation)) throw new TypeError("exportPptxWithOfficeKit expects a Presentation instance.");
  if (!presentation.slides?.items?.length) throw new OfficeKitCodecError("Presentation must contain at least one slide.", [], { code: "missing_slides" });
  const state = presentation[PRESENTATION_STATE];
  assertTrustedPresentationState(state);
  const sourceStates = presentationSourceSlideStateMap(presentation, state);
  if (!state) {
    const unsupported = unsupportedPresentationFeatures(presentation);
    if (unsupported.length) {
      throw new OfficeKitCodecError(`OfficeKit cannot author these source-free PPTX features: ${unsupported.slice(0, 8).join(", ")}${unsupported.length > 8 ? `, and ${unsupported.length - 8} more` : ""}. Export fails closed; use supported features or import a trustworthy source package for opaque preservation.`, [], { code: "unsupported_presentation_features" });
    }
  } else {
    if (presentationAdvancedSnapshot(presentation) !== state.advancedSnapshot) {
      throw new OfficeKitCodecError("Imported presentation theme and comment wire family are source-bound and read-only in OfficeKit 0.2.", [], { code: "unsupported_presentation_edit" });
    }
    // A source-bound canvas resize is explicit and intentionally narrow: the
    // native codec changes only p:presentation/p:sldSz. It never treats a
    // changed canvas as permission to rescale every slide/master coordinate.
  }

  const customShows = presentationCustomShows(presentation, state);
  const sections = presentationSections(presentation, state);
  const viewProperties = presentationViewPropertiesForEnvelope(presentation, state);
  const customShowLinks = presentationCustomShowLinkContext(customShows, state);
  const assetCatalog = createPresentationAssetCatalog(state?.assets || [], { shareBytes: true });
  const masters = presentationMasters(presentation, state, assetCatalog, customShowLinks);
  const layouts = presentationLayouts(presentation, state, assetCatalog, customShowLinks);
  const slides = presentation.slides.items.map((slide, slideIndex) => {
    const sourceState = sourceStates?.sourceBySlide.get(slide);
    const cloneState = sourceStates?.cloneBySlide.get(slide);
    const bindingState = sourceState || cloneState?.source;
    let retainedEntries = cloneState?.entries || bindingState?.entries;
    let deletedEntries = [];
    let authoredElements = [];
    if (bindingState) {
      if (cloneState && slide.name !== bindingState.name) throw new OfficeKitCodecError(`Source-preserving PPTX export cannot rename pending clone slide ${slideIndex + 1}.`, [], { code: "unsupported_presentation_slide_clone" });
      if ((slide.layoutId || "") !== (bindingState.wire.layoutId || "")) throw new OfficeKitCodecError(`Source-preserving PPTX export cannot change slide ${slideIndex + 1}'s layout binding.`, [], { code: cloneState ? "unsupported_presentation_slide_clone" : "presentation_slide_layout_binding_changed" });
      const commentsChanged = presentationSlideCommentSnapshot(slide) !== bindingState.commentSnapshot;
      const addingLegacyComments = !cloneState &&
        presentation.commentFormat === "legacy" &&
        !bindingState.wire.legacyComments?.length &&
        !bindingState.wire.modernComments?.length &&
        slide.comments.items.length > 0 &&
        bindingState.wire.source?.legacyCommentsAddable === true;
      const editingLegacyComments = !cloneState &&
        presentation.commentFormat === "legacy" &&
        sourceBoundLegacyCommentTextOnlyEdit(bindingState, slide, slideIndex);
      if (commentsChanged && presentation.commentFormat === "legacy" && !addingLegacyComments && !editingLegacyComments) {
        throw new OfficeKitCodecError(`Source-preserving PPTX export can change only existing legacy comment text on slide ${slideIndex + 1}; author, timestamp, coordinate, package-local identity, order, and thread topology are source-bound.`, [], { code: "unsupported_presentation_edit" });
      }
      if (commentsChanged && presentation.commentFormat === "modern" && (!bindingState.wire.modernComments?.length || cloneState)) {
        throw new OfficeKitCodecError(`Imported presentation slide ${slideIndex + 1} comments are source-bound outside the bounded modern text/status edit profile.`, [], { code: "unsupported_presentation_edit" });
      }
      const current = directSlideElements(slide);
      const entries = cloneState?.entries || bindingState.entries;
      const sourceModels = new Set(entries.map((entry) => entry.model));
      const entryByModel = new Map(entries.map((entry) => [entry.model, entry]));
      retainedEntries = current.filter((element) => sourceModels.has(element)).map((element) => entryByModel.get(element));
      deletedEntries = entries.filter((entry) => !current.includes(entry.model));
      authoredElements = current.filter((element) => !sourceModels.has(element));
      const allowedCloneDeletionIds = cloneState?.allowedDeletedIds instanceof Set
        ? cloneState.allowedDeletedIds
        : undefined;
      const typedDeletions = !cloneState && deletedEntries.every((entry) =>
        entry.model[PRESENTATION_ELEMENT_DELETED] === true);
      const authorizedCloneDeletions = Boolean(cloneState && allowedCloneDeletionIds &&
        deletedEntries.length === allowedCloneDeletionIds.size &&
        deletedEntries.every((entry) => allowedCloneDeletionIds.has(entry.wire.id)));
      if (current.length !== retainedEntries.length + authoredElements.length ||
          (cloneState && authoredElements.length > 0) ||
          (!cloneState && !typedDeletions) ||
          (cloneState && deletedEntries.length > 0 && !authorizedCloneDeletions) ||
          (cloneState && allowedCloneDeletionIds && !authorizedCloneDeletions)) {
        throw new OfficeKitCodecError(`Source-preserving PPTX export requires slide ${slideIndex + 1}'s original ${entries.length}-element topology.`, [], { code: cloneState ? "unsupported_presentation_slide_clone" : "presentation_element_topology_changed" });
      }
      const expectedRetainedEntries = allowedCloneDeletionIds
        ? entries.filter((entry) => !allowedCloneDeletionIds.has(entry.wire.id))
        : entries;
      if (cloneState && (retainedEntries.length !== expectedRetainedEntries.length ||
          retainedEntries.some((entry, index) => entry !== expectedRetainedEntries[index]))) {
        throw new OfficeKitCodecError(`Pending presentation clone ${slideIndex + 1} cannot reorder source elements before its first export and reimport.`, [], { code: "unsupported_presentation_slide_clone" });
      }
      if (!cloneState) {
        const authoredIds = new Set();
        for (const element of authoredElements) {
          if (!authoredIds.add(element.id)) {
            throw new OfficeKitCodecError(`Presentation slide ${slideIndex + 1} contains duplicate authored overlay identity ${element.id}.`, [], { code: "invalid_presentation_element" });
          }
          assertSourceBoundAuthoredOverlayElement(element, slideIndex);
        }
      }
      if (!bindingState.wire.speakerNotes && slide.speakerNotes?.text && !bindingState.wire.source?.speakerNotesAddable) {
        throw new OfficeKitCodecError(`Source-preserving PPTX export cannot add speaker notes to slide ${slideIndex + 1} because its presentation notes graph is not safely extensible.`, [], { code: "unsupported_presentation_edit" });
      }
      if (cloneState) {
        for (const entry of cloneState.entries) {
          if (presentationCloneHasPendingNativeReplacement(entry.model) ||
              presentationCloneElementSnapshot(entry.model) !== entry.cloneModelSnapshot) {
            throw new OfficeKitCodecError(`Imported presentation clone ${slideIndex + 1} must remain untouched until it has been exported and imported again.`, [], { code: "unsupported_presentation_slide_clone" });
          }
        }
      }
    }
    const legacyComments = presentation.commentFormat === "legacy"
      ? presentationLegacyComments(slide, Number(bindingState?.wire.source?.slideIndex ?? slideIndex))
      : [];
    const elements = bindingState
      ? [
        ...retainedEntries.map((entry) => {
          // A pending source-bound clone is deliberately immutable until its
          // first export/reimport boundary.  Reuse the exact source wire for
          // every element instead of reserializing the semantic projection;
          // this keeps unsupported geometry and opaque descendants byte-for-
          // byte eligible for the native OPC graph copy.  The C# clone codec
          // performs the ownership proof and copies the original SlidePart.
          if (cloneState) {
            registerPresentationCloneAssets(entry.model, assetCatalog);
            return entry.wire;
          }
          if (entry.wire.content.case === "shape") {
            if (presentationImportedShapeSnapshot(entry.model) === entry.modelSnapshot) {
              return entry.wire;
            }
            if (entry.wire.content.value.placeholder) {
              return presentationSlidePlaceholder(entry.model, entry.wire, entry.placeholderSnapshot, assetCatalog, customShowLinks);
            }
            return presentationShape(entry.model, entry.wire, assetCatalog, customShowLinks);
          }
          if (entry.wire.content.case === "image") {
            if (presentationImageReadOnlySnapshot(entry.model) !== entry.snapshot) {
              throw new OfficeKitCodecError(`Presentation image ${entry.model.id} changed outside its embedded rectangular image boundary.`, [], { code: "unsupported_presentation_edit" });
            }
            return presentationImage(entry.model, entry.wire, assetCatalog);
          }
          if (entry.wire.content.case === "table") {
            if (presentationTableReadOnlySnapshot(entry.model) !== entry.snapshot) {
              throw new OfficeKitCodecError(`Presentation table ${entry.model.id} changed outside its name/frame/plain-text boundary.`, [], { code: "unsupported_presentation_edit" });
            }
            return presentationTable(entry.model, entry.wire);
          }
          if (entry.wire.content.case === "connector") return presentationConnector(entry.model, entry.wire, cloneState?.sourceIdByCloneId);
          if (entry.wire.content.case === "chart") return presentationChart(entry.model, entry.wire);
          if (entry.wire.content.case === "group") {
            // The native validator still checks an unchanged, source-bound
            // group's recursively projected children whenever another object
            // on the presentation is edited.  Register those immutable image
            // and picture-bullet bytes even when the group itself can reuse
            // its original wire.
            registerPresentationCloneAssets(entry.model, assetCatalog);
            if (presentationImportedGroupSnapshot(entry.model) === entry.modelSnapshot) {
              return entry.wire;
            }
            return presentationGroup(entry.model, entry.wire, assetCatalog, cloneState?.sourceIdByCloneId, customShowLinks);
          }
          return presentationOpaque(entry.model, entry.wire, entry.snapshot, assetCatalog);
        }),
        ...authoredElements.map((element) => presentationElement(element, undefined, assetCatalog, undefined, customShowLinks)),
      ]
      : directSlideElements(slide)
        .filter((element) => element instanceof Shape || element instanceof ImageElement || element instanceof TableElement || element instanceof ChartElement || element instanceof GroupShape || slide.connectors.items.includes(element))
        .map((element) => presentationElement(element, undefined, assetCatalog, undefined, customShowLinks));
    const modernComments = presentation.commentFormat === "modern"
      ? presentationModernComments(slide, slideIndex, elements, bindingState?.wire.modernComments || [])
      : [];
    const speakerNotes = presentationSpeakerNotes(
      slide,
      bindingState?.wire.speakerNotes,
      assetCatalog,
      customShowLinks,
    );
    const requested = {
      id: sourceState?.wire.id || slide.id,
      name: slide.name,
      source: sourceState?.wire.source,
      ...(slide.visibilityCapability.known ? { hidden: slide.hidden } : {}),
      ...(slide.layoutId ? { layoutId: slide.layoutId } : {}),
      ...(hasPresentationBackground(slide.background) ? { background: wireBackground(slide.background, `slide ${slideIndex + 1}`, assetCatalog) } : {}),
      ...(slide.transition?.configured ? { transition: wirePresentationTransition(slide.transition) } : {}),
      ...(slide.animations.count ? { animations: slide.animations.items.map((animation) => wirePresentationAnimation(animation, `slide ${slideIndex + 1}`)) } : {}),
      ...(slide.morph.configured ? { morph: wirePresentationMorph(slide.morph.value, `slide ${slideIndex + 1}`) } : {}),
      ...(speakerNotes ? { speakerNotes } : {}),
      ...(legacyComments.length ? { legacyComments } : {}),
      ...(modernComments.length ? { modernComments } : {}),
      elements,
      ...(deletedEntries.length ? {
        elementDeletions: deletedEntries.map((entry) => ({
          id: entry.wire.id,
          source: entry.wire.source,
        })),
      } : {}),
    };
    if (!cloneState) return requested;
    const omittedElementIds = cloneState.allowedDeletedIds instanceof Set ? cloneState.allowedDeletedIds : undefined;
    if (!presentationCloneMatches(requested, cloneState.source.wire, omittedElementIds)) {
      throw new OfficeKitCodecError(`Imported presentation clone ${slideIndex + 1} must remain untouched until it has been exported and imported again.`, [], { code: "unsupported_presentation_slide_clone" });
    }
    delete requested.source;
    requested.cloneSource = cloneState.source.wire.source;
    return requested;
  });
  return {
    protocolVersion,
    family: ArtifactFamily.PRESENTATION,
    source: state?.source,
    assets: assetCatalog.assets(),
    opaqueOpc: state?.opaqueOpc,
    diagnostics: state?.diagnostics || [],
    payload: {
      case: "presentation",
      value: {
        id: presentation.id,
        name: state?.name || "",
        slideWidthEmu: emuFromPixels(presentation.slideSize.width, "slideSize.width"),
        slideHeightEmu: emuFromPixels(presentation.slideSize.height, "slideSize.height"),
        slides,
        masters,
        layouts,
        customShows,
        ...(state?.customShowsOpaque ? { customShowsOpaque: true } : {}),
        sections,
        ...(state?.sectionsOpaque ? { sectionsOpaque: true } : {}),
        ...(viewProperties ? { viewProperties } : {}),
      },
    },
  };
}

function presentationWireBytes(schema, value) {
  return toBinary(schema, create(schema, value));
}

function samePresentationWire(schema, left, right) {
  const leftBytes = presentationWireBytes(schema, left);
  const rightBytes = presentationWireBytes(schema, right);
  return leftBytes.length === rightBytes.length && leftBytes.every((value, index) => value === rightBytes[index]);
}

function clonePresentationWire(schema, value) {
  return fromBinary(schema, presentationWireBytes(schema, value));
}

function presentationTextLeafRuns(shape) {
  const leaves = [];
  let textLeafIndex = 0;
  for (const [paragraphIndex, paragraph] of (shape?.textBody?.paragraphs || []).entries()) {
    for (const [runIndex, run] of (paragraph.runs || []).entries()) {
      if (run.content?.case === "lineBreak") continue;
      if (run.content?.case === "text") {
        leaves.push({ paragraphIndex, runIndex, textLeafIndex, run });
      }
      // Native a:fld also owns one a:t and therefore consumes a leaf index,
      // even though v1 never grants field-text editing through this plan.
      if (run.content?.case === "text" || run.content?.case === "field") textLeafIndex += 1;
    }
  }
  return leaves;
}

function presentationNativeLeafError(code, message, details = {}) {
  return new OfficeKitCodecError(message, [], { code, ...details });
}

function assertNativeLeafTextValue(value) {
  if (typeof value !== "string") throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation native text leaf value must be a string.");
  if (Buffer.byteLength(value, "utf8") > 1_048_576) throw presentationNativeLeafError("presentation_native_leaf_value_too_large", "Presentation native text leaf value exceeds 1048576 UTF-8 bytes.");
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/u.test(value) || hasUnpairedUtf16Surrogate(value)) {
    throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation native text leaf value contains invalid XML text characters or an unpaired UTF-16 surrogate.");
  }
}

function assertNativeLeafFontFamilyValue(value) {
  if (typeof value !== "string" || value.length < 1 || value.length > 255 || value.trim() !== value || value.startsWith("+") || /[\u0000-\u001f\u007f]/u.test(value) || hasUnpairedUtf16Surrogate(value)) {
    throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation native font-family leaf value must be a trimmed literal typeface name of 1 through 255 characters.");
  }
}

function assertNativeLeafBooleanValue(value) {
  if (typeof value !== "boolean") {
    throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation native font-style leaf value must be a boolean.");
  }
}

function hasUnpairedUtf16Surrogate(value) {
  for (let index = 0; index < value.length; index += 1) {
    const unit = value.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const low = value.charCodeAt(index + 1);
      if (!(low >= 0xdc00 && low <= 0xdfff)) return true;
      index += 1;
    } else if (unit >= 0xdc00 && unit <= 0xdfff) return true;
  }
  return false;
}

// Component candidates are design evidence, not a second mutation surface.
// Keep the descriptor semantic and deliberately omit source XML, relationship
// IDs, asset bytes, and absolute package paths. A candidate can therefore tell
// an Agent that a repeated visual primitive exists without granting permission
// to synthesize an unsafe partial graph.
function componentFrameSize(frame) {
  return frame && Number.isFinite(Number(frame.width)) && Number.isFinite(Number(frame.height))
    ? { width: Number(frame.width), height: Number(frame.height) }
    : undefined;
}

function componentParagraphStyle(paragraph = {}) {
  return {
    level: paragraph.level,
    alignment: paragraph.alignment,
    style: paragraph.style,
    bulletCharacter: paragraph.bulletCharacter,
    autoNumber: paragraph.autoNumber,
    bulletNone: paragraph.bulletNone,
    tabStops: paragraph.tabStops,
    marginLeft: paragraph.marginLeft,
    indent: paragraph.indent,
    lineSpacing: paragraph.lineSpacing,
    spaceBefore: paragraph.spaceBefore,
    spaceBeforePercent: paragraph.spaceBeforePercent,
    spaceAfter: paragraph.spaceAfter,
    spaceAfterPercent: paragraph.spaceAfterPercent,
    runs: (paragraph.runs || []).map((run) => ({
      style: run.style,
      link: run.link,
      field: run.field ? { type: run.field.type } : undefined,
      break: run.break === true,
    })),
  };
}

function componentShapeDescriptor(shape) {
  const layout = shape.layoutJson();
  return {
    kind: "shape",
    geometry: layout.geometry,
    frame: componentFrameSize(layout.frame),
    transform: layout.transform,
    customAdjustments: layout.customAdjustments,
    customGuides: layout.customGuides,
    customConnectionSites: layout.customConnectionSites,
    customAdjustmentHandles: layout.customAdjustmentHandles,
    customPaths: layout.customPaths,
    textRectangle: layout.textRectangle,
    bodyProperties: layout.bodyProperties,
    style: layout.style ? { ...layout.style, text: layout.style.text } : undefined,
    paragraphs: (layout.paragraphs || []).map(componentParagraphStyle),
  };
}

function componentImageDescriptor(image) {
  const layout = image.layoutJson();
  return {
    kind: "image",
    frame: componentFrameSize(layout.frame),
    fit: layout.fit,
    crop: layout.crop,
    geometry: layout.geometry,
    borderRadius: layout.borderRadius,
    transform: layout.transform,
  };
}

function componentTableDescriptor(table) {
  const layout = table.layoutJson();
  return {
    kind: "table",
    frame: componentFrameSize(layout.frame),
    rows: layout.rows,
    columns: layout.columns,
    mergeRanges: layout.mergeRanges,
    style: layout.style,
    styleOptions: layout.styleOptions,
  };
}

function componentChartDescriptor(chart) {
  const layout = chart.layoutJson();
  return {
    kind: "chart",
    frame: componentFrameSize(layout.frame),
    chartType: layout.chartType,
    series: (layout.series || []).map((series) => ({
      namePresent: Boolean(series.name),
      axisGroup: series.axisGroup,
      pointCount: Array.isArray(series.values) ? series.values.length : 0,
      style: { color: series.color, line: series.line, marker: series.marker, points: series.points },
    })),
    axes: layout.axes,
    legend: layout.legend,
    dataLabels: layout.dataLabels,
    styleId: layout.styleId,
    varyColors: layout.varyColors,
    barOptions: layout.barOptions,
    lineOptions: layout.lineOptions,
  };
}

function componentConnectorDescriptor(connector) {
  if (connector.startTargetId || connector.endTargetId) {
    return { supported: false, reason: "connector has attached endpoint identities" };
  }
  return {
    kind: "connector",
    frame: componentFrameSize(connector.position),
    geometry: connector.geometry,
    line: connector.line,
    transform: connector.transform,
  };
}

function componentDescriptor(element) {
  try {
    if (element instanceof Shape) return { supported: true, value: componentShapeDescriptor(element) };
    if (element instanceof ImageElement) return { supported: true, value: componentImageDescriptor(element) };
    if (element instanceof TableElement) return { supported: true, value: componentTableDescriptor(element) };
    if (element instanceof ChartElement) return { supported: true, value: componentChartDescriptor(element) };
    if (isPresentationConnectorElement(element)) {
      const descriptor = componentConnectorDescriptor(element);
      return descriptor.supported === false ? descriptor : { supported: true, value: descriptor };
    }
    if (element instanceof GroupShape) {
      const children = [];
      for (const child of element.children) {
        const descriptor = componentDescriptor(child);
        if (descriptor.supported !== true) return { supported: false, reason: descriptor.reason || "group contains an opaque or unsupported descendant" };
        children.push(descriptor.value);
      }
      if (!children.length) return { supported: false, reason: "group has no inspectable descendants" };
      return {
        supported: true,
        value: {
          kind: "group",
          frame: componentFrameSize(element.position),
          childFrame: componentFrameSize(element.childFrame),
          children,
        },
      };
    }
    if (element?.kind === "nativeObject") return { supported: false, reason: "opaque native object" };
    return { supported: false, reason: "element kind is outside the inspect-only component profile" };
  } catch (error) {
    return { supported: false, reason: `descriptor could not be computed: ${error.message}` };
  }
}

function canonicalComponentValue(value) {
  if (Array.isArray(value)) return value.map(canonicalComponentValue);
  if (!value || typeof value !== "object") return value;
  return Object.fromEntries(Object.entries(value)
    .filter(([, entry]) => entry !== undefined)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([key, entry]) => [key, canonicalComponentValue(entry)]));
}

function presentationComponentDescendantIds(entry) {
  const ids = new Set();
  const visit = (wire) => {
    if (!wire?.id || ids.has(wire.id)) return;
    ids.add(wire.id);
    if (wire.content?.case === "group") {
      for (const child of wire.content.value.children || []) visit(child);
    }
  };
  visit(entry?.wire);
  return ids;
}

function presentationComponentEditCapability(state, occurrence, nativeLeafRecords) {
  const sourceState = (state.slides || []).find((entry) => entry.wire?.id === occurrence.slideId);
  const sourceEntry = sourceState?.entries?.find((entry) => entry.wire?.id === occurrence.targetId);
  if (!sourceEntry || !Array.isArray(nativeLeafRecords)) {
    return { supported: false, reason: "component occurrence has no source-bound native-leaf index" };
  }
  const ids = presentationComponentDescendantIds(sourceEntry);
  const leaves = nativeLeafRecords
    .filter((leaf) => ids.has(leaf.targetId) || (leaf.parentGroupId && ids.has(leaf.parentGroupId)))
    .sort((left, right) => left.leafId.localeCompare(right.leafId));
  if (!leaves.length) {
    return { supported: false, reason: "component occurrence has no codec-issued editable leaves" };
  }
  return {
    supported: true,
    mode: "nativeLeaves",
    leafCount: leaves.length,
    leafIds: leaves.map((leaf) => leaf.leafId),
    leafKinds: [...new Set(leaves.map((leaf) => leaf.leafKind))].sort(),
  };
}

function createPresentationComponentCapability(presentation, state) {
  const revisionSha256 = String(state.opaqueOpc?.sourcePackage?.sha256 || state.source?.packageSha256 || "").toLowerCase();
  if (!/^[0-9a-f]{64}$/u.test(revisionSha256)) return undefined;
  const nativeLeafRecords = presentation[PRESENTATION_NATIVE_LEAF_CAPABILITY]?.inspect?.() || [];
  const groups = new Map();
  for (const slideState of state.slides || []) {
    for (const entry of slideState.entries || []) {
      const descriptor = componentDescriptor(entry.model);
      const signatureValue = descriptor.supported
        ? canonicalComponentValue(descriptor.value)
        : { kind: entry.wire.content.case, blocked: descriptor.reason || "opaque or unsupported component graph" };
      const signature = JSON.stringify(signatureValue);
      const group = groups.get(signature) || { signatureValue, descriptor, occurrences: [] };
      group.occurrences.push({
        slide: slideState.slide.index + 1,
        slideId: slideState.wire.id,
        targetId: entry.wire.id,
        elementKind: entry.wire.content.case,
        sourceShapeTreeIndex: Number(entry.wire.source?.shapeTreeIndex),
        expectedElementSha256: entry.wire.source?.elementSha256,
        expectedSemanticSha256: entry.wire.source?.semanticSha256,
      });
      groups.set(signature, group);
    }
  }
  const records = [];
  for (const group of groups.values()) {
    if (group.occurrences.length < 2) continue;
    const sortedOccurrences = [...group.occurrences].sort((left, right) =>
      left.slideId.localeCompare(right.slideId) || left.targetId.localeCompare(right.targetId));
    const repeatedOnOneSlide = new Set(sortedOccurrences.map((occurrence) => occurrence.slideId)).size !== sortedOccurrences.length;
    const blockedReason = group.descriptor.supported !== true
      ? group.descriptor.reason || "opaque or unsupported component graph"
      : repeatedOnOneSlide
        ? "candidate occurs more than once on one slide and is ambiguous without an explicit selection"
        : "";
    const inspectedOccurrences = sortedOccurrences.map((occurrence) => ({
      ...occurrence,
      ownership: { sourceBound: true, closedGraph: !blockedReason, mutableDescendantsShared: false },
      reuseCapability: blockedReason
        ? { supported: false, reason: blockedReason }
        : componentReusePreflight(state, occurrence),
      editCapability: blockedReason
        ? { supported: false, reason: blockedReason }
        : presentationComponentEditCapability(state, occurrence, nativeLeafRecords),
    }));
    const reusableOccurrenceCount = inspectedOccurrences.filter((occurrence) => occurrence.reuseCapability.supported === true).length;
    const editableOccurrenceCount = inspectedOccurrences.filter((occurrence) => occurrence.editCapability.supported === true).length;
    const reuseCapability = blockedReason
      ? { supported: false, reason: blockedReason }
      : reusableOccurrenceCount > 0
        ? { supported: true, occurrenceCount: reusableOccurrenceCount }
        : { supported: false, reason: inspectedOccurrences[0]?.reuseCapability.reason || "no occurrence has an independently deletable source graph" };
    const candidateId = `pc_${createHash("sha256").update(`${revisionSha256}\0${JSON.stringify(sortedOccurrences)}\0${JSON.stringify(group.signatureValue)}`).digest("hex").slice(0, 32)}`;
    records.push(Object.freeze({
      kind: "componentCandidate",
      id: candidateId,
      candidateId,
      sourceRevisionSha256: revisionSha256,
      signature: createHash("sha256").update(JSON.stringify(group.signatureValue), "utf8").digest("hex"),
      descriptor: group.descriptor.supported ? group.signatureValue : { kind: group.signatureValue.kind },
      status: blockedReason ? "blocked" : "inspect-only",
      occurrences: inspectedOccurrences,
      reuseCapability,
      mutationCapability: {
        supported: false,
        reason: blockedReason || reusableOccurrenceCount > 0
          ? "Component candidates are not directly mutable as a whole; choose an occurrence whose editCapability is supported and pass its issued leaves to presentation.editComponentOccurrence, or use presentation.reuseSourceComponent for a new source-bound slide."
          : reuseCapability.reason,
      },
      editCapability: blockedReason
        ? { supported: false, reason: blockedReason }
        : editableOccurrenceCount > 0
          ? { supported: true, mode: "nativeLeaves", occurrenceCount: editableOccurrenceCount }
          : { supported: false, reason: inspectedOccurrences[0]?.editCapability.reason || "no occurrence has codec-issued editable leaves" },
      ...(blockedReason ? { blockedReason } : {}),
    }));
  }
  records.sort((left, right) => left.candidateId.localeCompare(right.candidateId));
  const byId = new Map(records.map((record) => [record.candidateId, record]));
  return Object.freeze({
    inspect: () => records.map((record) => structuredClone(record)),
    resolve(candidateId) {
      if (candidateId == null || candidateId === "") return undefined;
      const record = byId.get(String(candidateId));
      return record ? structuredClone(record) : undefined;
    },
  });
}

function componentReusePreflight(state, occurrence) {
  const sourceState = (state.slides || []).find((entry) => entry.wire?.id === occurrence.slideId);
  const sourceSlide = sourceState?.slide;
  if (!sourceState || !sourceSlide || !Array.isArray(sourceState.entries)) {
    return { supported: false, reason: "source slide is not available in this revision" };
  }
  const sourceEntry = sourceState.entries.find((entry) => entry.wire?.id === occurrence.targetId);
  if (!sourceEntry || sourceEntry.model?.parentGroup) {
    return { supported: false, reason: "candidate is not a direct top-level slide element" };
  }
  if (sourceEntry.wire?.content?.case === "opaque") {
    return { supported: false, reason: "candidate target is an opaque native object" };
  }
  const retainedNativeIds = presentationComponentNativeIds(sourceEntry.model);
  if (new Set(retainedNativeIds).size !== retainedNativeIds.length) {
    return { supported: false, reason: "candidate target contains duplicate native drawing IDs" };
  }
  const retainedNativeIdSet = new Set(retainedNativeIds);
  const deletedNativeIds = sourceState.entries
    .filter((entry) => entry !== sourceEntry)
    .flatMap((entry) => presentationComponentNativeIds(entry.model));
  if (deletedNativeIds.some((id) => retainedNativeIdSet.has(id))) {
    return { supported: false, reason: "candidate target shares a native drawing ID with a deleted sibling" };
  }
  const duplicateIdOnly = (capability) => capability?.supported !== true &&
    /native drawing ID .* ambiguous/u.test(String(capability?.blockedReason || capability?.reason || ""));
  const directElements = [
    ...sourceSlide.connectors.items,
    ...sourceSlide.shapes.items,
    ...sourceSlide.tables.items,
    ...sourceSlide.charts.items,
    ...sourceSlide.images.items,
    ...sourceSlide.groups.items,
    ...sourceSlide.nativeObjects.items,
  ];
  if (!directElements.includes(sourceEntry.model)) {
    return { supported: false, reason: "candidate is not bound to a direct slide element" };
  }
  const cloneCapability = sourceSlide.cloneCapability;
  const sharedOwnedPartsOnly = typeof cloneCapability?.blockedReason === "string" &&
    /^owned part .* is also referenced from /u.test(cloneCapability.blockedReason);
  if (!cloneCapability?.sourceBound || cloneCapability.known !== true ||
      (cloneCapability.supported !== true && !sharedOwnedPartsOnly)) {
    return { supported: false, reason: cloneCapability?.blockedReason || "source slide cannot be cloned safely" };
  }
  const removedSourceIds = new Set(sourceState.entries
    .filter((entry) => entry !== sourceEntry)
    .map((entry) => entry.wire?.id)
    .filter(Boolean));
  for (const entry of sourceState.entries) {
    if (entry === sourceEntry) continue;
    const capability = entry.model?.deletionCapability;
    if (!capability?.sourceBound || capability.known !== true || capability.supported !== true) {
      if (duplicateIdOnly(capability)) continue;
      return { supported: false, reason: `sibling ${entry.wire?.id || "<unknown>"} cannot be removed safely${capability?.blockedReason ? `: ${capability.blockedReason}` : ""}` };
    }
  }
  for (const entry of sourceState.entries) {
    if (entry === sourceEntry || entry.wire?.content?.case !== "connector" || removedSourceIds.has(entry.wire.id)) continue;
    const connector = entry.model;
    if ([connector.startTargetId, connector.endTargetId].some((targetId) => targetId && removedSourceIds.has(targetId))) {
      return { supported: false, reason: "retained connector would point at a removed sibling" };
    }
  }
  return { supported: true };
}

function presentationComponentNativeIds(model) {
  const ids = [];
  const visit = (value) => {
    if (!value || typeof value !== "object") return;
    const nativeId = Number(value.nativeId);
    if (Number.isInteger(nativeId) && nativeId > 0) ids.push(nativeId);
    if (Array.isArray(value.children)) for (const child of value.children) visit(child);
  };
  visit(model);
  return ids;
}

function createPresentationNativeLeafCapability(presentation, state) {
  const revisionSha256 = String(state.opaqueOpc?.sourcePackage?.sha256 || state.source?.packageSha256 || "").toLowerCase();
  if (!/^[0-9a-f]{64}$/u.test(revisionSha256)) return undefined;
  const registry = new Map();
  const records = [];
  const snapshotDataUrlHashes = new Map();
  const snapshotModel = (model) => presentationNativeLeafModelSnapshot(model, snapshotDataUrlHashes);
  const rootSourceSnapshots = new WeakMap();
  const rootSourceSnapshot = (rootEntry) => {
    let snapshot = rootSourceSnapshots.get(rootEntry);
    if (snapshot === undefined) {
      snapshot = snapshotModel(rootEntry.model);
      rootSourceSnapshots.set(rootEntry, snapshot);
    }
    return snapshot;
  };
  const connectedTargetIds = new Set();
  const collectConnectedTargets = (wire) => {
    if (wire.content.case === "connector") {
      if (wire.content.value.startTargetId) connectedTargetIds.add(wire.content.value.startTargetId);
      if (wire.content.value.endTargetId) connectedTargetIds.add(wire.content.value.endTargetId);
    } else if (wire.content.case === "group") {
      for (const child of wire.content.value.children || []) collectConnectedTargets(child);
    }
  };
  for (const slideState of state.slides) for (const entry of slideState.entries) collectConnectedTargets(entry.wire);
  const registerLeaf = ({ wire, model, slideState, shapeTreePath, parentGroupId, rootEntry, leafKind, expectedValue, value, unit, details, compilerBinding, normalize, isNoop, apply }) => {
    const expectedHash = createHash("sha256").update(expectedValue, "utf8").digest("hex");
    const seed = [revisionSha256, slideState.wire.id, wire.id, shapeTreePath.join("/"), wire.source.elementSha256, leafKind,
      details?.textLeafIndex ?? "", details?.nativeLeafIndex ?? "", details?.seriesIndex ?? "", details?.pointIndex ?? "", expectedHash].join("\0");
    const leafId = `nl_${createHash("sha256").update(seed).digest("hex").slice(0, 32)}`;
    const record = Object.freeze({
      kind: "nativeLeaf",
      leafKind,
      id: leafId,
      leafId,
      targetId: wire.id,
      ...(parentGroupId ? { parentGroupId } : {}),
      slide: slideState.slide.index + 1,
      ...(details || {}),
      value,
      ...(unit ? { unit } : {}),
      expectedHash,
      revisionSha256,
    });
    registry.set(leafId, Object.freeze({
      ...record,
      expectedValue,
      wire,
      slideState,
      shapeTreePath: [...shapeTreePath],
      rootEntry,
      rootSourceSnapshot: rootSourceSnapshot(rootEntry),
      compilerBinding,
      normalize,
      isNoop,
      apply,
    }));
    records.push(record);
  };
  const addElementLeaves = (wire, model, slideState, shapeTreePath, parentGroupId, rootEntry) => {
    if (wire.content.case === "group") {
      const children = wire.content.value.children || [];
      if (!Array.isArray(model?.children) || model.children.length !== children.length) return;
      for (let index = 0; index < children.length; index += 1) {
        const child = children[index];
        const childModel = model.children[index];
        if (childModel?.id !== child.id || !child.source) continue;
        addElementLeaves(child, childModel, slideState, [...shapeTreePath, child.source.shapeTreeIndex], wire.id, rootEntry);
      }
      return;
    }
    if (wire.content.case === "opaque") {
      const diagramBinding = wire.content.value.diagramText;
      const modelDiagramBinding = model?._diagramTextSourceBinding?.();
      const currentDiagramLeaves = model?._diagramTextRunRecords?.();
      if (diagramBinding || modelDiagramBinding || currentDiagramLeaves) {
        if (!diagramBinding || !modelDiagramBinding || !Array.isArray(currentDiagramLeaves) ||
            diagramBinding.partPath !== modelDiagramBinding.partPath || diagramBinding.contentType !== modelDiagramBinding.contentType ||
            diagramBinding.sourceSha256 !== modelDiagramBinding.sourceSha256 || diagramBinding.relationshipId !== modelDiagramBinding.relationshipId ||
            diagramBinding.nodes.length !== modelDiagramBinding.nodes.length) return;
        let textLeafIndex = 0;
        const diagramCandidates = [];
        for (let nodeIndex = 0; nodeIndex < diagramBinding.nodes.length; nodeIndex += 1) {
          const node = diagramBinding.nodes[nodeIndex];
          const modelNode = modelDiagramBinding.nodes[nodeIndex];
          const runTexts = node.runTexts?.length ? node.runTexts : [node.text];
          if (node.modelId !== modelNode.id || node.text !== modelNode.text || runTexts.length !== modelNode.runs.length ||
              runTexts.some((text, runIndex) => text !== modelNode.runs[runIndex])) return;
          for (let runIndex = 0; runIndex < runTexts.length; runIndex += 1) {
            const current = currentDiagramLeaves[textLeafIndex];
            if (!current || current.textLeafIndex !== textLeafIndex || current.nodeId !== modelNode.id ||
                current.nodeIndex !== nodeIndex || current.runIndex !== runIndex || current.text !== runTexts[runIndex]) return;
            diagramCandidates.push({
              modelNode,
              nodeIndex,
              runIndex,
              sourceText: runTexts[runIndex],
              textLeafIndex,
            });
            textLeafIndex += 1;
          }
        }
        if (textLeafIndex !== currentDiagramLeaves.length) return;
        for (const candidate of diagramCandidates) {
          registerLeaf({
              wire,
              model,
              slideState,
              shapeTreePath,
              parentGroupId,
              rootEntry,
              leafKind: "diagramText",
              expectedValue: candidate.sourceText,
              value: candidate.sourceText,
              details: { nodeId: candidate.modelNode.id, nodeIndex: candidate.nodeIndex, runIndex: candidate.runIndex, textLeafIndex: candidate.textLeafIndex },
              compilerBinding: {
                targetPartPath: diagramBinding.partPath,
                expectedTargetPartSha256: diagramBinding.sourceSha256,
                relationshipId: diagramBinding.relationshipId,
                diagramModelId: candidate.modelNode.id,
                diagramRunIndex: candidate.runIndex,
              },
              normalize(next) {
                assertNativeLeafTextValue(next);
                if (next.trim() !== next) {
                  throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation diagramText native leaf v1 cannot introduce leading or trailing whitespace.");
                }
                return { raw: next, publicValue: next };
              },
              apply(next) { model._setDiagramTextRun(candidate.modelNode.id, candidate.runIndex, next); },
          });
        }
      }
      const binding = wire.content.value.nativeChart;
      const modelBinding = model?._nativeChartSourceBinding?.();
      const currentLeaves = model?._nativeChartTitleRecords?.();
      const currentDataPoints = model?._nativeChartDataPointRecords?.();
      if (binding || modelBinding || Array.isArray(currentLeaves) || Array.isArray(currentDataPoints)) {
        if (!binding || !modelBinding || !Array.isArray(currentLeaves) ||
            !Array.isArray(currentDataPoints) ||
            binding.partPath !== modelBinding.partPath || binding.contentType !== modelBinding.contentType ||
            binding.sourceSha256 !== modelBinding.sourceSha256 || binding.relationshipId !== modelBinding.relationshipId ||
            binding.titleLeaves.length !== currentLeaves.length || binding.embeddedPackagePartPath !== modelBinding.embeddedPackagePartPath ||
            binding.embeddedPackageSourceSha256 !== modelBinding.embeddedPackageSourceSha256 ||
            binding.embeddedPackageRelationshipId !== modelBinding.embeddedPackageRelationshipId ||
            binding.dataPoints.length !== currentDataPoints.length) return;
        for (let index = 0; index < binding.titleLeaves.length; index += 1) {
        const leaf = binding.titleLeaves[index];
        const current = currentLeaves[index];
        if (leaf.textLeafIndex !== index || current.textLeafIndex !== index || leaf.text !== current.text) return;
        registerLeaf({
          wire,
          model,
          slideState,
          shapeTreePath,
          parentGroupId,
          rootEntry,
          leafKind: "chartTitleText",
          expectedValue: leaf.text,
          value: leaf.text,
          details: { textLeafIndex: index },
          compilerBinding: {
            targetPartPath: binding.partPath,
            expectedTargetPartSha256: binding.sourceSha256,
            relationshipId: binding.relationshipId,
          },
          normalize(next) { assertNativeLeafTextValue(next); return { raw: next, publicValue: next }; },
          apply(next) { model._setNativeChartTitleLeaf(index, next); },
        });
        }
        for (let index = 0; index < binding.dataPoints.length; index += 1) {
        const point = binding.dataPoints[index];
        const current = currentDataPoints[index];
        if (point.seriesIndex !== current.seriesIndex || point.pointIndex !== current.pointIndex || point.value !== current.value ||
            point.formula !== current.formula || point.worksheetPartPath !== current.worksheetPartPath ||
            point.worksheetSourceSha256 !== current.worksheetSourceSha256 || point.worksheetName !== current.worksheetName ||
            point.cellReference !== current.cellReference) return;
        registerLeaf({
          wire,
          model,
          slideState,
          shapeTreePath,
          parentGroupId,
          rootEntry,
          leafKind: "chartDataValue",
          expectedValue: point.value,
          value: Number(point.value),
          details: { seriesIndex: point.seriesIndex, pointIndex: point.pointIndex },
          compilerBinding: {
            targetPartPath: binding.partPath,
            expectedTargetPartSha256: binding.sourceSha256,
            relationshipId: binding.relationshipId,
            embeddedPackagePartPath: binding.embeddedPackagePartPath,
            expectedEmbeddedPackageSha256: binding.embeddedPackageSourceSha256,
            embeddedPackageRelationshipId: binding.embeddedPackageRelationshipId,
            embeddedWorksheetPartPath: point.worksheetPartPath,
            expectedEmbeddedWorksheetSha256: point.worksheetSourceSha256,
            embeddedCellReference: point.cellReference,
            chartSeriesIndex: point.seriesIndex,
            chartPointIndex: point.pointIndex,
            chartFormula: point.formula,
          },
          normalize(next) {
            const token = typeof next === "number" ? String(next) : String(next ?? "").trim();
            if (!/^-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[Ee][+-]?[0-9]+)?$/u.test(token) || !Number.isFinite(Number(token)) || token.length > 128) {
              throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation chartDataValue native leaf requires a finite numeric value.");
            }
            return { raw: token, publicValue: Number(token) };
          },
          apply(next) { model._setNativeChartDataPoint(point.seriesIndex, point.pointIndex, next); },
        });
        }
        return;
      }
      const nativeTextBinding = model?._nativeTextSourceBinding?.();
      const currentNativeTextLeaves = model?._nativeTextRecords?.();
      if (Array.isArray(nativeTextBinding) || Array.isArray(currentNativeTextLeaves)) {
        if (!Array.isArray(nativeTextBinding) || !Array.isArray(currentNativeTextLeaves) ||
            nativeTextBinding.length !== currentNativeTextLeaves.length) return;
        for (let index = 0; index < nativeTextBinding.length; index += 1) {
          const sourceLeaf = nativeTextBinding[index];
          const currentLeaf = currentNativeTextLeaves[index];
          if (sourceLeaf.textLeafIndex !== index || currentLeaf.textLeafIndex !== index || sourceLeaf.text !== currentLeaf.text) return;
          registerLeaf({
            wire,
            model,
            slideState,
            shapeTreePath,
            parentGroupId,
            rootEntry,
            leafKind: "nativeText",
            expectedValue: sourceLeaf.text,
            value: sourceLeaf.text,
            details: { textLeafIndex: index },
            normalize(next) { assertNativeLeafTextValue(next); return { raw: next, publicValue: next }; },
            apply(next) { model._setNativeTextLeaf(index, next); },
          });
        }
        // An opaque native object may expose more than one bounded leaf
        // family (for example group text and child fill tokens). Continue
        // collecting the other families instead of making the first one win.
      }
      const nativeLineBinding = model?._nativeLineSourceBinding?.();
      const currentNativeLineLeaves = model?._nativeLineRecords?.();
      if (Array.isArray(nativeLineBinding) || Array.isArray(currentNativeLineLeaves)) {
        const nativeKind = wire.content.value.nativeKind || model?.nativeKind;
        if (nativeKind !== "connector" || !Array.isArray(nativeLineBinding) ||
            !Array.isArray(currentNativeLineLeaves) || nativeLineBinding.length !== currentNativeLineLeaves.length) return;
        for (let index = 0; index < nativeLineBinding.length; index += 1) {
          const sourceLeaf = nativeLineBinding[index];
          const currentLeaf = currentNativeLineLeaves[index];
          const leafKind = sourceLeaf.leafKind || "lineRgb";
          if (sourceLeaf.lineLeafIndex !== index || currentLeaf.lineLeafIndex !== index ||
              (currentLeaf.leafKind || "lineRgb") !== leafKind ||
              sourceLeaf.expectedValue !== currentLeaf.expectedValue) return;
          registerLeaf({
            wire,
            model,
            slideState,
            shapeTreePath,
            parentGroupId,
            rootEntry,
            leafKind,
            expectedValue: sourceLeaf.expectedValue,
            value: sourceLeaf.value,
            details: { lineLeafIndex: index },
            normalize(next) {
              if (leafKind === "lineWidthEmu") {
                if (typeof next !== "string" && typeof next !== "number") {
                  throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation lineWidthEmu native leaf requires a non-negative integer EMU value.");
                }
                const token = String(next).trim();
                let integer;
                try { integer = BigInt(token); }
                catch { throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation lineWidthEmu native leaf requires a non-negative integer EMU value."); }
                if (String(integer) !== token || integer < 0n || integer > 20_116_800n) {
                  throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation lineWidthEmu native leaf is outside the safe EMU range.");
                }
                return { raw: String(integer), publicValue: Number(integer) };
              }
              if (leafKind === "lineScheme") {
                const token = String(next ?? "").trim();
                const canonical = NATIVE_SCHEME_COLOR_CANONICAL[token.toLowerCase()];
                if (!canonical) throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation lineScheme native leaf requires a supported theme color token.");
                return { raw: canonical, publicValue: canonical };
              }
              const match = /^#?([0-9a-f]{6})$/iu.exec(String(next ?? "").trim());
              if (!match) throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation lineRgb native leaf requires a six-digit RGB color.");
              const normalized = match[1].toUpperCase();
              return { raw: normalized, publicValue: `#${normalized.toLowerCase()}` };
            },
            isNoop(next) {
              return leafKind === "lineScheme"
                ? next === sourceLeaf.expectedValue
                : next.toUpperCase() === sourceLeaf.expectedValue.toUpperCase();
            },
            apply(next) { model._setNativeLineLeaf(index, next); },
          });
        }
        // Continue so a source-bound group can expose both text and style
        // leaves from the same preserved XML root.
      }
      const nativeStyleBinding = model?._nativeStyleSourceBinding?.();
      const currentNativeStyleLeaves = model?._nativeStyleRecords?.();
      if (Array.isArray(nativeStyleBinding) || Array.isArray(currentNativeStyleLeaves)) {
        if (!Array.isArray(nativeStyleBinding) || !Array.isArray(currentNativeStyleLeaves) ||
            nativeStyleBinding.length !== currentNativeStyleLeaves.length) return;
        for (let index = 0; index < nativeStyleBinding.length; index += 1) {
          const sourceLeaf = nativeStyleBinding[index];
          const currentLeaf = currentNativeStyleLeaves[index];
          if (sourceLeaf.nativeLeafIndex !== index || currentLeaf.nativeLeafIndex !== index ||
              sourceLeaf.leafKind !== currentLeaf.leafKind || sourceLeaf.expectedValue !== currentLeaf.expectedValue) return;
          const leafKind = sourceLeaf.leafKind;
          registerLeaf({
            wire,
            model,
            slideState,
            shapeTreePath,
            parentGroupId,
            rootEntry,
            leafKind,
            expectedValue: sourceLeaf.expectedValue,
            value: sourceLeaf.value,
            details: { nativeLeafIndex: index },
            normalize(next) {
              if (leafKind === "lineWidthEmu") {
                if (typeof next !== "string" && typeof next !== "number") {
                  throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation style line width requires a non-negative integer EMU value.");
                }
                const token = String(next).trim();
                let integer;
                try { integer = BigInt(token); }
                catch { throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation style line width requires a non-negative integer EMU value."); }
                if (String(integer) !== token || integer < 0n || integer > 20_116_800n) {
                  throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation style line width is outside the safe EMU range.");
                }
                return { raw: String(integer), publicValue: Number(integer) };
              }
              if (leafKind === "fillScheme" || leafKind === "lineScheme") {
                const canonical = NATIVE_SCHEME_COLOR_CANONICAL[String(next ?? "").trim().toLowerCase()];
                if (!canonical) throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation style scheme native leaf requires a supported theme color token.");
                return { raw: canonical, publicValue: canonical };
              }
              const match = /^#?([0-9a-f]{6})$/iu.exec(String(next ?? "").trim());
              if (!match) throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation style RGB native leaf requires a six-digit RGB color.");
              const normalized = match[1].toUpperCase();
              return { raw: normalized, publicValue: `#${normalized.toLowerCase()}` };
            },
            isNoop(next) {
              return leafKind === "lineWidthEmu"
                ? next === sourceLeaf.expectedValue
                : leafKind === "fillScheme" || leafKind === "lineScheme"
                ? next === sourceLeaf.expectedValue
                : next.toUpperCase() === sourceLeaf.expectedValue.toUpperCase();
            },
            apply(next) { model._setNativeStyleLeaf(index, next); },
          });
        }
      }
      if (wire.content.value.nativeKind === "picture" && wire.source?.editable === true) {
        for (const [field, leafKind] of PRESENTATION_SCALAR_LEAF_FIELDS.filter(([, candidate]) => candidate.endsWith("Emu"))) {
          const raw = String(wire.content.value[field] ?? "");
          if (!/^-?[0-9]+$/u.test(raw)) continue;
          const frameField = ({ leftEmu: "left", topEmu: "top", widthEmu: "width", heightEmu: "height" })[leafKind];
          registerLeaf({
            wire, model, slideState, shapeTreePath, parentGroupId, rootEntry, leafKind,
            expectedValue: raw,
            value: Number(raw),
            unit: "emu",
            normalize(next) {
              if (typeof next !== "string" && typeof next !== "number") {
                throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", `Presentation ${leafKind} native leaf requires an integer EMU value.`);
              }
              const token = String(next).trim();
              let integer;
              try { integer = BigInt(token); }
              catch { throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", `Presentation ${leafKind} native leaf requires an integer EMU value.`); }
              if (String(integer) !== token || integer < BigInt(Number.MIN_SAFE_INTEGER) || integer > BigInt(Number.MAX_SAFE_INTEGER) || ((leafKind === "widthEmu" || leafKind === "heightEmu") && integer <= 0n)) {
                throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", `Presentation ${leafKind} native leaf is outside the safe integer geometry range.`);
              }
              return { raw: String(integer), publicValue: Number(integer) };
            },
            apply(next) { model.position = { ...model.position, [frameField]: Number(next) / EMU_PER_PIXEL }; },
          });
        }
        return;
      }
    }
    const isShape = wire.content.case === "shape";
    const isImage = wire.content.case === "image";
    if ((!isShape && !isImage) || (isShape
      ? wire.source?.editable !== true && wire.source?.textEditable !== true
      : wire.source?.editable !== true)) return;
    const registerImportedShapeColorLeaves = () => {
      if (!isShape || wire.source?.editable === true) return;
      const scheme = NATIVE_SCHEME_COLOR_CANONICAL[String(wire.content.value.fillScheme || "").toLowerCase()];
      if (scheme) {
        registerLeaf({
          wire, model, slideState, shapeTreePath, parentGroupId, rootEntry, leafKind: "fillScheme",
          expectedValue: scheme,
          value: scheme,
          normalize(next) {
            const canonical = NATIVE_SCHEME_COLOR_CANONICAL[String(next ?? "").trim().toLowerCase()];
            if (!canonical) throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation fillScheme native leaf requires a supported theme color token.");
            return { raw: canonical, publicValue: canonical };
          },
          isNoop(next) { return next === scheme; },
          apply(next) { model.fill = next; },
        });
      }
      const lineWidth = String(wire.content.value.lineWidthEmu ?? "");
      // A missing a:ln is projected as the protobuf default 0. Only issue a
      // source-bound width leaf when the source proves a visible, positive
      // width; the export-time proof still requires one bounded a:ln.
      if (/^[1-9][0-9]*$/u.test(lineWidth) && BigInt(lineWidth) <= 20_116_800n) {
        registerLeaf({
          wire, model, slideState, shapeTreePath, parentGroupId, rootEntry, leafKind: "lineWidthEmu",
          expectedValue: lineWidth,
          value: Number(lineWidth),
          unit: "emu",
          normalize(next) {
            if (typeof next !== "string" && typeof next !== "number") {
              throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation lineWidthEmu native leaf requires a non-negative integer EMU value.");
            }
            const token = String(next).trim();
            let integer;
            try { integer = BigInt(token); }
            catch { throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation lineWidthEmu native leaf requires a non-negative integer EMU value."); }
            if (String(integer) !== token || integer < 0n || integer > 20_116_800n) {
              throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation lineWidthEmu native leaf is outside the safe EMU range.");
            }
            return { raw: String(integer), publicValue: Number(integer) };
          },
          isNoop(next) { return next === lineWidth; },
          apply(next) { model.line = { ...(model.line || {}), width: Number(next) / EMU_PER_POINT }; },
        });
      }
      if (wire.source?.textEditable !== true) return;
      for (const [field, leafKind] of [["fillRgb", "fillRgb"], ["lineRgb", "lineRgb"]]) {
        // The semantic projection carries opacity separately when the native
        // color token has an alpha/effect child. The codec deliberately keeps
        // those source-bound paints opaque: do not issue a leaf that the
        // export-time proof will reject after the Agent has selected it.
        if (leafKind === "fillRgb" && wire.content.value.fillOpacityThousandthPercent !== undefined) continue;
        const raw = String(wire.content.value[field] ?? "");
        if (!/^[0-9A-F]{6}$/iu.test(raw)) continue;
        registerLeaf({
          wire, model, slideState, shapeTreePath, parentGroupId, rootEntry, leafKind,
          expectedValue: raw,
          value: `#${raw.toLowerCase()}`,
          normalize(next) {
            const match = /^#?([0-9a-f]{6})$/iu.exec(String(next ?? "").trim());
            if (!match) throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", `Presentation ${leafKind} native leaf requires a six-digit RGB color.`);
            const normalized = match[1].toUpperCase();
            return { raw: normalized, publicValue: `#${normalized.toLowerCase()}` };
          },
          isNoop(next) { return next.toUpperCase() === raw.toUpperCase(); },
          apply(next) {
            if (leafKind === "fillRgb") model.fill = `#${next.toLowerCase()}`;
            else model.line = { ...(model.line || {}), fill: `#${next.toLowerCase()}` };
          },
        });
      }
    };
    if (isShape) {
      for (const leaf of presentationTextLeafRuns(wire.content.value)) {
        const value = leaf.run.content.value;
        registerLeaf({
          wire, model, slideState, shapeTreePath, parentGroupId, rootEntry, leafKind: "text", expectedValue: value, value,
          details: { paragraphIndex: leaf.paragraphIndex, runIndex: leaf.runIndex, textLeafIndex: leaf.textLeafIndex },
          normalize(next) { assertNativeLeafTextValue(next); return { raw: next, publicValue: next }; },
          apply(next) {
            const paragraphs = model.text.paragraphs;
            const run = paragraphs[leaf.paragraphIndex]?.runs?.[leaf.runIndex];
            if (!run || typeof run.text !== "string") throw presentationNativeLeafError("presentation_native_leaf_stale", "Presentation native text leaf no longer resolves to the imported text run.");
            run.text = next;
            model.text.paragraphs = paragraphs;
          },
        });
        const fontSizePoints = Number(leaf.run.fontSizePoints);
        // Slide placeholders keep their owner-local paragraph topology under a
        // stricter source-bound contract; do not issue a font leaf that would
        // make the placeholder exporter mistake an authorized style change for
        // an unapproved topology rewrite. Ordinary shapes remain supported.
        if (!model.placeholder && Number.isFinite(fontSizePoints) && fontSizePoints > 0 && fontSizePoints <= 768) {
          const expectedValue = String(Math.round(fontSizePoints * 100));
          registerLeaf({
            wire, model, slideState, shapeTreePath, parentGroupId, rootEntry, leafKind: "fontSizePoints",
            expectedValue,
            value: fontSizePoints,
            unit: "pt",
            details: { paragraphIndex: leaf.paragraphIndex, runIndex: leaf.runIndex, textLeafIndex: leaf.textLeafIndex },
            normalize(next) {
              if (typeof next !== "string" && typeof next !== "number") {
                throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation fontSizePoints native leaf requires a finite positive point value.");
              }
              const token = String(next).trim();
              if (!/^(?:0|[1-9][0-9]*)(?:\.[0-9]{1,2})?$/u.test(token)) {
                throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation fontSizePoints native leaf requires at most two decimal places.");
              }
              const points = Number(token);
              if (!Number.isFinite(points) || points <= 0 || points > 768) {
                throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation fontSizePoints native leaf is outside the safe point range.");
              }
              const hundredths = Math.round(points * 100);
              return { raw: String(hundredths), publicValue: hundredths / 100 };
            },
            isNoop(next) { return next === expectedValue; },
            apply(next) {
              // Mutate the imported local paragraph graph in place.  Using the
              // public getter/setter here would normalize inherited placeholder
              // formatting and turn a leaf-only edit into a false topology
              // change during source-bound export.
              const paragraphs = model.text._paragraphs;
              const run = paragraphs[leaf.paragraphIndex]?.runs?.[leaf.runIndex];
              if (!run || typeof run !== "object" || run.break || run.field) {
                throw presentationNativeLeafError("presentation_native_leaf_stale", "Presentation fontSizePoints native leaf no longer resolves to the imported text run.");
              }
              run.style = { ...(run.style || {}), fontSize: (Number(next) / 100) / POINTS_PER_PIXEL };
            },
          });
        }
        // A run family is issued only when the imported wire contains an
        // explicit, literal typeface. Theme tokens (for example +mn-lt) and
        // malformed/inherited font graphs remain source-owned. Like the
        // font-size leaf above, skip placeholders so an inherited owner style
        // cannot be mistaken for a local run token during source-bound export.
        if (!model.placeholder) {
          for (const [field, leafKind] of [["fontFamily", "fontFamily"], ["fontFamilyEastAsia", "fontFamilyEastAsia"]]) {
            const family = leaf.run[field];
            if (typeof family !== "string" || family.length < 1 || family.length > 255 || family.trim() !== family || family.startsWith("+") || /[\u0000-\u001f\u007f]/u.test(family) || hasUnpairedUtf16Surrogate(family)) continue;
            registerLeaf({
              wire, model, slideState, shapeTreePath, parentGroupId, rootEntry, leafKind,
              expectedValue: family,
              value: family,
              details: { paragraphIndex: leaf.paragraphIndex, runIndex: leaf.runIndex, textLeafIndex: leaf.textLeafIndex },
              normalize(next) {
                assertNativeLeafFontFamilyValue(next);
                return { raw: next, publicValue: next };
              },
              isNoop(next) { return next === family; },
              apply(next) {
                const paragraphs = model.text._paragraphs;
                const run = paragraphs[leaf.paragraphIndex]?.runs?.[leaf.runIndex];
                if (!run || typeof run !== "object" || run.break || run.field) {
                  throw presentationNativeLeafError("presentation_native_leaf_stale", "Presentation font-family native leaf no longer resolves to the imported text run.");
                }
                run.style = { ...(run.style || {}), [field]: next };
              },
            });
          }
          for (const [field, leafKind] of [["bold", "fontBold"], ["italic", "fontItalic"]]) {
            const enabled = leaf.run[field];
            if (typeof enabled !== "boolean") continue;
            const expectedValue = enabled ? "1" : "0";
            registerLeaf({
              wire, model, slideState, shapeTreePath, parentGroupId, rootEntry, leafKind,
              expectedValue,
              value: enabled,
              details: { paragraphIndex: leaf.paragraphIndex, runIndex: leaf.runIndex, textLeafIndex: leaf.textLeafIndex },
              normalize(next) {
                assertNativeLeafBooleanValue(next);
                return { raw: next ? "1" : "0", publicValue: next };
              },
              isNoop(next) { return next === expectedValue; },
              apply(next) {
                const paragraphs = model.text._paragraphs;
                const run = paragraphs[leaf.paragraphIndex]?.runs?.[leaf.runIndex];
                if (!run || typeof run !== "object" || run.break || run.field) {
                  throw presentationNativeLeafError("presentation_native_leaf_stale", "Presentation font-style native leaf no longer resolves to the imported text run.");
                }
                run.style = { ...(run.style || {}), [field]: next === "1" };
              },
            });
          }
        }
      }
    }
    if (wire.source.editable !== true) {
      registerImportedShapeColorLeaves();
      return;
    }
    const scalarFields = isImage
      ? PRESENTATION_SCALAR_LEAF_FIELDS.filter(([, leafKind]) => leafKind.endsWith("Emu") && leafKind !== "lineWidthEmu")
      : PRESENTATION_SCALAR_LEAF_FIELDS;
    for (const [field, leafKind] of scalarFields) {
      if (leafKind !== "lineWidthEmu" && leafKind.endsWith("Emu") && connectedTargetIds.has(wire.id)) continue;
      const raw = String(wire.content.value[field] ?? "");
      if (leafKind === "lineWidthEmu" && /^[1-9][0-9]*$/u.test(raw) && BigInt(raw) <= 20_116_800n) {
        registerLeaf({
          wire, model, slideState, shapeTreePath, parentGroupId, rootEntry, leafKind, expectedValue: raw,
          value: Number(raw), unit: "emu",
          normalize(next) {
            if (typeof next !== "string" && typeof next !== "number") {
              throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation lineWidthEmu native leaf requires a non-negative integer EMU value.");
            }
            const token = String(next).trim();
            let integer;
            try { integer = BigInt(token); }
            catch { throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation lineWidthEmu native leaf requires a non-negative integer EMU value."); }
            if (String(integer) !== token || integer < 0n || integer > 20_116_800n) {
              throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation lineWidthEmu native leaf is outside the safe EMU range.");
            }
            return { raw: String(integer), publicValue: Number(integer) };
          },
          isNoop(next) { return next === raw; },
          apply(next) { model.line = { ...(model.line || {}), width: Number(next) / EMU_PER_POINT }; },
        });
      } else if ((leafKind === "fillRgb" || leafKind === "lineRgb") && /^[0-9A-F]{6}$/iu.test(raw)) {
        registerLeaf({
          wire, model, slideState, shapeTreePath, parentGroupId, rootEntry, leafKind, expectedValue: raw, value: `#${raw.toLowerCase()}`,
          normalize(next) {
            const match = /^#?([0-9a-f]{6})$/iu.exec(String(next ?? "").trim());
            if (!match) throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", `Presentation ${leafKind} native leaf requires a six-digit RGB color.`);
            const normalized = match[1].toUpperCase();
            return { raw: normalized, publicValue: `#${normalized.toLowerCase()}` };
          },
          isNoop(next) { return next.toUpperCase() === raw.toUpperCase(); },
          apply(next) {
            if (leafKind === "fillRgb") model.fill = `#${next.toLowerCase()}`;
            else model.line = { ...model.line, fill: `#${next.toLowerCase()}` };
          },
        });
      } else if (leafKind !== "lineWidthEmu" && leafKind.endsWith("Emu") && /^-?[0-9]+$/u.test(raw)) {
        const frameField = ({ leftEmu: "left", topEmu: "top", widthEmu: "width", heightEmu: "height" })[leafKind];
        registerLeaf({
          wire, model, slideState, shapeTreePath, parentGroupId, rootEntry, leafKind, expectedValue: raw, value: Number(raw), unit: "emu",
          normalize(next) {
            if (typeof next !== "string" && typeof next !== "number") {
              throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", `Presentation ${leafKind} native leaf requires an integer EMU value.`);
            }
            const token = String(next).trim();
            let integer;
            try { integer = BigInt(token); }
            catch { throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", `Presentation ${leafKind} native leaf requires an integer EMU value.`); }
            if (String(integer) !== token || integer < BigInt(Number.MIN_SAFE_INTEGER) || integer > BigInt(Number.MAX_SAFE_INTEGER) || ((leafKind === "widthEmu" || leafKind === "heightEmu") && integer <= 0n)) {
              throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", `Presentation ${leafKind} native leaf is outside the safe integer geometry range.`);
            }
            return { raw: String(integer), publicValue: Number(integer) };
          },
          apply(next) { model.position = { ...model.position, [frameField]: Number(next) / EMU_PER_PIXEL }; },
        });
      }
    }
  };
  for (const slideState of state.slides) {
    for (const entry of slideState.entries) {
      if (!entry.wire.source) continue;
      addElementLeaves(entry.wire, entry.model, slideState, [entry.wire.source.shapeTreeIndex], undefined, entry);
    }
  }
  const prepare = (targetId, leafId, update) => {
      if (typeof targetId !== "string" || !targetId || typeof leafId !== "string" || !leafId) {
        throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation native-leaf editing requires non-empty targetId and leafId strings.");
      }
      if (!update || typeof update !== "object" || Array.isArray(update)) {
        throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation native-leaf update must be an object.");
      }
      const keys = Object.keys(update).sort();
      if (keys.length !== 2 || keys[0] !== "expectedHash" || keys[1] !== "value") {
        throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation native-leaf update accepts exactly expectedHash and value.");
      }
      const leaf = registry.get(leafId);
      if (!leaf || leaf.targetId !== targetId) {
        throw presentationNativeLeafError("presentation_native_leaf_not_issued", "Presentation native leaf was not issued for this target and source revision.");
      }
      if (update.expectedHash !== leaf.expectedHash) {
        throw presentationNativeLeafError("presentation_native_leaf_stale", "Presentation native leaf expectedHash does not match the imported source revision.");
      }
      const normalized = leaf.normalize(update.value);
      if (leaf.isNoop ? leaf.isNoop(normalized.raw) : normalized.raw === leaf.expectedValue) {
        throw presentationNativeLeafError("presentation_native_leaf_noop", "Presentation native-leaf edit must change its source value.");
      }
      const authorizedBefore = state.authorizedNativeLeafSnapshots?.get(leaf.rootEntry.wire.id) ?? leaf.rootSourceSnapshot;
      if (snapshotModel(leaf.rootEntry.model) !== authorizedBefore) {
        throw presentationNativeLeafError("presentation_native_leaf_concurrent_change", "Presentation native-leaf editing requires the target ownership tree to remain unchanged outside previously authorized native leaves.");
      }
      return { leaf, normalized };
  };
  const applyPrepared = ({ leaf, normalized }) => {
      leaf.apply(normalized.raw);
      state.authorizedNativeLeafSnapshots ??= new Map();
      state.authorizedNativeLeafSnapshots.set(leaf.rootEntry.wire.id, snapshotModel(leaf.rootEntry.model));
      state.pendingNativeLeafEdits ??= new Map();
      state.pendingNativeLeafEdits.set(leaf.leafId, Object.freeze({ leaf, value: normalized.raw }));
      return Object.freeze({
        kind: "nativeLeafEdit",
        targetId: leaf.targetId,
        leafId: leaf.leafId,
        leafKind: leaf.leafKind,
        expectedHash: leaf.expectedHash,
        revisionSha256,
        oldValue: leaf.value,
        value: normalized.publicValue,
        ...(leaf.unit ? { unit: leaf.unit } : {}),
      });
  };
  return Object.freeze({
    inspect: () => records.map((record) => ({ ...record })),
    edit(targetId, leafId, update) {
      return applyPrepared(prepare(targetId, leafId, update));
    },
    editMany(edits) {
      if (!Array.isArray(edits) || edits.length === 0 || edits.length > 256) {
        throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation native-leaf batch requires one through 256 edits.");
      }
      const leafIds = new Set();
      const prepared = edits.map((edit) => {
        if (!edit || typeof edit !== "object" || Array.isArray(edit)) {
          throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation native-leaf batch entries must be objects.");
        }
        const item = prepare(edit.targetId, edit.leafId, {
          expectedHash: edit.expectedHash,
          value: edit.value,
        });
        if (leafIds.has(item.leaf.leafId)) {
          throw presentationNativeLeafError("invalid_presentation_native_leaf_edit", "Presentation native-leaf batch cannot edit the same leaf twice.");
        }
        leafIds.add(item.leaf.leafId);
        return item;
      });
      return Object.freeze(prepared.map(applyPrepared));
    },
  });
}

function compactPresentationSnapshotDataUrls(value, hashes) {
  if (Array.isArray(value)) {
    for (const item of value) compactPresentationSnapshotDataUrls(item, hashes);
    return;
  }
  if (!value || typeof value !== "object") return;
  for (const [key, child] of Object.entries(value)) {
    if (key === "dataUrl" && typeof child === "string") {
      let identity = hashes?.get(child);
      if (identity === undefined) {
        identity = `sha256:${createHash("sha256").update(child).digest("hex")}:${child.length}`;
        hashes?.set(child, identity);
      }
      value[key] = identity;
    } else {
      compactPresentationSnapshotDataUrls(child, hashes);
    }
  }
}

function presentationNativeLeafModelSnapshot(model, dataUrlHashes) {
  const layout = model.layoutJson();
  compactPresentationSnapshotDataUrls(layout, dataUrlHashes);
  return JSON.stringify({
    id: model.id,
    nativeId: model.nativeId,
    creationId: model.creationId,
    layout,
  });
}

function compileIssuedPresentationNativeLeafOperation(pending, sourceSha256) {
  const { leaf, value } = pending;
  const source = leaf.wire.source;
  const slideSource = leaf.slideState.wire.source;
  if (!source || !slideSource || !source.elementSha256 || !source.semanticSha256 || !slideSource.partPath || !slideSource.slideXmlSha256) return undefined;
  const textLeafIndex = leaf.textLeafIndex ?? 0;
  const nativeLeafIndex = leaf.nativeLeafIndex ?? 0;
  const operationSeed = leaf.leafKind === "text"
    ? [sourceSha256, slideSource.partPath, leaf.shapeTreePath.join("/"), textLeafIndex, leaf.expectedValue, value].join("\0")
    : [sourceSha256, slideSource.partPath, leaf.shapeTreePath.join("/"), leaf.leafKind, nativeLeafIndex, JSON.stringify(leaf.compilerBinding || {}), leaf.expectedValue, value].join("\0");
  return {
    operationId: `pptx-${leaf.leafKind}-${createHash("sha256").update(operationSeed).digest("hex").slice(0, 20)}`,
    slideId: leaf.slideState.wire.id,
    slidePartPath: slideSource.partPath,
    expectedSlideSha256: slideSource.slideXmlSha256,
    targetId: leaf.wire.id,
    shapeTreeIndex: leaf.shapeTreePath[0],
    shapeTreePath: [...leaf.shapeTreePath],
    leafKind: leaf.leafKind,
    expectedElementSha256: source.elementSha256,
    expectedSemanticSha256: source.semanticSha256,
    textLeafIndex,
    nativeLeafIndex,
    expectedTextSha256: createHash("sha256").update(leaf.expectedValue, "utf8").digest("hex"),
    expectedValue: leaf.expectedValue,
    value,
    ...(leaf.compilerBinding || {}),
  };
}

function compilePresentationTextLeafOperation(original, requested, sourceSlide, sourceSha256, shapeTreePath) {
  if (original.content.case !== "shape" || requested.content.case !== "shape") return undefined;
  const originalLeaves = presentationTextLeafRuns(original.content.value);
  const requestedLeaves = presentationTextLeafRuns(requested.content.value);
  if (originalLeaves.length !== requestedLeaves.length) return undefined;
  const changed = [];
  for (let index = 0; index < originalLeaves.length; index += 1) {
    const before = originalLeaves[index];
    const after = requestedLeaves[index];
    if (before.paragraphIndex !== after.paragraphIndex || before.runIndex !== after.runIndex || before.textLeafIndex !== after.textLeafIndex) return undefined;
    if (!samePresentationWire(PresentationTextRunSchema, before.run, after.run)) changed.push({ before, after });
  }
  if (changed.length !== 1) return undefined;
  const [{ before, after }] = changed;
  if (before.run.content?.case !== "text" || after.run.content?.case !== "text" || before.run.content.value === after.run.content.value) return undefined;
  const restoredRun = clonePresentationWire(PresentationTextRunSchema, after.run);
  restoredRun.content.value = before.run.content.value;
  if (!samePresentationWire(PresentationTextRunSchema, restoredRun, before.run)) return undefined;
  const restored = clonePresentationWire(PresentationElementSchema, requested);
  restored.content.value.text = original.content.value.text;
  restored.content.value.textBody.paragraphs[before.paragraphIndex].runs[before.runIndex] = before.run;
  if (!samePresentationWire(PresentationElementSchema, restored, original)) return undefined;
  const source = original.source;
  const slideSource = sourceSlide.source;
  if (!source || !slideSource || !source.elementSha256 || !source.semanticSha256 || !slideSource.partPath || !slideSource.slideXmlSha256) return undefined;
  const operationSeed = [sourceSha256, slideSource.partPath, shapeTreePath.join("/"), before.textLeafIndex, before.run.content.value, after.run.content.value].join("\0");
  return {
    operationId: `pptx-text-${createHash("sha256").update(operationSeed).digest("hex").slice(0, 20)}`,
    slideId: sourceSlide.id,
    slidePartPath: slideSource.partPath,
    expectedSlideSha256: slideSource.slideXmlSha256,
    targetId: original.id,
    shapeTreeIndex: shapeTreePath[0],
    shapeTreePath,
    leafKind: "text",
    expectedElementSha256: source.elementSha256,
    expectedSemanticSha256: source.semanticSha256,
    textLeafIndex: before.textLeafIndex,
    expectedTextSha256: createHash("sha256").update(before.run.content.value, "utf8").digest("hex"),
    expectedValue: before.run.content.value,
    value: after.run.content.value,
  };
}

const PRESENTATION_SCALAR_LEAF_FIELDS = Object.freeze([
  Object.freeze(["fillRgb", "fillRgb"]),
  Object.freeze(["lineRgb", "lineRgb"]),
  Object.freeze(["lineWidthEmu", "lineWidthEmu"]),
  Object.freeze(["leftEmu", "leftEmu"]),
  Object.freeze(["topEmu", "topEmu"]),
  Object.freeze(["widthEmu", "widthEmu"]),
  Object.freeze(["heightEmu", "heightEmu"]),
]);

function restoreEquivalentPresentationScalarLeaves(original, requested) {
  if (original.content.case !== "shape" || requested.content.case !== "shape") return requested;
  let restored;
  for (const [field, leafKind] of PRESENTATION_SCALAR_LEAF_FIELDS) {
    if (leafKind !== "fillRgb" && leafKind !== "lineRgb") continue;
    const before = String(original.content.value[field] ?? "");
    const after = String(requested.content.value[field] ?? "");
    if (before !== after && /^[0-9a-f]{6}$/iu.test(before) && /^[0-9a-f]{6}$/iu.test(after) && before.toLowerCase() === after.toLowerCase()) {
      restored ??= clonePresentationWire(PresentationElementSchema, requested);
      restored.content.value[field] = original.content.value[field];
    }
  }
  return restored ?? requested;
}

function compilePresentationScalarLeafOperation(original, requested, sourceSlide, sourceSha256, shapeTreePath) {
  const contentCase = original.content.case;
  if (requested.content.case !== contentCase || !new Set(["shape", "image"]).has(contentCase) || original.source?.editable !== true) return undefined;
  const beforeElement = original.content.value;
  const afterElement = requested.content.value;
  const scalarFields = contentCase === "image"
      ? PRESENTATION_SCALAR_LEAF_FIELDS.filter(([, leafKind]) => leafKind.endsWith("Emu") && leafKind !== "lineWidthEmu")
      : PRESENTATION_SCALAR_LEAF_FIELDS;
  const changed = scalarFields.filter(([field]) => String(beforeElement[field] ?? "") !== String(afterElement[field] ?? ""));
  if (changed.length !== 1) return undefined;
  const [[field, leafKind]] = changed;
  const expectedValue = String(beforeElement[field] ?? "");
  const value = String(afterElement[field] ?? "");
  if ((leafKind === "fillRgb" || leafKind === "lineRgb") && (!/^[0-9A-F]{6}$/iu.test(expectedValue) || !/^[0-9A-F]{6}$/iu.test(value))) return undefined;
  if (leafKind.endsWith("Emu") && (!/^-?[0-9]+$/u.test(expectedValue) || !/^-?[0-9]+$/u.test(value))) return undefined;
  const restored = clonePresentationWire(PresentationElementSchema, requested);
  restored.content.value[field] = beforeElement[field];
  if (!samePresentationWire(PresentationElementSchema, restored, original)) return undefined;
  const source = original.source;
  const slideSource = sourceSlide.source;
  if (!source || !slideSource || !source.elementSha256 || !source.semanticSha256 || !slideSource.partPath || !slideSource.slideXmlSha256) return undefined;
  const operationSeed = [sourceSha256, slideSource.partPath, shapeTreePath.join("/"), leafKind, expectedValue, value].join("\0");
  return {
    operationId: `pptx-${leafKind}-${createHash("sha256").update(operationSeed).digest("hex").slice(0, 20)}`,
    slideId: sourceSlide.id,
    slidePartPath: slideSource.partPath,
    expectedSlideSha256: slideSource.slideXmlSha256,
    targetId: original.id,
    shapeTreeIndex: shapeTreePath[0],
    shapeTreePath,
    leafKind,
    expectedElementSha256: source.elementSha256,
    expectedSemanticSha256: source.semanticSha256,
    textLeafIndex: 0,
    expectedTextSha256: createHash("sha256").update(expectedValue, "utf8").digest("hex"),
    expectedValue,
    value,
  };
}

function compilePresentationImageAssetOperation(original, requested, sourceSlide, sourceSha256, shapeTreePath) {
  if (original.content.case !== "image" || requested.content.case !== "image" || original.source?.editable !== true) return undefined;
  const beforeImage = original.content.value;
  const afterImage = requested.content.value;
  if (!beforeImage.assetId || !afterImage.assetId || beforeImage.assetId === afterImage.assetId) return undefined;
  const restored = clonePresentationWire(PresentationElementSchema, requested);
  restored.content.value.assetId = beforeImage.assetId;
  restored.content.value.crop = beforeImage.crop;
  if (!samePresentationWire(PresentationElementSchema, restored, original)) return undefined;
  const source = original.source;
  const slideSource = sourceSlide.source;
  if (!slideSource || !source.elementSha256 || !source.semanticSha256 || !slideSource.partPath || !slideSource.slideXmlSha256) return undefined;
  const operationSeed = [sourceSha256, slideSource.partPath, shapeTreePath.join("/"), "imageAsset", beforeImage.assetId, afterImage.assetId,
    JSON.stringify(beforeImage.crop || null), JSON.stringify(afterImage.crop || null)].join("\0");
  return {
    operationId: `pptx-imageAsset-${createHash("sha256").update(operationSeed).digest("hex").slice(0, 20)}`,
    slideId: sourceSlide.id,
    slidePartPath: slideSource.partPath,
    expectedSlideSha256: slideSource.slideXmlSha256,
    targetId: original.id,
    shapeTreeIndex: shapeTreePath[0],
    shapeTreePath,
    leafKind: "imageAsset",
    expectedElementSha256: source.elementSha256,
    expectedSemanticSha256: source.semanticSha256,
    textLeafIndex: 0,
    expectedTextSha256: createHash("sha256").update(beforeImage.assetId, "utf8").digest("hex"),
    expectedValue: beforeImage.assetId,
    value: afterImage.assetId,
    imageReplacement: {
      assetId: afterImage.assetId,
      ...(afterImage.crop ? { crop: afterImage.crop } : {}),
    },
  };
}

function compilePresentationImageSvgAssetOperation(original, requested, sourceSlide, sourceSha256, shapeTreePath) {
  if (original.content.case !== "image" || requested.content.case !== "image" || original.source?.editable !== true) return undefined;
  const beforeImage = original.content.value;
  const afterImage = requested.content.value;
  if (!beforeImage.svgAssetId || !afterImage.svgAssetId || beforeImage.svgAssetId === afterImage.svgAssetId) return undefined;
  const restored = clonePresentationWire(PresentationElementSchema, requested);
  restored.content.value.svgAssetId = beforeImage.svgAssetId;
  if (!samePresentationWire(PresentationElementSchema, restored, original)) return undefined;
  const source = original.source;
  const slideSource = sourceSlide.source;
  if (!slideSource || !source.elementSha256 || !source.semanticSha256 || !slideSource.partPath || !slideSource.slideXmlSha256) return undefined;
  const operationSeed = [sourceSha256, slideSource.partPath, shapeTreePath.join("/"), "imageSvgAsset", beforeImage.svgAssetId, afterImage.svgAssetId].join("\0");
  return {
    operationId: `pptx-imageSvgAsset-${createHash("sha256").update(operationSeed).digest("hex").slice(0, 20)}`,
    slideId: sourceSlide.id,
    slidePartPath: slideSource.partPath,
    expectedSlideSha256: slideSource.slideXmlSha256,
    targetId: original.id,
    shapeTreeIndex: shapeTreePath[0],
    shapeTreePath,
    leafKind: "imageSvgAsset",
    expectedElementSha256: source.elementSha256,
    expectedSemanticSha256: source.semanticSha256,
    textLeafIndex: 0,
    expectedTextSha256: createHash("sha256").update(beforeImage.svgAssetId, "utf8").digest("hex"),
    expectedValue: beforeImage.svgAssetId,
    value: afterImage.svgAssetId,
    imageReplacement: {
      assetId: beforeImage.assetId,
      svgAssetId: afterImage.svgAssetId,
    },
  };
}

function compilePresentationElementDeletionOperation(original, sourceSlide, sourceSha256, shapeTreePath) {
  const source = original.source;
  const slideSource = sourceSlide.source;
  const capability = source?.deletionCapability;
  if (shapeTreePath.length !== 1 || capability?.supported !== true || !Number.isSafeInteger(capability.nativeId) || capability.nativeId <= 0 ||
      !slideSource || !source.elementSha256 || !source.semanticSha256 || !slideSource.partPath || !slideSource.slideXmlSha256) return undefined;
  const expectedValue = original.id;
  const operationSeed = [sourceSha256, slideSource.partPath, shapeTreePath[0], "deleteElement", source.elementSha256, capability.nativeId].join("\0");
  return {
    operationId: `pptx-deleteElement-${createHash("sha256").update(operationSeed).digest("hex").slice(0, 20)}`,
    slideId: sourceSlide.id,
    slidePartPath: slideSource.partPath,
    expectedSlideSha256: slideSource.slideXmlSha256,
    targetId: original.id,
    shapeTreeIndex: shapeTreePath[0],
    shapeTreePath,
    leafKind: "deleteElement",
    expectedElementSha256: source.elementSha256,
    expectedSemanticSha256: source.semanticSha256,
    textLeafIndex: 0,
    expectedTextSha256: createHash("sha256").update(expectedValue, "utf8").digest("hex"),
    expectedValue,
    value: "",
    elementDeletion: { expectedNativeId: capability.nativeId },
  };
}

function compilePresentationTableCellOperation(original, requested, sourceSlide, sourceSha256, shapeTreePath) {
  if (original.content.case !== "table" || requested.content.case !== "table" || original.source?.editable !== true) return undefined;
  const beforeTable = original.content.value;
  const afterTable = requested.content.value;
  if (beforeTable.rows.length !== afterTable.rows.length || beforeTable.columnWidthsEmu.length !== afterTable.columnWidthsEmu.length) return undefined;
  const changed = [];
  for (let rowIndex = 0; rowIndex < beforeTable.rows.length; rowIndex += 1) {
    const beforeRow = beforeTable.rows[rowIndex];
    const afterRow = afterTable.rows[rowIndex];
    if (beforeRow.cells.length !== afterRow.cells.length || beforeRow.cells.length !== beforeTable.columnWidthsEmu.length) return undefined;
    for (let columnIndex = 0; columnIndex < beforeRow.cells.length; columnIndex += 1) {
      if (beforeRow.cells[columnIndex].text !== afterRow.cells[columnIndex].text) {
        changed.push({ rowIndex, columnIndex, before: beforeRow.cells[columnIndex].text, after: afterRow.cells[columnIndex].text });
      }
    }
  }
  if (changed.length !== 1) return undefined;
  const [cell] = changed;
  const restored = clonePresentationWire(PresentationElementSchema, requested);
  restored.content.value.rows[cell.rowIndex].cells[cell.columnIndex].text = cell.before;
  if (!samePresentationWire(PresentationElementSchema, restored, original)) return undefined;
  const source = original.source;
  const slideSource = sourceSlide.source;
  if (!slideSource || !source.elementSha256 || !source.semanticSha256 || !slideSource.partPath || !slideSource.slideXmlSha256) return undefined;
  const textLeafIndex = cell.rowIndex * beforeTable.columnWidthsEmu.length + cell.columnIndex;
  const operationSeed = [sourceSha256, slideSource.partPath, shapeTreePath.join("/"), "tableCellText", textLeafIndex, cell.before, cell.after].join("\0");
  return {
    operationId: `pptx-tableCellText-${createHash("sha256").update(operationSeed).digest("hex").slice(0, 20)}`,
    slideId: sourceSlide.id,
    slidePartPath: slideSource.partPath,
    expectedSlideSha256: slideSource.slideXmlSha256,
    targetId: original.id,
    shapeTreeIndex: shapeTreePath[0],
    shapeTreePath,
    leafKind: "tableCellText",
    expectedElementSha256: source.elementSha256,
    expectedSemanticSha256: source.semanticSha256,
    textLeafIndex,
    expectedTextSha256: createHash("sha256").update(cell.before, "utf8").digest("hex"),
    expectedValue: cell.before,
    value: cell.after,
  };
}

function compilePresentationElementEditOperations(original, requested, sourceSlide, sourceSha256, shapeTreePath) {
  requested = restoreEquivalentPresentationScalarLeaves(original, requested);
  if (samePresentationWire(PresentationElementSchema, original, requested)) return [];
  if (original.content.case === "table" && requested.content.case === "table") {
    const operation = compilePresentationTableCellOperation(original, requested, sourceSlide, sourceSha256, shapeTreePath);
    return operation ? [operation] : undefined;
  }
  if (original.content.case === requested.content.case && new Set(["shape", "image"]).has(original.content.case)) {
    const scalarOperation = compilePresentationScalarLeafOperation(original, requested, sourceSlide, sourceSha256, shapeTreePath);
    if (scalarOperation) return [scalarOperation];
    if (original.content.case === "image") {
      const operation = compilePresentationImageAssetOperation(original, requested, sourceSlide, sourceSha256, shapeTreePath);
      if (operation) return [operation];
      const svgOperation = compilePresentationImageSvgAssetOperation(original, requested, sourceSlide, sourceSha256, shapeTreePath);
      return svgOperation ? [svgOperation] : undefined;
    }
    const operation = compilePresentationTextLeafOperation(original, requested, sourceSlide, sourceSha256, shapeTreePath);
    return operation ? [operation] : undefined;
  }
  if (original.content.case !== "group" || requested.content.case !== "group") return undefined;
  const originalChildren = original.content.value.children || [];
  const requestedChildren = requested.content.value.children || [];
  if (originalChildren.length !== requestedChildren.length) return undefined;
  const restored = clonePresentationWire(PresentationElementSchema, requested);
  const operations = [];
  for (let index = 0; index < originalChildren.length; index += 1) {
    const originalChild = originalChildren[index];
    const requestedChild = requestedChildren[index];
    if (originalChild.id !== requestedChild.id || !requestedChild.source || requestedChild.source.shapeTreeIndex !== originalChild.source?.shapeTreeIndex) return undefined;
    if (samePresentationWire(PresentationElementSchema, originalChild, requestedChild)) continue;
    const childOperations = compilePresentationElementEditOperations(
      originalChild,
      requestedChild,
      sourceSlide,
      sourceSha256,
      [...shapeTreePath, originalChild.source.shapeTreeIndex],
    );
    if (!childOperations) return undefined;
    operations.push(...childOperations);
    restored.content.value.children[index] = originalChild;
  }
  if (!operations.length || !samePresentationWire(PresentationElementSchema, restored, original)) return undefined;
  return operations;
}

// Compile the bounded source graph delta into the first stable Edit Plan IR.
// Returning undefined means the change is outside this v1 leaf profile and the
// caller must use an existing typed operation or fail closed; it never grants
// permission to fall back to raw XML or a second Office codec.
export function compilePresentationEditPlan(presentation, protocolVersion) {
  if (!(presentation instanceof Presentation)) throw new TypeError("compilePresentationEditPlan expects a Presentation instance.");
  const state = presentation[PRESENTATION_STATE];
  if (!state || state.clones?.length || presentation.slides.items.length !== state.slides.length ||
      state.slides.some((entry, index) => presentation.slides.items[index] !== entry.slide)) return undefined;
  const snapshot = state.opaqueOpc?.sourcePackage;
  const sourceSha256 = String(snapshot?.sha256 || state.source?.packageSha256 || "").toLowerCase();
  if (!(snapshot?.data instanceof Uint8Array) || !/^[0-9a-f]{64}$/.test(sourceSha256)) return undefined;
  if (state.slides.some((entry) => presentationImportedSlideShellWithoutElementsSnapshot(entry.slide) !==
      presentationImportedSlideShellWithoutElementsSnapshot(entry.shellSnapshot))) return undefined;
  const envelope = presentationEnvelope(presentation, protocolVersion);
  const restoredArtifact = clonePresentationWire(PresentationArtifactSchema, envelope.payload.value);
  const requestedSlides = restoredArtifact.slides;
  const operations = [];
  const pendingByRootId = new Map();
  for (const pending of state.pendingNativeLeafEdits?.values?.() || []) {
    const rootId = pending.leaf.rootEntry.wire.id;
    if (!pendingByRootId.has(rootId)) pendingByRootId.set(rootId, []);
    pendingByRootId.get(rootId).push(pending);
  }
  for (let slideIndex = 0; slideIndex < state.slides.length; slideIndex += 1) {
    const sourceSlide = state.slides[slideIndex];
    const requestedById = new Map(requestedSlides[slideIndex].elements.map((element, elementIndex) => [element.id, { element, elementIndex }]));
    const deletedEntries = [];
    for (let sourceElementIndex = 0; sourceElementIndex < sourceSlide.entries.length; sourceElementIndex += 1) {
      const entry = sourceSlide.entries[sourceElementIndex];
      const requestedEntry = requestedById.get(entry.wire.id);
      if (!requestedEntry) {
        if (entry.model[PRESENTATION_ELEMENT_DELETED] !== true) return undefined;
        const operation = compilePresentationElementDeletionOperation(
          entry.wire,
          sourceSlide.wire,
          sourceSha256,
          [entry.wire.source.shapeTreeIndex],
        );
        if (!operation) return undefined;
        operations.push(operation);
        deletedEntries.push({ sourceElementIndex, wire: entry.wire });
        continue;
      }
      const { element: requested, elementIndex } = requestedEntry;
      const issuedEdits = pendingByRootId.get(entry.wire.id);
      if (issuedEdits?.length) {
        if (presentationNativeLeafModelSnapshot(entry.model) !== state.authorizedNativeLeafSnapshots?.get(entry.wire.id)) return undefined;
        const issuedOperations = issuedEdits.map((pending) => compileIssuedPresentationNativeLeafOperation(pending, sourceSha256));
        if (issuedOperations.some((operation) => !operation)) return undefined;
        operations.push(...issuedOperations);
        restoredArtifact.slides[slideIndex].elements[elementIndex] = entry.wire;
        continue;
      }
      if (samePresentationWire(PresentationElementSchema, entry.wire, requested)) continue;
      const entryOperations = compilePresentationElementEditOperations(
        entry.wire,
        requested,
        sourceSlide.wire,
        sourceSha256,
        [entry.wire.source.shapeTreeIndex],
      );
      if (!entryOperations) return undefined;
      operations.push(...entryOperations);
      restoredArtifact.slides[slideIndex].elements[elementIndex] = entry.wire;
    }
    if (deletedEntries.length) {
      const declared = requestedSlides[slideIndex].elementDeletions || [];
      if (declared.length !== deletedEntries.length || deletedEntries.some(({ wire }) =>
        !declared.some((deletion) => deletion.id === wire.id && samePresentationWire(PresentationElementSourceBindingSchema, deletion.source, wire.source)))) return undefined;
      for (const { sourceElementIndex, wire } of deletedEntries)
        restoredArtifact.slides[slideIndex].elements.splice(sourceElementIndex, 0, wire);
      restoredArtifact.slides[slideIndex].elementDeletions = [];
    }
  }
  operations.sort((left, right) =>
    left.slidePartPath.localeCompare(right.slidePartPath) ||
    left.shapeTreePath.join("/").localeCompare(right.shapeTreePath.join("/"), undefined, { numeric: true }) ||
    left.leafKind.localeCompare(right.leafKind) ||
    (left.textLeafIndex ?? 0) - (right.textLeafIndex ?? 0) ||
    (left.nativeLeafIndex ?? 0) - (right.nativeLeafIndex ?? 0));
  const sourceArtifact = state.sourceArtifact;
  if (!sourceArtifact || !samePresentationWire(PresentationArtifactSchema, restoredArtifact, sourceArtifact)) return undefined;
  const requestedAssetIds = new Set(operations
    .filter((operation) => operation.leafKind === "imageAsset" || operation.leafKind === "imageSvgAsset")
    .map((operation) => operation.leafKind === "imageSvgAsset"
      ? operation.imageReplacement?.svgAssetId
      : operation.imageReplacement?.assetId)
    .filter(Boolean));
  const assets = (envelope.assets || []).filter((asset) => requestedAssetIds.has(asset.id));
  if (assets.length !== requestedAssetIds.size) return undefined;
  return {
    schema: "office-kit/pptx-edit-plan/v1",
    sourceRevisionSha256: sourceSha256,
    sourceBytes: snapshot.data,
    operations,
    wire: {
      expectedSourceSha256: sourceSha256,
      operations,
      assets,
    },
  };
}

export function presentationRequiresNativeLeafEditPlan(presentation) {
  const state = presentation?.[PRESENTATION_STATE];
  return state?.pendingNativeLeafEdits?.size > 0 || state?.slides?.some(({ entries }) =>
    entries.some(({ model }) => model[PRESENTATION_ELEMENT_DELETED] === true));
}

function presentationNativeKind(elementName) {
  return ({ pic: "picture", graphicFrame: "graphicFrame", grpSp: "group", cxnSp: "connector", contentPart: "contentPart" })[elementName] || elementName || "nativeObject";
}

function modelRun(run, customShowLinks) {
  const hyperlink = run.hyperlink?.case === "runHyperlink" ? modelHyperlink(run.hyperlink.value, customShowLinks) : undefined;
  const content = run.content?.case === "lineBreak"
    ? { break: true }
    : run.content?.case === "field"
      ? { field: { id: run.content.value.id, type: run.content.value.type, text: run.content.value.text } }
      : { text: run.content?.case === "text" ? run.content.value : "" };
  return {
    ...content,
    style: {
      ...(run.bold === undefined ? {} : { bold: run.bold }),
      ...(run.italic === undefined ? {} : { italic: run.italic }),
      ...(run.fontSizePoints === undefined ? {} : { fontSize: run.fontSizePoints / POINTS_PER_PIXEL }),
      ...(run.fontFamily === undefined ? {} : { fontFamily: run.fontFamily }),
      ...(run.fontFamilyEastAsia === undefined ? {} : { fontFamilyEastAsia: run.fontFamilyEastAsia }),
      ...(run.colorRgb === undefined ? {} : { color: `#${run.colorRgb}` }),
    },
    ...(hyperlink ? { link: hyperlink } : {}),
  };
}

function modelHyperlink(link, customShowLinks) {
  const customShowName = link.target?.case === "customShowId" ? customShowLinks?.get(link.target.value) : undefined;
  if (link.target?.case === "customShowId" && !customShowName) {
    throw new OfficeKitCodecError(`Presentation run hyperlink references missing custom show ${link.target.value}.`, [], { code: "invalid_presentation_artifact" });
  }
  const target = link.target?.case === "uri"
    ? { uri: link.target.value }
    : link.target?.case === "slideId"
      ? { slideId: link.target.value }
      : link.target?.case === "action"
        ? { action: link.target.value }
        : link.target?.case === "customShowId"
          ? { customShow: customShowName }
        : {};
  return {
    ...target,
    ...(link.tooltip === undefined ? {} : { tooltip: link.tooltip }),
    ...(link.targetFrame === undefined ? {} : { targetFrame: link.targetFrame }),
    ...(link.history === undefined ? {} : { history: link.history }),
    ...(link.highlightClick === undefined ? {} : { highlightClick: link.highlightClick }),
    ...(link.returnToSlide === undefined ? {} : { returnToSlide: link.returnToSlide }),
  };
}

function modelBullet(bullet, assetCatalog) {
  if (bullet?.case === "noBullet") return { bulletNone: true };
  if (bullet?.case === "bulletCharacter") return { bulletCharacter: bullet.value };
  if (bullet?.case === "autoNumber") return { autoNumber: { type: bullet.value.scheme, ...(bullet.value.startAt === undefined ? {} : { startAt: bullet.value.startAt }) } };
  if (bullet?.case === "pictureBullet") {
    if (bullet.value.source?.case === "assetId") return { bulletImage: { dataUrl: assetCatalog.dataUrl(bullet.value.source.value), relationshipMode: "embed" } };
    if (bullet.value.source?.case === "uri") return { bulletImage: { uri: validatePictureBulletUri(bullet.value.source.value), relationshipMode: "link" } };
    throw new OfficeKitCodecError("Presentation picture bullet has no source.", [], { code: "invalid_presentation_asset" });
  }
  return {};
}

function modelBulletStyle(paragraph) {
  return {
    ...(paragraph.bulletFont?.case === "bulletFontFamily" ? { bulletFont: paragraph.bulletFont.value } : {}),
    ...(paragraph.bulletFont?.case === "bulletFontFollowText" ? { bulletFontFollowText: true } : {}),
    ...(paragraph.bulletColor?.case === "bulletColorRgb" ? { bulletColor: `#${paragraph.bulletColor.value}` } : {}),
    ...(paragraph.bulletColor?.case === "bulletColorScheme" ? { bulletColor: paragraph.bulletColor.value } : {}),
    ...(paragraph.bulletColor?.case === "bulletColorFollowText" ? { bulletColorFollowText: true } : {}),
    ...(paragraph.bulletSize?.case === "bulletSizePoints" ? { bulletSize: paragraph.bulletSize.value / POINTS_PER_PIXEL } : {}),
    ...(paragraph.bulletSize?.case === "bulletSizePercent" ? { bulletSizePercent: paragraph.bulletSize.value } : {}),
    ...(paragraph.bulletSize?.case === "bulletSizeFollowText" ? { bulletSizeFollowText: true } : {}),
  };
}

function modelParagraphLayout(paragraph) {
  return {
    ...(paragraph.leftMargin?.case === "marginLeftEmu" ? { marginLeft: Number(paragraph.leftMargin.value) / EMU_PER_PIXEL } : {}),
    ...(paragraph.indentation?.case === "indentEmu" ? { indent: Number(paragraph.indentation.value) / EMU_PER_PIXEL } : {}),
  };
}

function modelParagraphSpacing(paragraph) {
  return {
    ...(paragraph.lineSpacing?.case === "lineSpacingPoints" ? { lineSpacing: paragraph.lineSpacing.value / POINTS_PER_PIXEL } : {}),
    ...(paragraph.lineSpacing?.case === "lineSpacingMultiplier" ? { lineSpacing: paragraph.lineSpacing.value } : {}),
    ...(paragraph.spaceBefore?.case === "spaceBeforePoints" ? { spaceBefore: paragraph.spaceBefore.value / POINTS_PER_PIXEL } : {}),
    ...(paragraph.spaceBefore?.case === "spaceBeforeMultiplier" ? { spaceBeforePercent: paragraph.spaceBefore.value } : {}),
    ...(paragraph.spaceAfter?.case === "spaceAfterPoints" ? { spaceAfter: paragraph.spaceAfter.value / POINTS_PER_PIXEL } : {}),
    ...(paragraph.spaceAfter?.case === "spaceAfterMultiplier" ? { spaceAfterPercent: paragraph.spaceAfter.value } : {}),
  };
}

function modelDefaultRunStyle(paragraph) {
  if (paragraph.defaultRunStyle?.case !== "defaultRunProperties") return {};
  const style = paragraph.defaultRunStyle.value;
  return {
    ...(style.bold === undefined ? {} : { bold: style.bold }),
    ...(style.italic === undefined ? {} : { italic: style.italic }),
    ...(style.fontSizePoints === undefined ? {} : { fontSize: style.fontSizePoints / POINTS_PER_PIXEL }),
    ...(style.fontFamily === undefined ? {} : { fontFamily: style.fontFamily }),
    ...(style.fontFamilyEastAsia === undefined ? {} : { fontFamilyEastAsia: style.fontFamilyEastAsia }),
    ...(style.color?.case === "colorRgb" ? { color: `#${style.color.value}` } : {}),
    ...(style.color?.case === "colorScheme" ? { color: style.color.value } : {}),
  };
}

function modelParagraph(paragraph, assetCatalog, { includeRuns = true, customShowLinks } = {}) {
  return {
    ...(includeRuns ? { runs: paragraph.runs.map((run) => modelRun(run, customShowLinks)) } : {}),
    level: paragraph.level ?? 0,
    ...(paragraph.alignment ? { alignment: paragraph.alignment } : {}),
    ...modelBullet(paragraph.bullet, assetCatalog),
    ...modelBulletStyle(paragraph),
    ...modelParagraphLayout(paragraph),
    ...modelParagraphSpacing(paragraph),
    ...(paragraph.tabStops?.length ? { tabStops: paragraph.tabStops.map((tab) => ({ position: Number(tab.positionEmu) / EMU_PER_PIXEL, alignment: tab.alignment })) } : {}),
    style: modelDefaultRunStyle(paragraph),
  };
}

function modelText(shape, assetCatalog, customShowLinks) {
  if (!shape.textBody) return shape.text;
  return shape.textBody.paragraphs.map((paragraph) => modelParagraph(paragraph, assetCatalog, { customShowLinks }));
}

function modelListStyles(shape, assetCatalog) {
  if (!shape.textBody) return {};
  return Object.fromEntries(shape.textBody.listStyles.map((style) => [style.level, modelParagraph(style, assetCatalog, { includeRuns: false })]));
}

function modelTextBodyProperties(shape) {
  const source = shape.textBody?.bodyProperties;
  if (!source) return {};
  const properties = {};
  const insets = {};
  for (const [key, choice] of [["left", source.leftInset], ["top", source.topInset], ["right", source.rightInset], ["bottom", source.bottomInset]]) {
    if (choice?.case?.endsWith("InsetEmu")) insets[key] = Number(choice.value) / EMU_PER_PIXEL;
  }
  if (Object.keys(insets).length) properties.insets = insets;
  if (source.anchor?.case === "verticalAnchor") properties.anchor = source.anchor.value;
  if (source.wrapping?.case === "wrap") properties.wrap = source.wrapping.value;
  if (source.autoFit?.case === "autoFitMode") properties.autoFit = source.autoFit.value;
  const normalAutoFit = {};
  if (source.normalAutoFit?.fontScale?.case === "fontScale1000") normalAutoFit.fontScale = source.normalAutoFit.fontScale.value / 1000;
  if (source.normalAutoFit?.lineSpacingReduction?.case === "lineSpacingReduction1000") normalAutoFit.lineSpacingReduction = source.normalAutoFit.lineSpacingReduction.value / 1000;
  if (Object.keys(normalAutoFit).length) properties.normalAutoFit = normalAutoFit;
  if (source.rotation?.case === "rotationAngle60000") properties.rotation = source.rotation.value / ROTATION_UNITS_PER_DEGREE;
  if (source.verticalText?.case === "verticalTextMode") properties.verticalText = source.verticalText.value;
  if (source.verticalOverflow?.case === "verticalOverflowMode") properties.verticalOverflow = source.verticalOverflow.value;
  if (source.horizontalOverflow?.case === "horizontalOverflowMode") properties.horizontalOverflow = source.horizontalOverflow.value;
  const columns = {};
  if (source.columnCount?.case === "columns") columns.count = source.columnCount.value;
  if (source.columnSpacing?.case === "columnSpacingEmu") columns.spacing = Number(source.columnSpacing.value) / EMU_PER_PIXEL;
  if (source.columnDirection?.case === "rightToLeftColumns") columns.rightToLeft = source.columnDirection.value;
  if (Object.keys(columns).length) properties.columns = columns;
  if (source.uprightText?.case === "upright") properties.upright = source.uprightText.value;
  return properties;
}

function modelMasterTextStyles(source, assetCatalog) {
  return Object.fromEntries(MASTER_STYLE_KINDS.map(([kind, levelsField]) => [
    kind,
    Object.fromEntries((source?.textStyles?.[levelsField] || []).map((style) => [style.level, modelParagraph(style, assetCatalog, { includeRuns: false })])),
  ]));
}

function modelPlaceholder(source, assetCatalog, customShowLinks) {
  const shape = { textBody: source.textBody };
  const transform = modelPlaceholderTransform(source.directFrame);
  return {
    id: source.id,
    name: source.name,
    type: source.type,
    idx: source.index,
    ...(source.directFrame ? { position: modelPlaceholderFrame(source.directFrame) } : {}),
    ...(Object.keys(transform).length ? { transform } : {}),
    text: modelText(shape, assetCatalog, customShowLinks),
    paragraphStyles: modelListStyles(shape, assetCatalog),
    textBodyProperties: modelTextBodyProperties(shape),
  };
}

function modelPlaceholderFrame(frame) {
  return {
    left: Number(frame.leftEmu) / EMU_PER_PIXEL,
    top: Number(frame.topEmu) / EMU_PER_PIXEL,
    width: Number(frame.widthEmu) / EMU_PER_PIXEL,
    height: Number(frame.heightEmu) / EMU_PER_PIXEL,
  };
}

function modelPresentationTransform(frame) {
  const transform = {};
  if (frame?.rotationAngle60000 != null) transform.rotationDegrees = frame.rotationAngle60000 / ROTATION_UNITS_PER_DEGREE;
  if (frame?.flipHorizontal != null) transform.flipHorizontal = Boolean(frame.flipHorizontal);
  if (frame?.flipVertical != null) transform.flipVertical = Boolean(frame.flipVertical);
  return transform;
}

function modelCustomGeometryGuides(guides) {
  return (guides || []).map((guide) => ({ name: guide.name, formula: guide.formula }));
}

function modelCustomGeometryPoint(point) {
  return {
    x: point.xReference ?? Number(point.x),
    y: point.yReference ?? Number(point.y),
  };
}

function modelCustomGeometryArc(arc) {
  return {
    widthRadius: arc.widthRadiusReference ?? Number(arc.widthRadius),
    heightRadius: arc.heightRadiusReference ?? Number(arc.heightRadius),
    startAngle: arc.startAngleReference ?? arc.startAngle,
    sweepAngle: arc.sweepAngleReference ?? arc.sweepAngle,
  };
}

function modelCustomGeometryConnectionSites(shape) {
  return (shape.customConnectionSites || []).map((site) => ({
    angle: site.angleReference ?? Number(site.angle60000) / ROTATION_UNITS_PER_DEGREE,
    x: site.xReference ?? Number(site.xEmu) / EMU_PER_PIXEL,
    y: site.yReference ?? Number(site.yEmu) / EMU_PER_PIXEL,
  }));
}

function modelCustomGeometryHandleBound(source, literalField, referenceField, divisor = 1) {
  if (source[referenceField] !== undefined) return source[referenceField];
  if (source[literalField] !== undefined) return Number(source[literalField]) / divisor;
  return undefined;
}

function modelCustomGeometryAdjustmentHandles(shape) {
  return (shape.customAdjustmentHandles || []).map((entry) => {
    const handle = entry.handle?.value;
    const kind = entry.handle?.case;
    if (!handle || (kind !== "xy" && kind !== "polar")) {
      throw new OfficeKitCodecError("OfficeKit Codec returned an invalid custom geometry adjustment handle.", [], { code: "invalid_presentation_geometry" });
    }
    const modeled = {
      kind,
      x: handle.position?.xReference ?? Number(handle.position?.x || 0n) / EMU_PER_PIXEL,
      y: handle.position?.yReference ?? Number(handle.position?.y || 0n) / EMU_PER_PIXEL,
    };
    if (kind === "xy") {
      if (handle.xAdjustment) modeled.xAdjustment = handle.xAdjustment;
      if (handle.yAdjustment) modeled.yAdjustment = handle.yAdjustment;
      for (const [field, literalField, referenceField] of [
        ["minX", "minX", "minXReference"], ["maxX", "maxX", "maxXReference"],
        ["minY", "minY", "minYReference"], ["maxY", "maxY", "maxYReference"],
      ]) {
        const value = modelCustomGeometryHandleBound(handle, literalField, referenceField);
        if (value !== undefined) modeled[field] = value;
      }
    } else {
      if (handle.radialAdjustment) modeled.radialAdjustment = handle.radialAdjustment;
      if (handle.angleAdjustment) modeled.angleAdjustment = handle.angleAdjustment;
      for (const [field, literalField, referenceField, divisor] of [
        ["minRadius", "minRadius", "minRadiusReference", 1], ["maxRadius", "maxRadius", "maxRadiusReference", 1],
        ["minAngle", "minAngle60000", "minAngleReference", ROTATION_UNITS_PER_DEGREE],
        ["maxAngle", "maxAngle60000", "maxAngleReference", ROTATION_UNITS_PER_DEGREE],
      ]) {
        const value = modelCustomGeometryHandleBound(handle, literalField, referenceField, divisor);
        if (value !== undefined) modeled[field] = value;
      }
    }
    return modeled;
  });
}

function modelCustomGeometryPaths(shape) {
  return (shape.customPaths || []).map((path, pathIndex) => {
    const modeled = {
      commands: path.commands.map((command) => {
        if (command.command.case === "moveTo") return { moveTo: modelCustomGeometryPoint(command.command.value) };
        if (command.command.case === "lineTo") return { lineTo: modelCustomGeometryPoint(command.command.value) };
        if (command.command.case === "quadraticBezierTo") return {
          quadraticBezTo: {
            x1: modelCustomGeometryPoint(command.command.value.control).x,
            y1: modelCustomGeometryPoint(command.command.value.control).y,
            x: modelCustomGeometryPoint(command.command.value.end).x,
            y: modelCustomGeometryPoint(command.command.value.end).y,
          },
        };
        if (command.command.case === "arcTo") return {
          arcTo: modelCustomGeometryArc(command.command.value),
        };
        if (command.command.case === "cubicBezierTo") return {
          cubicBezTo: {
            x1: modelCustomGeometryPoint(command.command.value.control1).x,
            y1: modelCustomGeometryPoint(command.command.value.control1).y,
            x2: modelCustomGeometryPoint(command.command.value.control2).x,
            y2: modelCustomGeometryPoint(command.command.value.control2).y,
            x: modelCustomGeometryPoint(command.command.value.end).x,
            y: modelCustomGeometryPoint(command.command.value.end).y,
          },
        };
        return { close: {} };
      }),
    };
    if (Number(path.width) > 0) modeled.width = Number(path.width);
    if (Number(path.height) > 0) modeled.height = Number(path.height);
    if (path.fillMode === PresentationCustomGeometryPath_FillMode.NORMAL) modeled.fillMode = "normal";
    else if (path.fillMode === PresentationCustomGeometryPath_FillMode.NONE) modeled.fillMode = "none";
    else if (path.fillMode !== PresentationCustomGeometryPath_FillMode.UNSPECIFIED) {
      throw new OfficeKitCodecError(`Presentation custom path ${pathIndex + 1} uses an unsupported fill mode.`, [], { code: "unsupported_presentation_features" });
    }
    if (path.stroke !== undefined) modeled.stroke = path.stroke;
    if (path.extrusionAllowed !== undefined) modeled.extrusionAllowed = path.extrusionAllowed;
    return modeled;
  });
}

function modelCustomGeometryTextRectangle(shape) {
  const rectangle = shape.textRectangle;
  if (!rectangle) return undefined;
  return Object.fromEntries(CUSTOM_TEXT_RECTANGLE_FIELDS.map(([field, literalField, referenceField]) => [
    field,
    rectangle[referenceField] ?? Number(rectangle[literalField]) / EMU_PER_PIXEL,
  ]));
}

function modelPlaceholderTransform(frame) {
  return modelPresentationTransform(frame);
}

function slidePlaceholderSnapshot(shape) {
  return JSON.stringify(shape.layoutJson());
}

function slidePlaceholderTextStructureSnapshot(shape) {
  const paragraphs = clonedPresentationValue(shape.text?.paragraphs || []);
  for (const paragraph of paragraphs) {
    for (const run of paragraph.runs || []) {
      if (Object.hasOwn(run, "text")) run.text = "";
      if (run.field) run.field.text = "";
    }
  }
  return JSON.stringify(paragraphs);
}

function slidePlaceholderReadOnlySnapshot(shape) {
  const snapshot = clonedPresentationValue(shape.layoutJson());
  delete snapshot.text;
  delete snapshot.paragraphs;
  snapshot.inheritedParagraphStyles = clonedPresentationValue(shape.text?.inheritedParagraphStyles || {});
  return JSON.stringify(snapshot);
}

function slidePlaceholderState(shape) {
  return {
    full: slidePlaceholderSnapshot(shape),
    readOnly: slidePlaceholderReadOnlySnapshot(shape),
    textStructure: slidePlaceholderTextStructureSnapshot(shape),
  };
}

function isPlainPresentationTextRequest(shape) {
  return JSON.stringify(shape.text?.paragraphs || []) === JSON.stringify(normalizePresentationParagraphs(shape.text?.value || ""));
}

function sourceBoundSlidePlaceholderTextBody(shape, originalShape, originalState, assetCatalog, customShowLinks) {
  if (slidePlaceholderTextStructureSnapshot(shape) === originalState.textStructure) {
    return presentationTextBody(shape, originalShape, assetCatalog, customShowLinks);
  }

  if (!isPlainPresentationTextRequest(shape)) {
    throw new OfficeKitCodecError(
      `Presentation slide placeholder ${shape.id} changed its source-bound paragraph, inline, or formatting topology. Use text.replace(...) for structured text, or text.set(...) with the source line-break topology intact.`,
      [],
      { code: "presentation_text_topology_changed" },
    );
  }

  // TextFrame.set(...) intentionally presents a plain request. Preserve the
  // imported run/paragraph formatting and map each newline-delimited segment
  // back to exactly one original text run. Ambiguous multi-run spans fail
  // closed; callers can use text.replace(...) for a precise structured edit.
  const textBody = clonedPresentationValue(originalShape.textBody);
  const spans = [[]];
  for (let paragraphIndex = 0; paragraphIndex < (textBody?.paragraphs || []).length; paragraphIndex += 1) {
    const paragraph = textBody.paragraphs[paragraphIndex];
    for (const run of paragraph.runs || []) {
      if (run.content?.case === "text") spans.at(-1).push(run);
      else if (run.content?.case === "lineBreak") spans.push([]);
      else {
        throw new OfficeKitCodecError(
          `Presentation slide placeholder ${shape.id} contains a field or unsupported inline that cannot be replaced through text.set(...).`,
          [],
          { code: "presentation_text_topology_changed" },
        );
      }
    }
    if (paragraphIndex + 1 < textBody.paragraphs.length) spans.push([]);
  }
  const segments = shape.text.value.split("\n");
  if (segments.length !== spans.length || spans.some((runs) => runs.length !== 1)) {
    throw new OfficeKitCodecError(
      `Presentation slide placeholder ${shape.id} cannot map text.set(...) onto its source-bound line-break and styled-run topology. Preserve the newline count or use text.replace(...).`,
      [],
      { code: "presentation_text_topology_changed" },
    );
  }
  for (let index = 0; index < segments.length; index += 1) {
    spans[index][0].content = { case: "text", value: segments[index] };
  }
  return textBody;
}

function presentationSlidePlaceholder(shape, original, originalState, assetCatalog, customShowLinks) {
  const currentState = slidePlaceholderState(shape);
  if (currentState.full === originalState.full) return original;
  if (original?.source?.textEditable !== true) {
    throw new OfficeKitCodecError(
      `Presentation slide placeholder ${shape.id} is source-bound and has no safely editable owner-local text graph.`,
      [],
      { code: "unsupported_presentation_edit" },
    );
  }
  if (currentState.readOnly !== originalState.readOnly) {
    throw new OfficeKitCodecError(
      `Presentation slide placeholder ${shape.id} may edit only its owner-local text; identity, geometry, formatting, and shape semantics remain source-bound.`,
      [],
      { code: "unsupported_presentation_edit" },
    );
  }
  const requested = clonedPresentationValue(original);
  const originalShape = original.content.value;
  requested.content.value.text = shape.text.value;
  // `requested` is a cloned protobuf message. Replacing one of its message
  // fields with a plain initializer does not make Buf recursively create that
  // child again, so canonicalize the rebuilt text body before binary proof.
  requested.content.value.textBody = create(
    PresentationTextBodySchema,
    sourceBoundSlidePlaceholderTextBody(shape, originalShape, originalState, assetCatalog, customShowLinks),
  );
  return requested;
}

function modelPresentationShapeLine(shape) {
  return {
    fill: shape.lineRgb ? `#${shape.lineRgb}` : "transparent",
    width: Number(shape.lineWidthEmu) / EMU_PER_POINT,
    style: shape.lineStyle || (shape.lineRgb ? "solid" : "none"),
    ...(shape.startArrow ? { head: {
      type: shape.startArrow,
      ...(shape.startArrowWidth ? { width: shape.startArrowWidth } : {}),
      ...(shape.startArrowLength ? { length: shape.startArrowLength } : {}),
    } } : {}),
    ...(shape.endArrow ? { tail: {
      type: shape.endArrow,
      ...(shape.endArrowWidth ? { width: shape.endArrowWidth } : {}),
      ...(shape.endArrowLength ? { length: shape.endArrowLength } : {}),
    } } : {}),
    ...(shape.lineCap ? { cap: shape.lineCap } : {}),
    ...(shape.lineJoin ? { join: shape.lineJoin } : {}),
  };
}

function modelPresentationOpaqueElement(element, assetCatalog, nativeGraph, sourcePart) {
  const opaque = element.content.value;
  return {
    kind: "nativeObject",
    id: element.id,
    name: element.name,
    _officeKitSharePartBytes: true,
    nativeKind: opaque.nativeKind || presentationNativeKind(opaque.elementName),
    text: opaque.text,
    position: {
      left: Number(opaque.leftEmu) / EMU_PER_PIXEL,
      top: Number(opaque.topEmu) / EMU_PER_PIXEL,
      width: Number(opaque.widthEmu) / EMU_PER_PIXEL,
      height: Number(opaque.heightEmu) / EMU_PER_PIXEL,
    },
    rawXml: opaque.rawXml,
    sourcePart,
    editable: false,
    placementCapability: {
      sourceBound: Boolean(element.source),
      known: true,
      supported: element.source?.editable === true && !["oleObject", "diagram"].includes(opaque.nativeKind),
      blockedReason: element.source?.editable === true && opaque.nativeKind === "oleObject"
        ? "embedded Office payload is editable only through its bounded replacement API"
        : element.source?.editable === true && opaque.nativeKind === "diagram"
          ? "diagram text is editable only through its bounded diagram-text API"
          : element.source?.editable === true ? "" : "opaque native frame is not proven safe to edit",
    },
    ...(opaque.oleWorkbook ? { oleWorkbook: {
      partPath: opaque.oleWorkbook.partPath,
      contentType: opaque.oleWorkbook.contentType,
      sourceSha256: opaque.oleWorkbook.sourceSha256,
      relationshipId: opaque.oleWorkbook.relationshipId,
    } } : {}),
    ...(opaque.oleOfficePackage ? { oleOfficePackage: {
      partPath: opaque.oleOfficePackage.partPath,
      contentType: opaque.oleOfficePackage.contentType,
      sourceSha256: opaque.oleOfficePackage.sourceSha256,
      relationshipId: opaque.oleOfficePackage.relationshipId,
      kind: opaque.oleOfficePackage.kind,
    } } : {}),
    ...(opaque.diagramText ? { diagramText: {
      partPath: opaque.diagramText.partPath,
      contentType: opaque.diagramText.contentType,
      sourceSha256: opaque.diagramText.sourceSha256,
      relationshipId: opaque.diagramText.relationshipId,
      nodes: (opaque.diagramText.nodes || []).map((node) => ({
        id: node.modelId,
        text: node.text,
        runs: node.runTexts?.length ? [...node.runTexts] : [node.text],
      })),
    } } : {}),
    ...(opaque.nativeChart ? { nativeChart: {
      partPath: opaque.nativeChart.partPath,
      contentType: opaque.nativeChart.contentType,
      sourceSha256: opaque.nativeChart.sourceSha256,
      relationshipId: opaque.nativeChart.relationshipId,
      titleLeaves: (opaque.nativeChart.titleLeaves || []).map((leaf) => ({
        textLeafIndex: leaf.textLeafIndex,
        text: leaf.text,
      })),
      embeddedPackagePartPath: opaque.nativeChart.embeddedPackagePartPath,
      embeddedPackageSourceSha256: opaque.nativeChart.embeddedPackageSourceSha256,
      embeddedPackageRelationshipId: opaque.nativeChart.embeddedPackageRelationshipId,
      dataPoints: (opaque.nativeChart.dataPoints || []).map((point) => ({
        seriesIndex: point.seriesIndex,
        pointIndex: point.pointIndex,
        value: point.value,
        formula: point.formula,
        worksheetPartPath: point.worksheetPartPath,
        worksheetSourceSha256: point.worksheetSourceSha256,
        worksheetName: point.worksheetName,
        cellReference: point.cellReference,
      })),
    } } : {}),
    ...nativeGraph(opaque, sourcePart),
  };
}

function modelPresentationGroupChild(element, assetCatalog, customShowLinks, nativeGraph, sourcePart) {
  const common = { id: element.id, name: element.name };
  if (element.content.case === "shape") {
    const shape = element.content.value;
    if (shape.placeholder) throw new OfficeKitCodecError(`Presentation group ${element.id} contains an unsupported placeholder child.`, [], { code: "invalid_presentation_group" });
    return {
      kind: "shape",
      ...common,
      geometry: shape.geometry || "rect",
      ...(shape.customAdjustments?.length ? { customAdjustments: modelCustomGeometryGuides(shape.customAdjustments) } : {}),
      ...(shape.customGuides?.length ? { customGuides: modelCustomGeometryGuides(shape.customGuides) } : {}),
      ...(shape.customConnectionSites?.length ? { customConnectionSites: modelCustomGeometryConnectionSites(shape) } : {}),
      ...(shape.customAdjustmentHandles?.length ? { customAdjustmentHandles: modelCustomGeometryAdjustmentHandles(shape) } : {}),
      ...(shape.customPaths?.length ? { customPaths: modelCustomGeometryPaths(shape) } : {}),
      ...(shape.textRectangle ? { textRectangle: modelCustomGeometryTextRectangle(shape) } : {}),
      position: {
        left: Number(shape.leftEmu) / EMU_PER_PIXEL,
        top: Number(shape.topEmu) / EMU_PER_PIXEL,
        width: Number(shape.widthEmu) / EMU_PER_PIXEL,
        height: Number(shape.heightEmu) / EMU_PER_PIXEL,
      },
      ...(shape.transform ? { transform: modelPresentationTransform(shape.transform) } : {}),
      fill: modelPresentationShapeFill(shape),
      line: modelPresentationShapeLine(shape),
      ...(shape.shadow ? { shadow: modelPresentationShadow(shape.shadow) } : {}),
      ...(shape.useBackgroundFill === undefined ? {} : { _officeKitUseBackgroundFill: shape.useBackgroundFill }),
      ...modelPresentationAccessibility(shape.accessibility),
      _officeKitAccessibilityEditable: element.source?.accessibilityEditable === true,
      text: modelText(shape, assetCatalog, customShowLinks),
      textBodyProperties: modelTextBodyProperties(shape),
    };
  }
  if (element.content.case === "image") {
    const image = element.content.value;
    return {
      kind: "image",
      ...common,
      position: {
        left: Number(image.leftEmu) / EMU_PER_PIXEL,
        top: Number(image.topEmu) / EMU_PER_PIXEL,
        width: Number(image.widthEmu) / EMU_PER_PIXEL,
        height: Number(image.heightEmu) / EMU_PER_PIXEL,
      },
      ...modelPresentationImageAccessibility(image),
      _officeKitAccessibilityEditable: element.source?.accessibilityEditable === true,
      contentType: assetCatalog.contentType(image.assetId),
      _officeKitDataUrlSource: assetCatalog.dataUrlSource(image.assetId),
      ...(image.svgAssetId ? { _officeKitSvgDataUrlSource: assetCatalog.dataUrlSource(image.svgAssetId) } : {}),
      fit: "stretch",
      ...(image.crop ? { crop: presentationImageCropFromWire(image.crop) } : {}),
      geometry: "rect",
      ...(image.transform ? { transform: modelPresentationTransform(image.transform) } : {}),
    };
  }
  if (element.content.case === "table") {
    const table = element.content.value;
    return {
      kind: "table",
      ...common,
      position: {
        left: Number(table.leftEmu) / EMU_PER_PIXEL,
        top: Number(table.topEmu) / EMU_PER_PIXEL,
        width: Number(table.widthEmu) / EMU_PER_PIXEL,
        height: Number(table.heightEmu) / EMU_PER_PIXEL,
      },
      values: table.rows.map((row) => row.cells.map((cell) => cell.text)),
      rows: table.rows.length,
      columns: table.columnWidthsEmu.length,
      styleOptions: { headerRow: table.firstRow === true, bandedRows: table.bandedRows === true },
      ...modelPresentationAccessibility(table.accessibility, "Imported Presentation table"),
      _officeKitAccessibilityEditable: element.source?.accessibilityEditable === true,
    };
  }
  if (element.content.case === "connector") {
    const connector = element.content.value;
    return {
      kind: "connector",
      ...common,
      connectorType: connector.connectorType || "straight",
      start: { x: Number(connector.startXEmu) / EMU_PER_PIXEL, y: Number(connector.startYEmu) / EMU_PER_PIXEL },
      end: { x: Number(connector.endXEmu) / EMU_PER_PIXEL, y: Number(connector.endYEmu) / EMU_PER_PIXEL },
      startTargetId: connector.startTargetId || undefined,
      endTargetId: connector.endTargetId || undefined,
      startSiteIndex: Number(connector.startConnectionSiteIndex || 0),
      endSiteIndex: Number(connector.endConnectionSiteIndex || 0),
      line: {
        fill: connector.lineRgb ? `#${connector.lineRgb}` : "transparent",
        width: Number(connector.lineWidthEmu) / EMU_PER_POINT,
        style: connector.lineStyle || "solid",
        ...(connector.startArrow ? { startArrow: connector.startArrow } : {}),
        ...(connector.endArrow ? { endArrow: connector.endArrow } : {}),
      },
      ...(connector.startArrow ? { head: { type: connector.startArrow, ...(connector.startArrowWidth ? { width: connector.startArrowWidth } : {}), ...(connector.startArrowLength ? { length: connector.startArrowLength } : {}) } } : {}),
      ...(connector.endArrow ? { tail: { type: connector.endArrow, ...(connector.endArrowWidth ? { width: connector.endArrowWidth } : {}), ...(connector.endArrowLength ? { length: connector.endArrowLength } : {}) } } : {}),
      ...(connector.lineCap ? { cap: connector.lineCap } : {}),
      ...(connector.lineJoin ? { join: connector.lineJoin } : {}),
      ...modelPresentationAccessibility(connector.accessibility, "Imported Presentation connector"),
      _officeKitAccessibilityEditable: element.source?.accessibilityEditable === true,
      _officeKitSourceBound: Boolean(element.source),
    };
  }
  if (element.content.case === "chart") return { kind: "chart", ...common, ...modelPresentationChart(element.content.value, element.source?.accessibilityEditable) };
  if (element.content.case === "group") return { kind: "groupShape", ...modelPresentationGroup(element, assetCatalog, customShowLinks, nativeGraph, sourcePart) };
  if (element.content.case === "opaque") return modelPresentationOpaqueElement(element, assetCatalog, nativeGraph, sourcePart);
  throw new OfficeKitCodecError(`Presentation group child ${element.id} has unsupported wire content ${element.content.case || "none"}.`, [], { code: "invalid_presentation_group" });
}

function modelPresentationGroup(element, assetCatalog, customShowLinks, nativeGraph, sourcePart) {
  const group = element.content.value;
  return {
    id: element.id,
    name: element.name,
    position: {
      left: Number(group.leftEmu) / EMU_PER_PIXEL,
      top: Number(group.topEmu) / EMU_PER_PIXEL,
      width: Number(group.widthEmu) / EMU_PER_PIXEL,
      height: Number(group.heightEmu) / EMU_PER_PIXEL,
    },
    childFrame: {
      left: Number(group.childLeftEmu) / EMU_PER_PIXEL,
      top: Number(group.childTopEmu) / EMU_PER_PIXEL,
      width: Number(group.childWidthEmu) / EMU_PER_PIXEL,
      height: Number(group.childHeightEmu) / EMU_PER_PIXEL,
    },
    ...modelPresentationAccessibility(group.accessibility, "Imported Presentation group"),
    _officeKitAccessibilityEditable: element.source?.accessibilityEditable === true,
    children: group.children.map((child) => modelPresentationGroupChild(child, assetCatalog, customShowLinks, nativeGraph, sourcePart)),
  };
}

export async function presentationFromEnvelope(envelope, options = {}) {
  if (envelope.family !== ArtifactFamily.PRESENTATION || envelope.payload.case !== "presentation") {
    throw new OfficeKitCodecError("OfficeKit response does not contain a presentation artifact.", [], { code: "invalid_presentation_artifact" });
  }
  const source = envelope.payload.value;
  if (source.customShowsOpaque && source.customShows?.length) {
    throw new OfficeKitCodecError("OfficeKit returned both opaque and semantic presentation custom shows.", [], { code: "invalid_presentation_artifact" });
  }
  if (source.sectionsOpaque && source.sections?.length) {
    throw new OfficeKitCodecError("OfficeKit returned both opaque and semantic PowerPoint sections.", [], { code: "invalid_presentation_artifact" });
  }
  const customShowLinks = new Map();
  for (const show of source.customShows || []) {
    if (!show.id || customShowLinks.has(show.id)) {
      throw new OfficeKitCodecError("OfficeKit returned an invalid or duplicate presentation custom-show ID.", [], { code: "invalid_presentation_artifact" });
    }
    customShowLinks.set(show.id, show.name);
  }
  const assetCatalog = createPresentationAssetCatalog(envelope.assets || [], { shareBytes: true });
  const assetBytesBySha256 = new Map((envelope.assets || []).map((asset) => [String(asset.sha256 || "").toLowerCase(), asset.data]));
  const nativeGraph = await materializePresentationNativeGraphs(envelope, { assetBytesBySha256 });
  const importedTheme = options?.importedThemeProfile;
  const presentation = Presentation.create({
    slideSize: { width: Number(source.slideWidthEmu) / EMU_PER_PIXEL, height: Number(source.slideHeightEmu) / EMU_PER_PIXEL },
    ...(importedTheme?.kind === "theme" ? {
      theme: {
        id: importedTheme.id,
        name: importedTheme.name,
        colors: importedTheme.colors,
        fonts: importedTheme.fonts,
        colorMap: importedTheme.colorMap,
      },
    } : {}),
  });
  presentation.id = source.id || presentation.id;
  const slideGuides = modelPresentationSlideGuides(source.viewProperties);
  presentation.view._setImportedProperties(modelPresentationView(source.viewProperties));
  const masterStates = [];
  if (source.masters?.length) {
    presentation.masters.items.length = 0;
    for (const sourceMaster of source.masters) {
      const model = presentation.masters.add({
        id: sourceMaster.id,
        name: sourceMaster.name,
        ...(sourceMaster.background ? { background: modelBackground(sourceMaster.background, assetCatalog) } : {}),
        placeholders: (sourceMaster.placeholders || []).map((placeholder) => modelPlaceholder(placeholder, assetCatalog, customShowLinks)),
        textParagraphStyles: modelMasterTextStyles(sourceMaster, assetCatalog),
        slideGuides,
      });
      if (!sourceMaster.background) model.background = undefined;
      for (let index = 0; index < sourceMaster.placeholders.length; index += 1) {
        if (!sourceMaster.placeholders[index].directFrame) model.placeholders[index].position = undefined;
      }
      masterStates.push({
        wire: sourceMaster,
        model,
        snapshot: masterReadOnlySnapshot(model),
      });
    }
  }
  const layoutStates = [];
  for (const sourceLayout of source.layouts || []) {
    const model = presentation.layouts.add({
      id: sourceLayout.id,
      name: sourceLayout.name,
      type: sourceLayout.type,
      masterId: sourceLayout.masterId,
      ...(sourceLayout.background ? { background: modelBackground(sourceLayout.background, assetCatalog) } : {}),
      placeholders: (sourceLayout.placeholders || []).map((placeholder) => modelPlaceholder(placeholder, assetCatalog, customShowLinks)),
      slideGuides,
    });
    layoutStates.push({
      wire: sourceLayout,
      model,
      snapshot: layoutReadOnlySnapshot(model),
    });
  }
  const slideStates = [];
  for (const sourceSlide of source.slides) {
    const slide = presentation.slides.add({
      name: sourceSlide.name,
      ...(sourceSlide.hidden === undefined ? {} : { hidden: sourceSlide.hidden }),
      ...(sourceSlide.background ? { background: modelBackground(sourceSlide.background, assetCatalog) } : {}),
      ...(sourceSlide.transition ? { transition: modelPresentationTransition(sourceSlide.transition, slideStates.length) } : {}),
      ...(sourceSlide.animations?.length ? { animations: sourceSlide.animations.map((animation) => modelPresentationAnimation(animation, `slide ${slideStates.length + 1}`)) } : {}),
      ...(sourceSlide.morph ? { morph: modelPresentationMorph(sourceSlide.morph, `slide ${slideStates.length + 1}`) } : {}),
    });
    Object.defineProperty(slide, PRESENTATION_SLIDE_VISIBILITY_CAPABILITY, {
      value: Object.freeze({
        sourceBound: true,
        known: sourceSlide.hidden !== undefined,
        editable: sourceSlide.source?.visibilityEditable === true,
      }),
    });
    const deletionCapability = sourceSlide.source?.deletionCapability;
    Object.defineProperty(slide, PRESENTATION_SLIDE_DELETION_CAPABILITY, {
      value: Object.freeze({
        sourceBound: true,
        known: Boolean(deletionCapability),
        supported: deletionCapability?.supported === true,
        blockedReason: deletionCapability?.blockedReason || "",
        ownedPartCount: Number(deletionCapability?.ownedPartCount || 0),
      }),
    });
    const cloneCapability = sourceSlide.source?.cloneCapability;
    const sourceRevisionSha256 = String(envelope.source?.packageSha256 || envelope.opaqueOpc?.sourcePackage?.sha256 || "").toLowerCase();
    Object.defineProperty(slide, PRESENTATION_SLIDE_CLONE_CAPABILITY, {
      value: Object.freeze({
        sourceBound: true,
        known: Boolean(cloneCapability),
        supported: cloneCapability?.supported === true,
        blockedReason: cloneCapability?.blockedReason || "",
        clonedPartCount: Number(cloneCapability?.clonedPartCount || 0),
        sharedPartCount: Number(cloneCapability?.sharedPartCount || 0),
        ...(sourceRevisionSha256 ? { sourceRevisionSha256 } : {}),
      }),
    });
    Object.defineProperty(slide, PRESENTATION_SLIDE_CONTINUATION_CAPABILITY, {
      value: Object.freeze({
        sourceBound: true,
        ready: true,
        profile: "bounded-overlay",
        requiresExportReopen: false,
        oneSlideMutationPerExport: true,
        shapeGeometries: Object.freeze(["textbox", "rect", "roundRect", "ellipse"]),
        embeddedImage: true,
        ...(sourceRevisionSha256 ? { sourceRevisionSha256 } : {}),
      }),
    });
    slide.id = sourceSlide.id || slide.id;
    slide.layoutId = sourceSlide.layoutId || undefined;
    slide.addNotes(sourceSlide.speakerNotes?.textBody
      ? modelText(sourceSlide.speakerNotes, assetCatalog, customShowLinks)
      : sourceSlide.speakerNotes?.text || "");
    Object.defineProperty(slide.speakerNotes, PRESENTATION_SPEAKER_NOTES_CAPABILITY, {
      value: Object.freeze({
        sourceBound: true,
        partPresent: Boolean(sourceSlide.speakerNotes),
        editable: Boolean(sourceSlide.speakerNotes?.source?.editable),
        addable: Boolean(!sourceSlide.speakerNotes && sourceSlide.source?.speakerNotesAddable),
      }),
    });
    Object.defineProperty(slide.transition, PRESENTATION_TRANSITION_CAPABILITY, {
      value: Object.freeze({
        sourceBound: true,
        partPresent: Boolean(sourceSlide.source?.transitionPresent),
        editable: Boolean(sourceSlide.source?.transitionEditable),
        addable: Boolean(sourceSlide.source?.transitionAddable),
      }),
    });
    Object.defineProperty(slide.animations, PRESENTATION_ANIMATIONS_CAPABILITY, {
      value: Object.freeze({
        sourceBound: true,
        present: Boolean(sourceSlide.source?.timingPresent),
        editable: Boolean(sourceSlide.source?.timingEditable),
        addable: Boolean(sourceSlide.source?.timingAddable),
        ...(sourceRevisionSha256 ? { sourceRevisionSha256 } : {}),
      }),
    });
    Object.defineProperty(slide.morph, PRESENTATION_MORPH_CAPABILITY, {
      value: Object.freeze({
        sourceBound: true,
        editable: Boolean(sourceSlide.source?.timingEditable && sourceSlide.morph),
        addable: Boolean(sourceSlide.source?.timingAddable && !sourceSlide.morph),
        ...(sourceRevisionSha256 ? { sourceRevisionSha256 } : {}),
      }),
    });
    const sourcePart = sourceSlide.source?.partPath;
    const entries = [];
    for (const element of sourceSlide.elements) {
      let model;
      if (element.content.case === "shape") {
        const shape = element.content.value;
        const placeholderIdentity = shape.placeholder;
        const layout = placeholderIdentity ? presentation.layouts.getItem(slide.layoutId) : undefined;
        // PowerPoint resolves slide placeholders against their linked layout by
        // idx. Type remains descriptive and deliberately does not participate
        // in this lookup.
        const inheritedPlaceholder = placeholderIdentity?.inheritsGeometry
          ? layout?.effectivePlaceholders().find((candidate) => candidate.idx === Number(placeholderIdentity.index))
          : undefined;
        const directFrame = shape.directFrame ? modelPlaceholderFrame(shape.directFrame) : undefined;
        const directTransform = shape.directFrame ? modelPlaceholderTransform(shape.directFrame) : undefined;
        const effectiveFrame = directFrame || inheritedPlaceholder?.position || {
          left: Number(shape.leftEmu) / EMU_PER_PIXEL,
          top: Number(shape.topEmu) / EMU_PER_PIXEL,
          width: Number(shape.widthEmu) / EMU_PER_PIXEL,
          height: Number(shape.heightEmu) / EMU_PER_PIXEL,
        };
        const geometrySource = directFrame
          ? "slide"
          : placeholderIdentity?.inheritsGeometry
            ? (inheritedPlaceholder?.geometrySource || "unresolved")
            : placeholderIdentity
              ? "slide-unrecognized"
              : undefined;
        const effectiveTransform = directFrame
          ? directTransform
          : placeholderIdentity
            ? inheritedPlaceholder?.transform
            : modelPresentationTransform(shape.transform);
        model = slide.shapes.add({
          id: element.id,
          name: element.name || inheritedPlaceholder?.name,
          geometry: shape.geometry || "rect",
          ...(shape.customAdjustments?.length ? { customAdjustments: modelCustomGeometryGuides(shape.customAdjustments) } : {}),
          ...(shape.customGuides?.length ? { customGuides: modelCustomGeometryGuides(shape.customGuides) } : {}),
          ...(shape.customConnectionSites?.length ? { customConnectionSites: modelCustomGeometryConnectionSites(shape) } : {}),
          ...(shape.customAdjustmentHandles?.length ? { customAdjustmentHandles: modelCustomGeometryAdjustmentHandles(shape) } : {}),
          ...(shape.customPaths?.length ? { customPaths: modelCustomGeometryPaths(shape) } : {}),
          ...(shape.textRectangle ? { textRectangle: modelCustomGeometryTextRectangle(shape) } : {}),
          position: { ...effectiveFrame },
          ...(effectiveTransform && Object.keys(effectiveTransform).length ? { transform: effectiveTransform } : {}),
          ...(placeholderIdentity ? { placeholder: {
            layoutId: slide.layoutId,
            type: placeholderIdentity.type,
            idx: Number(placeholderIdentity.index),
            geometrySource,
            textEditable: element.source?.textEditable === true,
          } } : {}),
          fill: modelPresentationShapeFill(shape),
          line: modelPresentationShapeLine(shape),
          ...(shape.shadow ? { shadow: modelPresentationShadow(shape.shadow) } : {}),
          ...(shape.useBackgroundFill === undefined ? {} : { _officeKitUseBackgroundFill: shape.useBackgroundFill }),
          ...modelPresentationAccessibility(shape.accessibility),
          _officeKitAccessibilityEditable: element.source?.accessibilityEditable === true,
          text: modelText(shape, assetCatalog, customShowLinks),
          textBodyProperties: modelTextBodyProperties(shape),
        });
        model.text.inheritedParagraphStyles = modelListStyles(shape, assetCatalog);
      } else if (element.content.case === "image") {
        const image = element.content.value;
        model = slide.images.add({
          id: element.id,
          name: element.name,
          position: {
            left: Number(image.leftEmu) / EMU_PER_PIXEL,
            top: Number(image.topEmu) / EMU_PER_PIXEL,
            width: Number(image.widthEmu) / EMU_PER_PIXEL,
            height: Number(image.heightEmu) / EMU_PER_PIXEL,
          },
          ...modelPresentationImageAccessibility(image),
          _officeKitAccessibilityEditable: element.source?.accessibilityEditable === true,
          contentType: assetCatalog.contentType(image.assetId),
          _officeKitDataUrlSource: assetCatalog.dataUrlSource(image.assetId),
          ...(image.svgAssetId ? { _officeKitSvgDataUrlSource: assetCatalog.dataUrlSource(image.svgAssetId) } : {}),
          fit: "stretch",
          ...(image.crop ? { crop: presentationImageCropFromWire(image.crop) } : {}),
          geometry: "rect",
          ...(image.transform ? { transform: modelPresentationTransform(image.transform) } : {}),
        });
      } else if (element.content.case === "table") {
        const table = element.content.value;
        model = slide.tables.add({
          id: element.id,
          name: element.name,
          position: {
            left: Number(table.leftEmu) / EMU_PER_PIXEL,
            top: Number(table.topEmu) / EMU_PER_PIXEL,
            width: Number(table.widthEmu) / EMU_PER_PIXEL,
            height: Number(table.heightEmu) / EMU_PER_PIXEL,
          },
          values: table.rows.map((row) => row.cells.map((cell) => cell.text)),
          rows: table.rows.length,
          columns: table.columnWidthsEmu.length,
          styleOptions: {
            headerRow: table.firstRow === true,
            bandedRows: table.bandedRows === true,
          },
          ...modelPresentationAccessibility(table.accessibility, "Imported Presentation table"),
          _officeKitAccessibilityEditable: element.source?.accessibilityEditable === true,
          mergeRanges: table.mergeRanges.map((range) => ({
            startRow: Number(range.startRow),
            endRow: Number(range.endRow),
            startColumn: Number(range.startColumn),
            endColumn: Number(range.endColumn),
          })),
        });
      } else if (element.content.case === "connector") {
        const connector = element.content.value;
        model = slide.connectors.add({
          id: element.id,
          name: element.name,
          connectorType: connector.connectorType || "straight",
          start: { x: Number(connector.startXEmu) / EMU_PER_PIXEL, y: Number(connector.startYEmu) / EMU_PER_PIXEL },
          end: { x: Number(connector.endXEmu) / EMU_PER_PIXEL, y: Number(connector.endYEmu) / EMU_PER_PIXEL },
          startTargetId: connector.startTargetId || undefined,
          endTargetId: connector.endTargetId || undefined,
          startSiteIndex: Number(connector.startConnectionSiteIndex || 0),
          endSiteIndex: Number(connector.endConnectionSiteIndex || 0),
          line: {
            fill: connector.lineRgb ? `#${connector.lineRgb}` : "transparent",
            width: Number(connector.lineWidthEmu) / EMU_PER_POINT,
            style: connector.lineStyle || "solid",
            ...(connector.startArrow ? { startArrow: connector.startArrow } : {}),
            ...(connector.endArrow ? { endArrow: connector.endArrow } : {}),
          },
          ...(connector.startArrow ? { head: { type: connector.startArrow, ...(connector.startArrowWidth ? { width: connector.startArrowWidth } : {}), ...(connector.startArrowLength ? { length: connector.startArrowLength } : {}) } } : {}),
          ...(connector.endArrow ? { tail: { type: connector.endArrow, ...(connector.endArrowWidth ? { width: connector.endArrowWidth } : {}), ...(connector.endArrowLength ? { length: connector.endArrowLength } : {}) } } : {}),
          ...(connector.lineCap ? { cap: connector.lineCap } : {}),
          ...(connector.lineJoin ? { join: connector.lineJoin } : {}),
          ...modelPresentationAccessibility(connector.accessibility, "Imported Presentation connector"),
          _officeKitAccessibilityEditable: element.source?.accessibilityEditable === true,
          _officeKitSourceBound: Boolean(element.source),
        });
      } else if (element.content.case === "chart") {
        const chart = modelPresentationChart(element.content.value, element.source?.accessibilityEditable);
        model = slide.charts.add(chart.chartType, {
          id: element.id,
          name: element.name,
          ...chart,
        });
      } else if (element.content.case === "group") {
        model = slide.groups.add(modelPresentationGroup(element, assetCatalog, customShowLinks, nativeGraph, sourcePart));
        markPresentationImportedGroupSnapshots(model, element, sourceRevisionSha256);
      } else if (element.content.case === "opaque") {
        const opaque = element.content.value;
        model = slide.nativeObjects.add({
          id: element.id,
          name: element.name,
          _officeKitSharePartBytes: true,
          nativeKind: opaque.nativeKind || presentationNativeKind(opaque.elementName),
          text: opaque.text,
          position: {
            left: Number(opaque.leftEmu) / EMU_PER_PIXEL,
            top: Number(opaque.topEmu) / EMU_PER_PIXEL,
            width: Number(opaque.widthEmu) / EMU_PER_PIXEL,
            height: Number(opaque.heightEmu) / EMU_PER_PIXEL,
          },
          rawXml: opaque.rawXml,
          sourcePart,
          editable: false,
        placementCapability: {
          sourceBound: Boolean(element.source),
          known: true,
          supported: element.source?.editable === true && !["oleObject", "diagram"].includes(opaque.nativeKind),
          blockedReason: element.source?.editable === true && opaque.nativeKind === "oleObject"
            ? "embedded Office payload is editable only through its bounded replacement API"
            : element.source?.editable === true && opaque.nativeKind === "diagram"
              ? "diagram text is editable only through its bounded diagram-text API"
              : element.source?.editable === true ? "" : "opaque native frame is not proven safe to edit",
        },
          ...(opaque.oleWorkbook ? { oleWorkbook: {
            partPath: opaque.oleWorkbook.partPath,
            contentType: opaque.oleWorkbook.contentType,
            sourceSha256: opaque.oleWorkbook.sourceSha256,
            relationshipId: opaque.oleWorkbook.relationshipId,
          } } : {}),
          ...(opaque.oleOfficePackage ? { oleOfficePackage: {
            partPath: opaque.oleOfficePackage.partPath,
            contentType: opaque.oleOfficePackage.contentType,
            sourceSha256: opaque.oleOfficePackage.sourceSha256,
            relationshipId: opaque.oleOfficePackage.relationshipId,
            kind: opaque.oleOfficePackage.kind,
          } } : {}),
          ...(opaque.diagramText ? { diagramText: {
            partPath: opaque.diagramText.partPath,
            contentType: opaque.diagramText.contentType,
            sourceSha256: opaque.diagramText.sourceSha256,
            relationshipId: opaque.diagramText.relationshipId,
            nodes: (opaque.diagramText.nodes || []).map((node) => ({
              id: node.modelId,
              text: node.text,
              runs: node.runTexts?.length ? [...node.runTexts] : [node.text],
            })),
          } } : {}),
          ...(opaque.nativeChart ? { nativeChart: {
            partPath: opaque.nativeChart.partPath,
            contentType: opaque.nativeChart.contentType,
            sourceSha256: opaque.nativeChart.sourceSha256,
            relationshipId: opaque.nativeChart.relationshipId,
            titleLeaves: (opaque.nativeChart.titleLeaves || []).map((leaf) => ({
              textLeafIndex: leaf.textLeafIndex,
              text: leaf.text,
            })),
            embeddedPackagePartPath: opaque.nativeChart.embeddedPackagePartPath,
            embeddedPackageSourceSha256: opaque.nativeChart.embeddedPackageSourceSha256,
            embeddedPackageRelationshipId: opaque.nativeChart.embeddedPackageRelationshipId,
            dataPoints: (opaque.nativeChart.dataPoints || []).map((point) => ({
              seriesIndex: point.seriesIndex,
              pointIndex: point.pointIndex,
              value: point.value,
              formula: point.formula,
              worksheetPartPath: point.worksheetPartPath,
              worksheetSourceSha256: point.worksheetSourceSha256,
              worksheetName: point.worksheetName,
              cellReference: point.cellReference,
            })),
          } } : {}),
          ...nativeGraph(opaque, sourcePart),
        });
      } else {
        throw new OfficeKitCodecError(`Presentation element ${element.id} has no supported wire content.`, [], { code: "invalid_presentation_artifact" });
      }
      const elementDeletionCapability = element.source?.deletionCapability;
      const elementZOrderCapability = element.source?.zOrderCapability;
      const deletionNativeId = Number(elementDeletionCapability?.nativeId || 0) || undefined;
      if (model.nativeId === undefined && deletionNativeId !== undefined) model.nativeId = deletionNativeId;
      Object.defineProperty(model, PRESENTATION_ELEMENT_ORDER_CAPABILITY, {
        value: Object.freeze({
          sourceBound: true,
          known: Boolean(elementZOrderCapability),
          editable: elementZOrderCapability?.supported === true,
          blockedReason: elementZOrderCapability
            ? elementZOrderCapability.blockedReason || ""
            : "Imported direct-element order capability is unavailable.",
          ...(sourceRevisionSha256 ? { sourceRevisionSha256 } : {}),
        }),
      });
      Object.defineProperty(model, PRESENTATION_ELEMENT_DELETION_CAPABILITY, {
        value: Object.freeze({
          sourceBound: true,
          known: Boolean(elementDeletionCapability),
          supported: elementDeletionCapability?.supported === true,
          blockedReason: elementDeletionCapability?.blockedReason || "",
          nativeId: deletionNativeId,
        }),
      });
      entries.push({
        wire: element,
        model,
        placeholderSnapshot: element.content.case === "shape" && element.content.value.placeholder
          ? slidePlaceholderState(model)
          : undefined,
        snapshot: element.content.case === "opaque"
          ? opaquePresentationSnapshot(model)
          : element.content.case === "image"
            ? presentationImageReadOnlySnapshot(model)
            : element.content.case === "table"
              ? presentationTableReadOnlySnapshot(model)
            : undefined,
      });
    }
    for (const entry of entries) capturePresentationConnectorEndpointState(entry.model);
    // Group layout inspection can resolve attached connector endpoints and is
    // therefore not observational until every imported connector has captured
    // its source-bound fingerprint. Snapshot only after that boundary so the
    // preservation optimization cannot mutate the model it is measuring.
    for (const entry of entries) {
      entry.modelSnapshot = entry.wire.content.case === "shape"
        ? presentationImportedShapeSnapshot(entry.model)
        : entry.wire.content.case === "group"
          ? presentationImportedGroupSnapshot(entry.model)
          : undefined;
    }
    for (const sourceThread of sourceSlide.modernComments || []) {
      presentation.commentFormat = "modern";
      const moniker = sourceThread.anchor?.monikers?.[0];
      const textRange = sourceThread.anchor?.kind === PresentationModernCommentAnchor_Kind.TEXT_RANGE;
      const target = slide.resolve(sourceThread.targetId);
      const targetElement = textRange ? slide.resolve(target?.parentId) : target;
      if (targetElement && moniker) {
        targetElement.nativeId = Number(moniker.nativeId);
        targetElement.creationId = moniker.creationId || undefined;
        targetElement.moniker = moniker.type;
      }
      slide.nativeSlideId = Number(sourceThread.anchor?.nativeSlideId || 0) || undefined;
      const sourceComments = [sourceThread.root, ...(sourceThread.replies || [])];
      const comments = sourceComments.map((comment) => ({
        nativeId: comment.id,
        authorId: comment.authorId,
        author: comment.author,
        initials: comment.initials || undefined,
        userId: comment.userId || undefined,
        providerId: comment.providerId || undefined,
        person: {
          id: comment.authorId,
          name: comment.author,
          initials: comment.initials || undefined,
          userId: comment.userId || undefined,
          providerId: comment.providerId || undefined,
        },
        text: comment.text,
        created: comment.createdAt,
        status: comment.status,
      }));
      const root = comments[0];
      slide.comments.addThread(sourceThread.targetId, root.text, {
        id: sourceThread.id,
        author: root.author,
        created: root.created,
        resolved: ["resolved", "closed"].includes(root.status),
        nativeFormat: "modern",
        nativeAnchor: {
          format: "modern",
          type: textRange ? "textRange" : "element",
          nativeId: Number(moniker?.nativeId || 0),
          creationId: moniker?.creationId || undefined,
          moniker: moniker?.type,
          nativeSlideId: Number(sourceThread.anchor?.nativeSlideId || 0),
          ...(textRange ? {
            textStart: Number(sourceThread.anchor?.textStart || 0),
            textLength: Number(sourceThread.anchor?.textLength || 0),
            cp: Number(sourceThread.anchor?.textStart || 0),
            length: Number(sourceThread.anchor?.textLength || 0),
            ...(sourceThread.anchor?.contextLength === undefined ? {} : { contextLength: Number(sourceThread.anchor.contextLength) }),
            ...(sourceThread.anchor?.contextHash === undefined ? {} : { contextHash: Number(sourceThread.anchor.contextHash) }),
          } : {}),
        },
        position: {
          x: Number(sourceThread.positionXEmu || 0),
          y: Number(sourceThread.positionYEmu || 0),
          unit: "emu",
        },
        comments,
      });
    }
    for (const sourceComment of sourceSlide.legacyComments || []) {
      const created = sourceComment.createdAt || new Date(0).toISOString();
      const nativeAuthorId = Number(sourceComment.nativeAuthorId || 0);
      const nativeIndex = Number(sourceComment.nativeIndex || 0);
      const positionXEmu = Number(sourceComment.positionXEmu || 0);
      const positionYEmu = Number(sourceComment.positionYEmu || 0);
      slide.comments.addThread(undefined, sourceComment.text, {
        id: sourceComment.id,
        author: sourceComment.author,
        created,
        nativeFormat: "legacy",
        nativeAnchor: {
          format: "legacy",
          nativeAuthorId,
          nativeIndex,
          positionXEmu,
          positionYEmu,
        },
        position: { x: positionXEmu / EMU_PER_PIXEL, y: positionYEmu / EMU_PER_PIXEL, unit: "px" },
        comments: [{
          nativeId: `legacy:${nativeAuthorId}:${nativeIndex}`,
          author: sourceComment.author,
          text: sourceComment.text,
          created,
        }],
      });
    }
    Object.defineProperty(slide.comments, PRESENTATION_LEGACY_COMMENTS_CAPABILITY, {
      value: Object.freeze({
        sourceBound: true,
        format: sourceSlide.source?.commentFamily || "legacy",
        partPresent: Boolean(sourceSlide.source?.commentPartPresent),
        editable: Boolean(
          sourceSlide.legacyComments?.length &&
          !sourceSlide.modernComments?.length &&
          sourceSlide.source?.legacyCommentsEditable
        ),
        addable: Boolean(
          !sourceSlide.legacyComments?.length &&
          !sourceSlide.modernComments?.length &&
          sourceSlide.source?.legacyCommentsAddable
        ),
      }),
    });
    slideStates.push({
      wire: sourceSlide,
      slide,
      name: slide.name,
      commentSnapshot: presentationSlideCommentSnapshot(slide),
      shellSnapshot: presentationImportedSlideShellSnapshot(slide),
      entries,
    });
  }
  const customShowStates = [];
  for (const sourceShow of source.customShows || []) {
    const model = presentation.customShows.add({
      id: sourceShow.id,
      name: sourceShow.name,
      nativeId: Number(sourceShow.nativeId),
      slideIds: [...sourceShow.slideIds],
    });
    customShowStates.push({ wire: sourceShow, model });
  }
  const sectionStates = [];
  for (const sourceSection of source.sections || []) {
    const model = presentation.sections.add({
      id: sourceSection.id,
      name: sourceSection.name,
      nativeId: sourceSection.nativeId,
      slideIds: [...sourceSection.slideIds],
    });
    sectionStates.push({ wire: sourceSection, model });
  }
  const presentationState = {
    source: envelope.source,
    opaqueOpc: envelope.opaqueOpc,
    assets: assetCatalog.assets(),
    diagnostics: envelope.diagnostics,
    sourceArtifact: source,
    name: source.name,
    slideWidthEmu: source.slideWidthEmu,
    slideHeightEmu: source.slideHeightEmu,
    viewProperties: source.viewProperties,
    customShowsOpaque: Boolean(source.customShowsOpaque),
    customShows: customShowStates,
    sectionsOpaque: Boolean(source.sectionsOpaque),
    sections: sectionStates,
    advancedSnapshot: presentationAdvancedSnapshot(presentation),
    masters: masterStates,
    layouts: layoutStates,
    slides: slideStates,
    clones: [],
  };
  Object.defineProperty(presentation, PRESENTATION_STATE, {
    configurable: true,
    value: presentationState,
    writable: true,
  });
  Object.defineProperty(presentation, PRESENTATION_SLIDE_DUPLICATOR, {
    configurable: true,
    value: (slide) => duplicateImportedPresentationSlide(presentation, presentationState, slide),
  });
  const revisionSha256 = String(presentationState.opaqueOpc?.sourcePackage?.sha256 || presentationState.source?.packageSha256 || "").toLowerCase();
  if (/^[0-9a-f]{64}$/u.test(revisionSha256)) {
    const defineLazyCapability = (symbol, create) => {
      Object.defineProperty(presentation, symbol, {
        configurable: true,
        get() {
          const capability = create();
          Object.defineProperty(presentation, symbol, {
            configurable: true,
            value: capability,
          });
          return capability;
        },
      });
    };
    defineLazyCapability(PRESENTATION_NATIVE_LEAF_CAPABILITY, () =>
      createPresentationNativeLeafCapability(presentation, presentationState));
    defineLazyCapability(PRESENTATION_COMPONENT_CAPABILITY, () =>
      createPresentationComponentCapability(presentation, presentationState));
  }
  return presentation;
}
