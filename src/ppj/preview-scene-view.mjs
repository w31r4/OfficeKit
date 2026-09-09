// Internal read-only painter view. No layout, PPJ lowering, OOXML parsing,
// filesystem, raster backend or public Presentation facade belongs here.
import { PresentationPreviewSceneSchema, PresentationElementSchema } from "../generated/office_kit/artifact/v1/office_artifact_pb.js";
import { readPpjPreviewScene } from "./preview-scene.mjs";
import { previewDiagnostic, previewPath } from "./preview-diagnostics.mjs";
import { OfficeKitCodecError } from "../codecs/office-kit-error.mjs";

const freeze = Object.freeze;
const contentFields = new Map(PresentationElementSchema.fields
  .filter(field => field.oneof?.localName === "content").map(field => [field.localName, field]));
const metadata = new WeakMap();
const fail = (message) => { throw new OfficeKitCodecError(message, [], { code: "preview.scene.units" }); };

function scaled(value, divisor, integer = false) {
  if (value === undefined) return undefined;
  if (typeof value === "bigint") {
    // Do not silently round an out-of-range coordinate before division.
    if (value > BigInt(Number.MAX_SAFE_INTEGER) || value < BigInt(Number.MIN_SAFE_INTEGER)) fail("Native coordinate exceeds safe numeric precision.");
    value = Number(value);
  }
  if (typeof value !== "number" || !Number.isFinite(value)) fail("Native visual units must be finite numbers or safe integers.");
  if (integer && !Number.isSafeInteger(value)) fail("Native integral units must retain safe integer precision.");
  return value / divisor;
}

// PPJ canvas coordinates are points (12700 EMU), not CSS px (9525 EMU).
export const scenePoints = (emu) => scaled(emu, 12700, true);
export const sceneDegrees = (angle60000) => scaled(angle60000, 60000, true);
export const sceneOpacity = (thousandthPercent) => scaled(thousandthPercent, 100000, true);
export const sceneFontPoints = (points) => scaled(points, 1);

function rectangle(value, child = false) {
  const keys = child ? ["childLeftEmu", "childTopEmu", "childWidthEmu", "childHeightEmu"]
    : ["leftEmu", "topEmu", "widthEmu", "heightEmu"];
  if (!keys.every(key => value?.[key] !== undefined)) return undefined;
  const [x, y, width, height] = keys.map(key => scenePoints(value[key]));
  return freeze({ x, y, width, height });
}

function transform(value) {
  if (!value) return undefined;
  return freeze({ native: value, rotation: sceneDegrees(value.rotationAngle60000),
    flipH: value.flipHorizontal, flipV: value.flipVertical });
}

function schemaMetadata(schema) {
  if (!metadata.has(schema)) metadata.set(schema, new Set([
    "$typeName", "$unknown", ...schema.fields.map(field => field.oneof?.localName ?? field.localName),
  ]));
  return metadata.get(schema);
}

/** Validate a native.mjs receipt, then expose typed views without copying or
 * flattening native fields. Views are not editable PPJ and imply no paint grade.
 * Asset bytes are borrowed from the verified receipt, never fetched again. */
export function createPpjSceneView(receipt, { limits } = {}) {
  const scene = readPpjPreviewScene({ ...receipt,
    assets: receipt.assets?.map(asset => ({ ...asset, contentType: asset.mimeType ?? asset.contentType })),
  }, receipt.file, { limits });
  const bindings = new Map(scene.bindings.map(binding => [binding.scenePath, binding]));
  const diagnostics = [];
  function limitation(scenePath, reason, value, owner) {
    const identity = owner?.semanticId ? { id: owner.semanticId, pageId: owner.pageId } : {};
    diagnostics.push(freeze({ ...previewDiagnostic({ ...identity, path: owner?.programPath || "$", reason,
      value, action: "Retain native evidence and review this field; a scene adapter does not prove paint support." }), scenePath }));
  }

  // Discover unknown descendants from the actual generated descriptors. A new
  // field is retained by native reference; unknown wire/JS data cannot vanish
  // or silently inherit a parent's future supported grade.
  function inspect(schema, value, scenePath, owner, depth = 0) {
    if (depth > 136) throw new OfficeKitCodecError("Scene inspection exceeds its depth budget.", [], { code: "preview.scene.budget" });
    if (schema === PresentationElementSchema) owner = bindings.get(scenePath);
    for (const key of Object.keys(value)) {
      if (!schemaMetadata(schema).has(key)) limitation(previewPath(scenePath, key), "preview.scene.unknown-field", value[key], owner);
    }
    for (const field of value.$unknown || []) limitation(`${scenePath}.$unknown[${field.no}]`, "preview.scene.unknown-wire-field", field.no, owner);
    for (const oneof of schema.oneofs) {
      const union = value[oneof.localName];
      if (union?.case && !oneof.fields.some(field => field.localName === union.case))
        limitation(previewPath(scenePath, oneof.localName), "preview.scene.unknown-oneof", union.case, owner);
      for (const key of Object.keys(union || {})) if (key !== "case" && key !== "value")
        limitation(previewPath(previewPath(scenePath, oneof.localName), key), "preview.scene.unknown-field", union[key], owner);
    }
    if (schema === PresentationElementSchema && !contentFields.has(value.content?.case))
      limitation(`${scenePath}.content`, "preview.scene.unknown-content", value.content?.case, owner);
    for (const field of schema.fields) {
      const current = field.oneof ? (value[field.oneof.localName]?.case === field.localName ? value[field.oneof.localName].value : undefined) : value[field.localName];
      if (current == null) continue;
      const fieldPath = previewPath(scenePath, field.localName);
      if (field.fieldKind === "message") inspect(field.message, current, fieldPath, owner, depth + 1);
      else if (field.fieldKind === "list" && field.listKind === "message")
        current.forEach((child, i) => inspect(field.message, child, previewPath(fieldPath, i), owner, depth + 1));
      else if (field.fieldKind === "enum" && !field.enum.values.some(item => item.number === current))
        limitation(fieldPath, "preview.scene.unknown-enum", current, owner);
    }
  }
  inspect(PresentationPreviewSceneSchema, scene, "$");

  function group(value, scenePath) {
    return freeze({ native: value, frame: rectangle(value), childFrame: rectangle(value, true),
      transform: transform(value.frameTransform),
      children: freeze(value.children.map((element, i) => node(element, `${scenePath}.children[${i}]`))) });
  }
  function node(element, scenePath) {
    const binding = bindings.get(scenePath), kind = contentFields.has(element.content.case) ? element.content.case : "unknown";
    const native = element.content.value;
    const endpoints = kind === "connector" ? freeze({
      start: freeze({ x: scenePoints(native.startXEmu), y: scenePoints(native.startYEmu) }),
      end: freeze({ x: scenePoints(native.endXEmu), y: scenePoints(native.endYEmu) }),
    }) : undefined;
    const frame = rectangle(native) ?? (endpoints ? freeze({
      x: Math.min(endpoints.start.x, endpoints.end.x), y: Math.min(endpoints.start.y, endpoints.end.y),
      width: Math.abs(endpoints.end.x - endpoints.start.x), height: Math.abs(endpoints.end.y - endpoints.start.y),
    }) : undefined);
    const childGroup = kind === "group" ? group(native, `${scenePath}.group`) : undefined;
    return freeze({ kind, native, element, binding, scenePath, nativeId: element.id,
      pageId: binding.pageId, semanticId: binding.semanticId || undefined, path: binding.programPath,
      hidden: element.hidden, frame, endpoints, transform: transform(native?.transform ?? native?.frameTransform),
      childFrame: childGroup?.childFrame, children: childGroup?.children ?? freeze([]),
      drawing: kind === "diagram" && native.drawing ? group(native.drawing, `${scenePath}.diagram.drawing`) : undefined,
    });
  }
  const payloads = new Map();
  const key = (mime, sha) => `${mime.toLowerCase()}\0${sha.toLowerCase()}`;
  for (const asset of receipt.assets) if (asset.data?.byteLength)
    payloads.set(key(asset.mimeType ?? asset.contentType, asset.sha256), asset.data);
  const assets = freeze(scene.assets.map(reference => freeze({ ...reference,
    data: payloads.get(key(reference.contentType, reference.sha256)),
  })));
  const byAssetId = new Map(assets.map(asset => [asset.nativeId, asset]));
  const pages = freeze(scene.presentation.slides.map((slide, i) => {
    const scenePath = `$.presentation.slides[${i}]`;
    const nodes = freeze(slide.elements.map((element, j) => node(element, `${scenePath}.elements[${j}]`)));
    const pageIds = new Set(nodes.map(element => element.pageId));
    // No ordinal or slide-name guess gives an empty/unmapped page edit authority.
    const pageId = pageIds.size === 1 ? [...pageIds][0] : undefined;
    if (pageIds.size === 0) limitation(scenePath, "preview.scene.unmapped-page", slide.id);
    if (pageIds.size > 1) limitation(scenePath, "preview.scene.ambiguous-page", pageIds.size);
    return freeze({ native: slide, nativeId: slide.id, pageId, scenePath, hidden: slide.hidden, nodes });
  }));
  return freeze({ scene, native: scene.presentation, pages, assets,
    asset: (nativeId) => byAssetId.get(nativeId),
    canvas: freeze({ width: scenePoints(scene.presentation.slideWidthEmu), height: scenePoints(scene.presentation.slideHeightEmu) }),
    // Deliberately not `supported` or `passed`: mapping is not drawing.
    paintAssessment: "unassessed", diagnostics: freeze(diagnostics),
  });
}
