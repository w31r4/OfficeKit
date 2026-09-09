import { createHash } from "node:crypto";
import { fromBinary, toBinary } from "@bufbuild/protobuf";
import {
  PresentationPreviewSceneSchema,
  PresentationArtifactSchema,
  PresentationPreviewSceneOrigin as Origin,
  PresentationPreviewAttribution as Attribution,
} from "../generated/office_kit/artifact/v1/office_artifact_pb.js";
import { OfficeKitCodecError } from "../codecs/office-kit-error.mjs";

export const PPJ_PREVIEW_SCENE_VERSION = 1;
const MAX_BYTES = 128 * 1024 * 1024;
const HASH = /^[a-f0-9]{64}$/u;
const hash = (bytes) => createHash("sha256").update(bytes).digest("hex");
const assetKey = (mime, sha) => `${mime.toLowerCase()}\0${sha.toLowerCase()}`;
const fail = (code, message) => { throw new OfficeKitCodecError(message, [], { code: `preview.scene.${code}` }); };

// Compact JSON-safe identity of an already validated scene. This is evidence,
// not the scene payload or source editing authority.
export function ppjPreviewSceneIdentity(scene) {
  if (!scene) fail("missing", "Preview scene identity is missing.");
  if (scene.version !== PPJ_PREVIEW_SCENE_VERSION) fail("version", "Preview scene identity version is incompatible.");
  const origin = scene.origin === Origin.AUTHORED_LOWERING ? "authored-lowering"
    : scene.origin === Origin.CANDIDATE_IMPORT ? "candidate-import" : undefined;
  if (!origin) fail("origin", "Preview scene identity has an unknown origin.");
  for (const key of ["sha256", "programSha256", "candidateSha256"])
    if (typeof scene[key] !== "string" || !HASH.test(scene[key])) fail("identity", `Preview scene ${key} is not a valid digest.`);
  return Object.freeze({ version: scene.version, origin, sha256: scene.sha256,
    programSha256: scene.programSha256, candidateSha256: scene.candidateSha256 });
}

// Loaded only for an explicit scene request. Native fields and optional presence
// stay intact; this module validates transport, not layout or paint support.
export function readPpjPreviewReceiptScene(receipt, options) {
  return readPpjPreviewScene({ ...receipt,
    assets: receipt.assets?.map(asset => ({ ...asset, contentType: asset.mimeType ?? asset.contentType })),
  }, receipt.file, options);
}

export function readPpjPreviewScene(program, candidate, { limits = {} } = {}) {
  const requestedLimit = BigInt(limits.maxUncompressedBytes ?? 0);
  const maxBytes = requestedLimit > 0n && requestedLimit < BigInt(MAX_BYTES) ? Number(requestedLimit) : MAX_BYTES;
  let scene = program.previewScene;
  if (!scene && program.previewSceneBytes instanceof Uint8Array) {
    if (program.previewSceneBytes.byteLength > maxBytes) fail("budget", "Preview scene exceeds its byte budget.");
    try {
      scene = fromBinary(PresentationPreviewSceneSchema, program.previewSceneBytes, { recursionLimit: 136 });
    } catch (cause) {
      throw new OfficeKitCodecError("Preview scene is not a valid native scene message.", [], { code: "preview.scene.invalid", cause });
    }
  }
  if (!scene) fail("missing", "Native codec returned no preview scene. Rebuild or install the matching OfficeKit codec.");
  if (scene.version !== PPJ_PREVIEW_SCENE_VERSION) fail("version", "Native preview scene version is incompatible. Rebuild or install the matching OfficeKit codec.");
  const expectedOrigin = program.sourceBound ? Origin.CANDIDATE_IMPORT : Origin.AUTHORED_LOWERING;
  if (scene.origin !== expectedOrigin) fail("origin", "Preview scene origin does not match the compile mode.");
  if (!HASH.test(scene.programSha256) || scene.programSha256 !== program.programSha256 ||
      !(program.programJson instanceof Uint8Array) || !program.programJson.byteLength || hash(program.programJson) !== scene.programSha256)
    fail("program-mismatch", "Preview scene does not match the canonical compiled program.");
  if (!HASH.test(scene.candidateSha256) || scene.candidateSha256 !== program.outputSha256 ||
      !(candidate instanceof Uint8Array) || !candidate.byteLength || hash(candidate) !== scene.candidateSha256)
    fail("candidate-mismatch", "Preview scene does not match the actual compiled candidate bytes.");
  if (!scene.presentation || !Array.isArray(scene.presentation.slides) || !Array.isArray(scene.bindings) || !Array.isArray(scene.assets))
    fail("invalid", "Preview scene is missing its presentation, bindings or asset references.");
  // The producer hashes deterministic protobuf, not JSON. Clearing the digest
  // on a shallow envelope copy leaves the received graph/optional values intact.
  let encoded;
  try { encoded = toBinary(PresentationPreviewSceneSchema, { ...scene, sha256: "" }); }
  catch (cause) {
    throw new OfficeKitCodecError("Preview scene cannot be serialized without loss.", [], { code: "preview.scene.invalid", cause });
  }
  if (encoded.byteLength + 66 > maxBytes) fail("budget", "Preview scene exceeds its byte budget.");
  if (!HASH.test(scene.sha256) || hash(encoded) !== scene.sha256) fail("digest-mismatch", "Preview scene digest does not match its native state or provenance.");
  validateBindings(scene);
  validateAssets(scene, program.assets);
  return freezeScene(scene);
}

function validateBindings(scene) {
  const nodes = new Map();
  function walk(elements, path, depth = 0) {
    if (depth > 128) fail("budget", "Preview scene exceeds its node depth budget.");
    for (const [index, element] of elements.entries()) {
      const scenePath = `${path}[${index}]`;
      if (nodes.size >= 100000) fail("budget", "Preview scene exceeds its visual node budget.");
      nodes.set(scenePath, { id: element.id, index });
      if (element.content.case === "group") walk(element.content.value.children, `${scenePath}.group.children`, depth + 1);
      if (element.content.case === "diagram" && element.content.value.drawing)
        walk(element.content.value.drawing.children, `${scenePath}.diagram.drawing.children`, depth + 1);
    }
  }
  for (const [index, slide] of scene.presentation.slides.entries()) walk(slide.elements, `$.presentation.slides[${index}].elements`);
  const paths = new Set();
  for (const binding of scene.bindings) {
    const node = nodes.get(binding.scenePath);
    if (!node || !binding.pageId || !binding.nativeId || node.id !== binding.nativeId || binding.zOrder !== node.index ||
        !binding.programPath.startsWith("$") || binding.programPath.includes("#component(") || paths.has(binding.scenePath) ||
        ![Attribution.DIRECT, Attribution.GENERATED, Attribution.UNMAPPED].includes(binding.attribution) ||
        (binding.attribution !== Attribution.UNMAPPED && !binding.semanticId))
      fail("bindings", "Preview bindings must identify actual nodes, original owner paths and explicit attribution.");
    paths.add(binding.scenePath);
  }
  if (paths.size !== nodes.size) fail("bindings", "Preview scene contains nodes without explicit ownership or unmapped attribution.");
}

function validateAssets(scene, assets) {
  if (!Array.isArray(assets)) fail("asset-mismatch", "Preview receipt has no asset list.");
  const payloads = new Map();
  for (const asset of assets) {
    if (typeof asset.contentType !== "string" || typeof asset.sha256 !== "string")
      fail("asset-mismatch", "Preview asset metadata is invalid.");
    if (!(asset.data instanceof Uint8Array) || !asset.data.byteLength) continue;
    const key = assetKey(asset.contentType, asset.sha256);
    if (!payloads.has(key)) payloads.set(key, []);
    payloads.get(key).push(asset);
  }
  const ids = new Set();
  const verified = new Set();
  for (const reference of scene.assets) {
    if (!reference.nativeId || ids.has(reference.nativeId) || !reference.contentType || !HASH.test(reference.sha256))
      fail("asset-mismatch", "Preview native asset references must be unique and have valid MIME/hash identities.");
    ids.add(reference.nativeId);
    const key = assetKey(reference.contentType, reference.sha256);
    const matches = payloads.get(key);
    if (!matches?.length) fail("asset-mismatch", "Preview scene asset has no matching candidate/compiler payload.");
    if (!verified.has(key)) {
      for (const asset of matches)
        if (hash(asset.data) !== reference.sha256) fail("asset-mismatch", "Preview asset bytes differ from their declared identity.");
      verified.add(key);
    }
  }
  validateNativeReferences(PresentationArtifactSchema, scene.presentation, ids);
}

function validateNativeReferences(schema, message, assets, depth = 0) {
  if (depth > 128) fail("budget", "Preview native message nesting exceeds its budget.");
  for (const field of schema.fields) {
    const value = field.oneof
      ? (message[field.oneof.localName]?.case === field.localName ? message[field.oneof.localName].value : undefined)
      : message[field.localName];
    if (value == null) continue;
    if (field.message && /(?:SourceBinding|Capability|PresentationElementDeletion)$/u.test(field.message.typeName)) {
      if (field.fieldKind !== "list" || value.length) fail("authority", "Read-only preview must not carry source editing authority.");
      continue;
    }
    if ((schema.typeName.endsWith("PresentationOpaqueElement") && field.name === "raw_xml" && value) ||
        (field.name === "replacement_asset_id" && value))
      fail("authority", "Read-only preview must not carry opaque XML or replacement requests.");
    if (field.fieldKind === "message") validateNativeReferences(field.message, value, assets, depth + 1);
    else if (field.fieldKind === "list" && field.listKind === "message")
      for (const child of value) validateNativeReferences(field.message, child, assets, depth + 1);
    else if (field.fieldKind === "map") fail("invalid", "Preview native map fields require an explicit transport policy.");
    else if (field.fieldKind === "scalar") {
      if (value instanceof Uint8Array && value.length) fail("authority", "Preview inline binary requires an explicit asset policy.");
      if (/(?:^|_)asset_id$/u.test(field.name) && value && !assets.has(value))
        fail("asset-mismatch", "Preview native field references an undeclared scene asset.");
    }
  }
}

function freezeScene(value) {
  // Bytes in preserved protobuf unknown fields cannot be frozen by JavaScript.
  // They remain transport evidence, never source-package editing authority.
  if (!value || typeof value !== "object" || ArrayBuffer.isView(value) || Object.isFrozen(value)) return value;
  for (const child of Object.values(value)) freezeScene(child);
  return Object.freeze(value);
}
