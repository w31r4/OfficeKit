import { createHash } from "node:crypto";
import path from "node:path";
import { inflateRawSync } from "node:zlib";
import { OfficeKitCodecError } from "./office-kit-error.mjs";
import { PRESENTATION_PART_BYTES_SOURCE } from "../presentation/native-objects.mjs";

function fail(code, message) {
  throw new OfficeKitCodecError(message, [], { code });
}

function safePartPath(value) {
  const raw = String(value || "");
  const segments = raw.split("/");
  const normalized = path.posix.normalize(raw);
  if (!raw || raw.startsWith("/") || raw.includes("\\") || [...raw].some((character) => character.charCodeAt(0) < 0x20) ||
      segments.some((segment) => !segment || segment === "." || segment === "..") || normalized !== raw) {
    fail("invalid_presentation_native_graph", `OfficeKit returned an unsafe native-object part path: ${value}`);
  }
  return normalized;
}

function relationshipKey(sourcePath, id) {
  return `${String(sourcePath || "")}\0${String(id || "")}`;
}

function modelRelationship(relationship) {
  return {
    id: relationship.id,
    type: relationship.type,
    target: relationship.target,
    targetMode: relationship.targetMode,
  };
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function lazyZipPartBytes(entry, metadata, partPath) {
  const compressed = entry?.compressedContent;
  const compressedSize = Number(entry?.compressedSize);
  const byteLength = Number(entry?.uncompressedSize);
  const compression = entry?.compression;
  const sourceSha256 = String(metadata.sha256 || "").toLowerCase();
  if (!ArrayBuffer.isView(compressed) || compressed.byteLength !== compressedSize ||
      !Number.isSafeInteger(byteLength) || byteLength < 0 || !/^[0-9a-f]{64}$/u.test(sourceSha256) ||
      (compression !== "\x00\x00" && compression !== "\x08\x00")) {
    fail("invalid_opc_package", `OfficeKit source package snapshot has unsupported ZIP metadata for native-object part ${partPath}.`);
  }
  let resolved;
  const resolve = () => {
    if (resolved) return resolved;
    try {
      if (compression === "\x00\x00") {
        // Stored entries alias the source package through the ZIP index. Copy before
        // exposing mutable public bytes so callers cannot mutate the package.
        resolved = Uint8Array.from(compressed);
      } else {
        const inflated = inflateRawSync(compressed, byteLength > 0 ? { maxOutputLength: byteLength } : undefined);
        resolved = new Uint8Array(inflated.buffer, inflated.byteOffset, inflated.byteLength);
      }
    } catch (error) {
      fail("invalid_opc_package", `OfficeKit source package native-object part ${partPath} could not be inflated: ${error.message}`);
    }
    if (resolved.byteLength !== byteLength) {
      fail("invalid_opc_package", `OfficeKit source package native-object part ${partPath} changed its declared byte length.`);
    }
    if (sha256(resolved) !== sourceSha256) {
      fail("presentation_native_part_hash_mismatch", `OfficeKit native-object part ${partPath} does not match its opaque graph hash.`);
    }
    return resolved;
  };
  const install = (target) => {
    let overridden = false;
    let override;
    const get = () => overridden ? override : resolve();
    const set = (value) => {
      overridden = true;
      override = value;
    };
    const binding = {
      install,
      trustedSnapshot(candidate) {
        const descriptor = Object.getOwnPropertyDescriptor(candidate, "bytes");
        return !overridden && resolved === undefined && descriptor?.get === get && descriptor?.set === set
          ? { byteLength, sha256: sourceSha256 }
          : undefined;
      },
    };
    Object.defineProperty(target, PRESENTATION_PART_BYTES_SOURCE, { value: binding });
    Object.defineProperty(target, "bytes", { configurable: true, enumerable: true, get, set });
  };
  return { install };
}

function inflateIndexedZipPart(entry, partPath) {
  const compressed = entry?.compressedContent;
  const byteLength = Number(entry?.uncompressedSize);
  if (!ArrayBuffer.isView(compressed) || !Number.isSafeInteger(byteLength) || byteLength < 0) {
    fail("invalid_opc_package", `OfficeKit source package snapshot has invalid ZIP metadata for native-object part ${partPath}.`);
  }
  try {
    const output = entry.compression === "\x00\x00"
      ? Uint8Array.from(compressed)
      : inflateRawSync(compressed, byteLength > 0 ? { maxOutputLength: byteLength } : undefined);
    const bytes = output.constructor === Uint8Array
      ? output
      : new Uint8Array(output.buffer, output.byteOffset, output.byteLength);
    if (bytes.byteLength !== byteLength) {
      fail("invalid_opc_package", `OfficeKit source package native-object part ${partPath} changed its declared byte length.`);
    }
    return bytes;
  } catch (error) {
    if (error instanceof OfficeKitCodecError) throw error;
    fail("invalid_opc_package", `OfficeKit source package native-object part ${partPath} could not be inflated: ${error.message}`);
  }
}

function indexedZipEntry(entry) {
  if (!entry?._data) return entry;
  return {
    compressedContent: entry._data.compressedContent,
    compressedSize: entry._data.compressedSize,
    uncompressedSize: entry._data.uncompressedSize,
    compression: entry._data.compression?.magic,
  };
}

// Materialize only the bounded part closure selected by the C# codec. The
// complete source package remains canonical in opaque_opc; byte extraction is
// needed solely so the ordinary JS presentation model can retain the same
// read-only native object until its next canonical OfficeKit export.
export async function materializePresentationNativeGraphs(envelope, options = {}) {
  const opaqueOpc = envelope.opaqueOpc;
  const assetBytesBySha256 = options.assetBytesBySha256 instanceof Map
    ? options.assetBytesBySha256
    : new Map();
  const opaqueElements = [];
  if (envelope.payload?.case === "presentation") {
    for (const slide of envelope.payload.value.slides) {
      for (const element of slide.elements) {
        if (element.content?.case === "opaque") opaqueElements.push(element.content.value);
      }
    }
  }
  const requestedPaths = new Set();
  for (const opaque of opaqueElements) {
    for (const partPath of opaque.preservedPartPaths || []) requestedPaths.add(safePartPath(partPath));
  }
  const partsByPath = new Map();
  for (const part of opaqueOpc?.parts || []) {
    const partPath = safePartPath(part.path);
    if (partsByPath.has(partPath)) fail("invalid_presentation_native_graph", `OfficeKit returned duplicate opaque part metadata for ${partPath}.`);
    partsByPath.set(partPath, part);
  }
  for (const partPath of requestedPaths) {
    if (!partsByPath.has(partPath)) fail("missing_presentation_native_part", `OfficeKit native-object graph references missing part metadata ${partPath}.`);
  }

  let zip;
  let zipIndex;
  const sourceBytes = opaqueOpc?.sourcePackage?.data;
  if ([...requestedPaths].some((partPath) => !(partsByPath.get(partPath)?.data?.length))) {
    if (!sourceBytes?.length) fail("missing_source_package", "OfficeKit native-object graph cannot be materialized because its source package snapshot is missing.");
    try {
      const { tryIndexOoxmlZipWithinBudget } = await import("../ooxml/zip-index.mjs");
      zipIndex = tryIndexOoxmlZipWithinBudget(sourceBytes, options, "PPTX");
      if (!zipIndex) {
        const { loadOoxmlZipWithinBudget } = await import("../ooxml/package.mjs");
        zip = await loadOoxmlZipWithinBudget(sourceBytes, options, "PPTX");
      }
    } catch (error) {
      fail("invalid_opc_package", `OfficeKit source package snapshot is not a readable ZIP package: ${error.message}`);
    }
  }

  const materializedParts = new Map();
  // Inflate selected parts one at a time. The codec already bounds their total
  // size, but parallel inflation could otherwise make that entire budget live
  // at once in JSZip plus the hydrated presentation graph.
  for (const partPath of requestedPaths) {
    const metadata = partsByPath.get(partPath);
    const sharedAssetBytes = assetBytesBySha256.get(String(metadata.sha256 || "").toLowerCase());
    let bytes = metadata.data?.length
      ? new Uint8Array(metadata.data)
      : sharedAssetBytes?.length
        ? new Uint8Array(sharedAssetBytes.buffer, sharedAssetBytes.byteOffset, sharedAssetBytes.byteLength)
        : undefined;
    let lazyBytes;
    if (!bytes) {
      const entry = zipIndex?.get(partPath) || zip?.file(partPath);
      if (!entry) fail("missing_presentation_native_part", `OfficeKit source package snapshot is missing native-object part ${partPath}.`);
      if (metadata.sha256) lazyBytes = lazyZipPartBytes(indexedZipEntry(entry), metadata, partPath);
      else bytes = zipIndex ? inflateIndexedZipPart(entry, partPath) : await entry.async("uint8array");
    }
    if (bytes && metadata.sha256 && sha256(bytes) !== metadata.sha256.toLowerCase()) {
      fail("presentation_native_part_hash_mismatch", `OfficeKit native-object part ${partPath} does not match its opaque graph hash.`);
    }
    const materialized = {
      path: partPath,
      contentType: metadata.contentType || "application/octet-stream",
      ...(bytes ? { bytes } : {}),
      sourceSha256: (metadata.sha256 || sha256(bytes)).toLowerCase(),
      relationships: (metadata.relationships || []).map(modelRelationship),
    };
    if (lazyBytes) lazyBytes.install(materialized);
    materializedParts.set(partPath, materialized);
  }

  const relationships = new Map();
  for (const relationship of opaqueOpc?.packageRelationships || []) {
    const key = relationshipKey(relationship.sourcePath, relationship.id);
    if (relationships.has(key)) fail("invalid_presentation_native_graph", `OfficeKit returned duplicate relationship ${relationship.id} from ${relationship.sourcePath}.`);
    relationships.set(key, relationship);
  }

  return function nativeGraph(opaque, sourcePart) {
    const references = (opaque.relationshipReferences || []).map((reference) => ({
      attribute: reference.attribute,
      id: reference.relationshipId,
      namespaceUri: reference.namespaceUri,
    }));
    const rootRelationships = [];
    const seenIds = new Set();
    for (const reference of references) {
      if (seenIds.has(reference.id)) continue;
      seenIds.add(reference.id);
      const relationship = relationships.get(relationshipKey(sourcePart, reference.id));
      if (!relationship) fail("missing_presentation_native_relationship", `OfficeKit native object in ${sourcePart} references missing relationship ${reference.id}.`);
      rootRelationships.push(modelRelationship(relationship));
    }
    const parts = (opaque.preservedPartPaths || []).map((partPath) => materializedParts.get(safePartPath(partPath)));
    if (parts.some((part) => !part)) fail("missing_presentation_native_part", `OfficeKit native-object graph contains unresolved part metadata.`);
    return { relationshipReferences: references, rootRelationships, parts };
  };
}

export function presentationNativeGraphSnapshot(object, { ignoredPartPaths = [] } = {}) {
  const ignored = new Set(ignoredPartPaths);
  return {
    relationshipReferences: object.relationshipReferences,
    rootRelationships: object.rootRelationships,
    parts: (object.parts || []).map((part) => {
      const trusted = ignored.has(part.path)
        ? undefined
        : part?.[PRESENTATION_PART_BYTES_SOURCE]?.trustedSnapshot?.(part);
      const bytes = ignored.has(part.path) || trusted ? undefined : part.bytes;
      return {
        path: part.path,
        contentType: part.contentType,
        relationships: part.relationships,
        sourceSha256: part.sourceSha256,
        ...(ignored.has(part.path) ? {} : {
          bytes: trusted?.byteLength ?? bytes?.length ?? 0,
          sha256: trusted?.sha256 ?? sha256(bytes || []),
        }),
      };
    }),
  };
}
