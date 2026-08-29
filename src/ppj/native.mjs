import {
  ArtifactFamily,
  CodecOperation,
} from "../generated/office_kit/artifact/v1/office_artifact_pb.js";
import {
  OFFICE_KIT_PROTOCOL_VERSION,
  codecLimits,
  invokeOfficeKitLazy,
} from "../codecs/office-kit-runtime.mjs";

export const PPJ_MAX_BYTES = 16 * 1024 * 1024;

function bytes(value, name) {
  if (value instanceof Uint8Array) return value;
  if (Buffer.isBuffer(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  throw new TypeError(`${name} must be a Uint8Array.`);
}

function programBytes(value) {
  const result = bytes(value, "PPJ program");
  if (result.byteLength === 0 || result.byteLength > PPJ_MAX_BYTES) {
    throw new RangeError(`PPJ must contain 1 through ${PPJ_MAX_BYTES} UTF-8 bytes.`);
  }
  return result;
}

function result(response) {
  const program = response.presentationProgram;
  if (!program) throw new Error("OfficeKit native codec returned no PPJ receipt.");
  return Object.freeze({
    file: response.file,
    programJson: program.programJson,
    programSha256: program.programSha256,
    nodeMapJson: program.nodeMapJson,
    sourceSha256: program.sourceSha256,
    outputSha256: program.outputSha256,
    changedParts: Object.freeze([...program.changedParts]),
    changedNodeIds: Object.freeze([...program.changedNodeIds]),
    assets: Object.freeze(program.assets.map((asset) => Object.freeze({
      id: asset.id,
      fileName: asset.fileName,
      mimeType: asset.contentType,
      sha256: asset.sha256,
      data: asset.data,
    }))),
    restoredEmbeddedProgram: Boolean(program.restoredEmbeddedProgram),
    sourceBound: Boolean(program.sourceBound),
    expandedElementCount: Number(program.expandedElementCount),
    diagnostics: Object.freeze(response.diagnostics.map((diagnostic) => Object.freeze({
      severity: diagnostic.severity,
      code: diagnostic.code,
      message: diagnostic.message,
      sourcePath: diagnostic.sourcePath,
      sourceIdentity: diagnostic.sourceIdentity,
    }))),
  });
}

export async function projectPptxToPpj(source, {
  sourceUri,
  assetRootUri,
  includeNodeMap = true,
  limits = {},
} = {}) {
  const file = bytes(source, "PPTX source");
  if (!sourceUri || !assetRootUri) throw new TypeError("PPJ projection requires relative sourceUri and assetRootUri values.");
  return invokeOfficeKitLazy(() => ({
    protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
    operation: CodecOperation.PROJECT_PPTX_TO_PPJ,
    family: ArtifactFamily.PRESENTATION,
    file,
    limits: codecLimits(limits),
    presentationProgram: {
      includeNodeMap: Boolean(includeNodeMap),
      sourceUri,
      assetRootUri,
    },
  }), {
    fileSidecar: true,
    consumeResponse: result,
  });
}

export async function compilePpjToPptx(program, {
  source = new Uint8Array(),
  assets = [],
  includeNodeMap = true,
  limits = {},
} = {}) {
  const file = bytes(source, "PPTX source");
  const suppliedAssets = assets.map((asset, index) => {
    if (!asset || typeof asset !== "object") throw new TypeError(`PPJ asset ${index + 1} must be an object.`);
    return {
      id: String(asset.id ?? ""),
      fileName: String(asset.fileName ?? ""),
      contentType: String(asset.mimeType ?? asset.contentType ?? ""),
      sha256: String(asset.sha256 ?? ""),
      data: bytes(asset.data, `PPJ asset ${asset.id || index + 1}`),
    };
  });
  return invokeOfficeKitLazy(() => ({
    protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
    operation: CodecOperation.COMPILE_PPJ_TO_PPTX,
    family: ArtifactFamily.PRESENTATION,
    file,
    limits: codecLimits(limits),
    presentationProgram: {
      programJson: programBytes(program),
      assets: suppliedAssets,
      includeNodeMap: Boolean(includeNodeMap),
    },
  }), {
    fileSidecar: true,
    consumeResponse: result,
  });
}
