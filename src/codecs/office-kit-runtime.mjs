import { create, fromBinary, toBinary } from "@bufbuild/protobuf";

import {
  CodecRequestSchema,
  CodecResponseSchema,
} from "../generated/office_kit/artifact/v1/office_artifact_pb.js";
import { FileBlob } from "../shared/file-blob.mjs";
import { OfficeKitCodecError } from "./office-kit-error.mjs";
import {
  OFFICE_KIT_NATIVE_TRANSPORT_VERSION,
  startOfficeKitNativeClient,
} from "./office-kit-native-client.mjs";

export const OFFICE_KIT_PROTOCOL_VERSION = 2;

let runtimePromise;

async function runtime() {
  while (true) {
    if (!runtimePromise) {
      runtimePromise = startOfficeKitNativeClient()
        .catch((error) => {
          runtimePromise = undefined;
          if (error instanceof OfficeKitCodecError) throw error;
          throw new OfficeKitCodecError("Bundled OfficeKit NativeAOT Codec could not be loaded.", [], { code: "runtime_unavailable", cause: error });
        });
    }
    const loadedPromise = runtimePromise;
    const loaded = await loadedPromise;
    if (loaded.tryAcquire()) return loaded;
    if (runtimePromise === loadedPromise) runtimePromise = undefined;
  }
}

function uint64(value, name) {
  if (value == null) return 0n;
  if (typeof value === "bigint") {
    if (value < 0n) throw new TypeError(`${name} must be a non-negative integer.`);
    return value;
  }
  if (!Number.isSafeInteger(value) || value < 0) throw new TypeError(`${name} must be a non-negative safe integer or bigint.`);
  return BigInt(value);
}

export function uint32(value, name) {
  if (value == null) return 0;
  if (!Number.isInteger(value) || value < 0 || value > 0xffff_ffff) throw new TypeError(`${name} must be an unsigned 32-bit integer.`);
  return value;
}

export function codecLimits(limits = {}) {
  return {
    maxInputBytes: uint64(limits.maxInputBytes, "maxInputBytes"),
    maxUncompressedBytes: uint64(limits.maxUncompressedBytes, "maxUncompressedBytes"),
    maxParts: uint32(limits.maxParts, "maxParts"),
    maxSheets: uint32(limits.maxSheets, "maxSheets"),
    maxCells: uint64(limits.maxCells, "maxCells"),
    maxCompressionRatio: uint32(limits.maxCompressionRatio, "maxCompressionRatio"),
  };
}

function bytesFrom(value) {
  if (value instanceof Uint8Array) {
    return value.constructor === Uint8Array
      ? value
      : new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  throw new TypeError("Expected FileBlob, Uint8Array, ArrayBuffer, or ArrayBuffer view.");
}

export async function inputBytes(value) {
  if (value instanceof FileBlob) return bytesFrom(value.bytes);
  return bytesFrom(value);
}

function responseFailure(response) {
  const message = response.diagnostics.length
    ? response.diagnostics.map((item) => `${item.code}: ${item.message}`).join("\n")
    : "OfficeKit codec returned an unspecified failure.";
  return new OfficeKitCodecError(message, response.diagnostics);
}

export async function invokeOfficeKit(request, { fileSidecar = false } = {}) {
  if (Object.hasOwn(request || {}, "allowLossy") || Object.hasOwn(request || {}, "allow_lossy")) {
    throw new TypeError("invokeOfficeKit no longer accepts allowLossy/allow_lossy; opaque Office content without a validated source package always fails closed.");
  }
  const loaded = await runtime();
  const loadedPromise = runtimePromise;
  let response;
  try {
    const wireRequest = create(CodecRequestSchema, request);
    const sidecar = fileSidecar && wireRequest.file?.byteLength ? bytesFrom(wireRequest.file) : undefined;
    if (sidecar) wireRequest.file = new Uint8Array();
    const wireRequestBytes = toBinary(CodecRequestSchema, wireRequest);
    const wireResponse = bytesFrom(await loaded.invoke(wireRequestBytes, sidecar));
    response = fromBinary(CodecResponseSchema, wireResponse);
  } catch (error) {
    loaded.kill();
    if (runtimePromise === loadedPromise) runtimePromise = undefined;
    if (error instanceof OfficeKitCodecError) throw error;
    throw new OfficeKitCodecError("OfficeKit native codec returned an invalid protobuf response.", [], { code: "runtime_protocol_mismatch", cause: error });
  } finally {
    loaded.release();
  }
  if (!response.ok) throw responseFailure(response);
  return response;
}

export function assertCodecOptions(options, allowed, apiName) {
  if (options == null || typeof options !== "object" || Array.isArray(options)) throw new TypeError(`${apiName} options must be an object.`);
  const unsupported = Object.keys(options).filter((key) => !allowed.has(key));
  if (unsupported.length) throw new TypeError(`${apiName} does not accept option${unsupported.length === 1 ? "" : "s"} ${unsupported.join(", ")}. OfficeKit is the only Office codec and lossy fallback is unavailable.`);
}

export async function officeKitStatus() {
  const loaded = await runtime();
  try {
    return {
      available: true,
      protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
      assemblyName: loaded.descriptor.assemblyName,
      backend: "native-aot",
      target: loaded.descriptor.target,
      transportVersion: OFFICE_KIT_NATIVE_TRANSPORT_VERSION,
      manifest: loaded.descriptor.manifest,
    };
  } finally {
    loaded.release();
  }
}
