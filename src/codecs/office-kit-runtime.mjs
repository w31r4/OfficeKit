import { FileBlob } from "../shared/file-blob.mjs";
import { OfficeKitCodecError } from "./office-kit-error.mjs";
import {
  isJavaScriptMemoryAllocationError,
  javaScriptMemoryBudgetError,
  OFFICE_KIT_DEFAULT_MAX_INPUT_BYTES,
} from "./office-kit-memory.mjs";
import {
  OFFICE_KIT_NATIVE_TRANSPORT_VERSION,
  startOfficeKitNativeClient,
} from "./office-kit-native-client.mjs";

export const OFFICE_KIT_PROTOCOL_VERSION = 2;

const runtimeStates = Object.freeze({
  office: { promise: undefined, invocationTail: Promise.resolve() },
  ppj: { promise: undefined, invocationTail: Promise.resolve() },
});

async function runtime(profile) {
  const state = runtimeStates[profile];
  if (!state) throw new TypeError(`Unknown OfficeKit native codec profile ${profile}.`);
  while (true) {
    if (!state.promise) {
      state.promise = startOfficeKitNativeClient({ profile })
        .catch((error) => {
          state.promise = undefined;
          if (error instanceof OfficeKitCodecError) throw error;
          throw new OfficeKitCodecError("Bundled OfficeKit NativeAOT Codec could not be loaded.", [], { code: "runtime_unavailable", cause: error });
        });
    }
    const loadedPromise = state.promise;
    const loaded = await loadedPromise;
    if (loaded.tryAcquire()) return loaded;
    if (state.promise === loadedPromise) state.promise = undefined;
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

function effectiveMaxInputBytes(limits = {}) {
  const normalized = codecLimits(limits);
  return normalized.maxInputBytes > 0n ? normalized.maxInputBytes : OFFICE_KIT_DEFAULT_MAX_INPUT_BYTES;
}

export function assertOfficeKitInputBudget(value, limits = {}, family = "Office") {
  const bytes = bytesFrom(value);
  const maximum = effectiveMaxInputBytes(limits);
  if (BigInt(bytes.byteLength) > maximum) {
    throw new OfficeKitCodecError(
      `${family} input has ${bytes.byteLength} bytes and exceeds max_input_bytes (${maximum}).`,
      [],
      { code: "input_budget_exceeded" },
    );
  }
  return bytes;
}

export async function boundedInputBytes(value, limits = {}, family = "Office") {
  return assertOfficeKitInputBudget(await inputBytes(value), limits, family);
}

export async function ownedInputBytes(value, limits = {}, family = "Office") {
  const bytes = await boundedInputBytes(value, limits, family);
  try {
    return Uint8Array.from(bytes);
  } catch (error) {
    if (isJavaScriptMemoryAllocationError(error) || error instanceof RangeError) {
      throw javaScriptMemoryBudgetError(`${family} source ownership`, error);
    }
    throw error;
  }
}

function responseFailure(response) {
  const message = response.diagnostics.length
    ? response.diagnostics.map((item) => `${item.code}: ${item.message}`).join("\n")
    : "OfficeKit codec returned an unspecified failure.";
  return new OfficeKitCodecError(message, response.diagnostics);
}

async function invokeOfficeKitExclusive(profile, request, { fileSidecar = false, consumeResponse } = {}) {
  if (request?.file?.byteLength) assertOfficeKitInputBudget(request.file, request.limits, "Office");
  const state = runtimeStates[profile];
  const loaded = await runtime(profile);
  const loadedPromise = state.promise;
  let response;
  let stage = "request encoding";
  try {
    const sidecar = fileSidecar && request.file?.byteLength ? bytesFrom(request.file) : undefined;
    let wireRequestBytes;
    let decodeResponse;
    if (profile === "ppj") {
      const wire = await import("./office-kit-ppj-wire.mjs");
      wireRequestBytes = wire.encodePpjCodecRequest(request, { omitFile: sidecar != null });
      decodeResponse = wire.decodePpjCodecResponse;
    } else {
      const [{ create, fromBinary, toBinary }, { CodecRequestSchema, CodecResponseSchema }] = await Promise.all([
        import("@bufbuild/protobuf"),
        import("../generated/office_kit/artifact/v1/office_artifact_pb.js"),
      ]);
      const wireRequest = create(CodecRequestSchema, request);
      if (sidecar) wireRequest.file = new Uint8Array();
      wireRequestBytes = toBinary(CodecRequestSchema, wireRequest);
      decodeResponse = (bytes) => fromBinary(CodecResponseSchema, bytes);
    }
    stage = "response decoding";
    const wireResponse = bytesFrom(await loaded.invoke(wireRequestBytes, sidecar));
    response = decodeResponse(wireResponse);
  } catch (error) {
    loaded.kill();
    if (state.promise === loadedPromise) state.promise = undefined;
    if (error instanceof OfficeKitCodecError) throw error;
    if (isJavaScriptMemoryAllocationError(error)) throw javaScriptMemoryBudgetError(stage, error);
    throw new OfficeKitCodecError("OfficeKit native codec returned an invalid protobuf response.", [], { code: "runtime_protocol_mismatch", cause: error });
  } finally {
    loaded.release();
  }
  if (!response.ok) throw responseFailure(response);
  if (typeof consumeResponse === "function") {
    try {
      return await consumeResponse(response);
    } catch (error) {
      if (isJavaScriptMemoryAllocationError(error)) throw javaScriptMemoryBudgetError("artifact hydration", error);
      throw error;
    }
  }
  return response;
}

function assertOfficeKitRequest(request) {
  if (Object.hasOwn(request || {}, "allowLossy") || Object.hasOwn(request || {}, "allow_lossy")) {
    throw new TypeError("invokeOfficeKit no longer accepts allowLossy/allow_lossy; opaque Office content without a validated source package always fails closed.");
  }
  if (request?.file?.byteLength) assertOfficeKitInputBudget(request.file, request.limits, "Office");
}

async function invokeProfileLazy(profile, createRequest, options = {}) {
  if (typeof createRequest !== "function") throw new TypeError("invokeOfficeKitLazy expects a request factory.");
  const state = runtimeStates[profile];
  const invokeCreatedRequest = () => {
    let request;
    try {
      request = createRequest();
    } catch (error) {
      if (isJavaScriptMemoryAllocationError(error)) throw javaScriptMemoryBudgetError("request construction", error);
      throw error;
    }
    assertOfficeKitRequest(request);
    return invokeOfficeKitExclusive(profile, request, options);
  };
  const operation = state.invocationTail.then(
    invokeCreatedRequest,
    invokeCreatedRequest,
  );
  // Never retain the most recent protobuf response through the queue tail.
  state.invocationTail = operation.then(() => undefined, () => undefined);
  return operation;
}

export async function invokeOfficeKitLazy(createRequest, options = {}) {
  return invokeProfileLazy("office", createRequest, options);
}

export async function invokeOfficeKitPpjLazy(createRequest, options = {}) {
  return invokeProfileLazy("ppj", createRequest, options);
}

export async function invokeOfficeKit(request, options = {}) {
  assertOfficeKitRequest(request);
  return invokeOfficeKitLazy(() => request, options);
}

export function assertCodecOptions(options, allowed, apiName) {
  if (options == null || typeof options !== "object" || Array.isArray(options)) throw new TypeError(`${apiName} options must be an object.`);
  const unsupported = Object.keys(options).filter((key) => !allowed.has(key));
  if (unsupported.length) throw new TypeError(`${apiName} does not accept option${unsupported.length === 1 ? "" : "s"} ${unsupported.join(", ")}. OfficeKit is the only Office codec and lossy fallback is unavailable.`);
}

export async function officeKitStatus() {
  const [office, ppj] = await Promise.all([runtime("office"), runtime("ppj")]);
  try {
    return {
      available: true,
      protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
      assemblyName: office.descriptor.assemblyName,
      backend: "native-aot",
      target: office.descriptor.target,
      transportVersion: OFFICE_KIT_NATIVE_TRANSPORT_VERSION,
      manifest: office.descriptor.manifest,
      profiles: Object.freeze({
        office: Object.freeze({ assemblyName: office.descriptor.assemblyName }),
        ppj: Object.freeze({ assemblyName: ppj.descriptor.assemblyName }),
      }),
    };
  } finally {
    office.release();
    ppj.release();
  }
}
