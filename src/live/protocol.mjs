import { officeLiveError } from "./errors.mjs";

export const LIVE_PROTOCOL = 1;
export const LIVE_HOSTS = Object.freeze(["excel", "powerpoint"]);
export const MAX_LIVE_REQUEST_BYTES = 1_000_000;
export const MAX_LIVE_RESPONSE_BYTES = 10_000_000;
export const MAX_LIVE_IMAGE_BYTES = 8_000_000;

export function createLiveSuccess({ result, audit } = {}) {
  return {
    protocol: LIVE_PROTOCOL,
    ok: true,
    result: result ?? {},
    audit: audit ?? {},
  };
}

export function createLiveFailure(error, { audit } = {}) {
  const normalized = error?.code
    ? error
    : officeLiveError("internal-error", error instanceof Error ? error.message : String(error));
  return {
    protocol: LIVE_PROTOCOL,
    ok: false,
    error: {
      code: normalized.code,
      message: normalized.message,
      retryable: Boolean(normalized.retryable),
      maybeApplied: Boolean(normalized.maybeApplied),
      ...(normalized.details === undefined ? {} : { details: normalized.details }),
    },
    ...(audit == null ? {} : { audit }),
  };
}

export function liveProtocolReference() {
  return {
    protocol: LIVE_PROTOCOL,
    hosts: LIVE_HOSTS,
    limits: {
      maxRequestBytes: MAX_LIVE_REQUEST_BYTES,
      maxResponseBytes: MAX_LIVE_RESPONSE_BYTES,
      maxImageBytes: MAX_LIVE_IMAGE_BYTES,
    },
  };
}

export function assertLiveHost(value) {
  if (!LIVE_HOSTS.includes(value)) throw officeLiveError("unsupported-host", `Unsupported Office Live host: ${String(value)}.`);
  return value;
}

export function validateLiveEnvelope(value, { host, validateRequest } = {}) {
  if (value == null || typeof value !== "object" || Array.isArray(value)) {
    throw officeLiveError("invalid-request", "Live request must be an object.");
  }
  if (value.protocol !== LIVE_PROTOCOL) {
    throw officeLiveError("invalid-request", `Live request protocol must be ${LIVE_PROTOCOL}.`);
  }
  if (typeof value.sessionId !== "string" || !/^[A-Za-z0-9][A-Za-z0-9._:-]{7,127}$/u.test(value.sessionId)) {
    throw officeLiveError("invalid-request", "sessionId is invalid.");
  }
  if (typeof value.idempotencyKey !== "string" || !/^[A-Za-z0-9][A-Za-z0-9._:-]{7,159}$/u.test(value.idempotencyKey)) {
    throw officeLiveError("invalid-request", "idempotencyKey is invalid.");
  }
  if (typeof value.operation !== "string" || value.operation.length === 0 || value.operation.length > 128) {
    throw officeLiveError("invalid-request", "operation is invalid.");
  }
  if (value.args == null || typeof value.args !== "object" || Array.isArray(value.args)) {
    throw officeLiveError("invalid-request", "args must be an object.");
  }
  if (typeof validateRequest === "function") return validateRequest(value);
  if (host !== undefined && value.host !== undefined && value.host !== host) {
    throw officeLiveError("forbidden", `Live request is not for the ${host} session.`);
  }
  return structuredClone(value);
}
