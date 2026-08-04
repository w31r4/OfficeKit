import process from "node:process";

import { bridgeRequest, ensureExcelBridge } from "./client.mjs";
import { doctorExcel } from "./cli.mjs";
import { excelLiveError } from "./errors.mjs";
import { validateExcelRequest } from "./protocol.mjs";
import { readExcelConfiguration, resolveExcelStatePaths } from "./state.mjs";

/**
 * Create the typed Excel Live surface exposed by a REPL context.
 *
 * Importing this module is deliberately separate from `src/cli/repl.mjs` so
 * file-only REPL sessions do not load the Excel bridge, certificates, or
 * Office.js-facing control plane. Calling one of these methods is the explicit
 * boundary at which a local Excel operation may start or reuse the bridge.
 */
export function createExcelLiveReplFacade({
  statePaths = resolveExcelStatePaths(),
  platform = process.platform,
  ensureBridge = ensureExcelBridge,
  doctor = doctorExcel,
  bridgeRequestFn = bridgeRequest,
} = {}) {
  return Object.freeze({
    doctor: async (options = {}) => {
      assertSupportedDesktopPlatform(platform);
      return doctor({ statePaths, platform, ensureBridge, ...options });
    },
    sessions: async () => {
      assertSupportedDesktopPlatform(platform);
      return withBridge(statePaths, ensureBridge, (state) =>
        bridgeRequestFn(state, "GET", "/v1/cli/sessions"));
    },
    execute: async (request) => {
      assertSupportedDesktopPlatform(platform);
      let validated;
      try {
        validated = validateExcelRequest(request);
      } catch (error) {
        throw error?.code ? error : excelLiveError("invalid-request", error.message);
      }
      const result = await withBridge(statePaths, ensureBridge, (state) =>
        bridgeRequestFn(state, "POST", "/v1/cli/execute", { request: validated }));
      assertSuccessfulBridgeResult(result);
      return result;
    },
    disconnect: async (sessionId) => {
      assertSupportedDesktopPlatform(platform);
      if (typeof sessionId !== "string" || sessionId.length === 0 || sessionId.length > 128) {
        throw excelLiveError("invalid-session", "Excel sessionId must be a non-empty string of at most 128 characters.");
      }
      const result = await withBridge(statePaths, ensureBridge, (state) =>
        bridgeRequestFn(state, "POST", "/v1/cli/disconnect", { sessionId }));
      assertSuccessfulBridgeResult(result);
      return result;
    },
  });
}

function assertSupportedDesktopPlatform(platform) {
  if (platform !== "darwin" && platform !== "win32") {
    throw excelLiveError(
      "unsupported-platform",
      "OfficeKit Excel Live Control currently supports Microsoft Excel desktop on macOS and Windows only.",
    );
  }
}

async function withBridge(statePaths, ensureBridge, operation) {
  const state = await readExcelConfiguration(statePaths);
  await ensureBridge(statePaths, state);
  return operation(state);
}

function assertSuccessfulBridgeResult(result) {
  if (result?.ok === true) return result;
  const error = result?.error ?? {};
  throw excelLiveError(
    typeof error.code === "string" ? error.code : "bridge-failure",
    typeof error.message === "string" ? error.message : "OfficeKit Excel bridge failed.",
    {
      retryable: Boolean(error.retryable),
      maybeApplied: Boolean(error.maybeApplied),
      details: error.details,
    },
  );
}
