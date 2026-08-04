import process from "node:process";

import { bridgeRequest, ensurePowerPointBridge } from "./client.mjs";
import { doctorPowerPoint } from "../live/cli.mjs";
import { officeLiveError } from "../live/errors.mjs";
import { validatePowerPointRequest } from "../live/adapters/powerpoint.mjs";
import { readPowerPointConfiguration, resolvePowerPointStatePaths } from "./state.mjs";

export function createPowerPointLiveReplFacade({
  statePaths = resolvePowerPointStatePaths(),
  platform = process.platform,
  ensureBridge = ensurePowerPointBridge,
  doctor = doctorPowerPoint,
  bridgeRequestFn = bridgeRequest,
} = {}) {
  return Object.freeze({
    doctor: async (options = {}) => {
      assertSupportedDesktopPlatform(platform);
      return doctor({ statePaths, platform, ensureBridge, ...options });
    },
    sessions: async () => {
      assertSupportedDesktopPlatform(platform);
      return withBridge(statePaths, ensureBridge, (state) => bridgeRequestFn(state, "GET", "/v1/cli/sessions"));
    },
    execute: async (request) => {
      assertSupportedDesktopPlatform(platform);
      const validated = validatePowerPointRequest(request);
      const result = await withBridge(statePaths, ensureBridge, (state) => bridgeRequestFn(state, "POST", "/v1/cli/execute", { request: validated }));
      assertSuccessfulBridgeResult(result);
      return result;
    },
    disconnect: async (sessionId) => {
      assertSupportedDesktopPlatform(platform);
      if (typeof sessionId !== "string" || sessionId.length === 0 || sessionId.length > 128) throw officeLiveError("invalid-session", "PowerPoint sessionId must be a non-empty string of at most 128 characters.");
      const result = await withBridge(statePaths, ensureBridge, (state) => bridgeRequestFn(state, "POST", "/v1/cli/disconnect", { sessionId }));
      assertSuccessfulBridgeResult(result);
      return result;
    },
  });
}

function assertSupportedDesktopPlatform(platform) {
  if (platform !== "win32" && platform !== "darwin") throw officeLiveError("unsupported-platform", "PowerPoint Live requires desktop PowerPoint on Windows or macOS; Windows is the first real acceptance platform.");
}

async function withBridge(statePaths, ensureBridge, operation) {
  const state = await readPowerPointConfiguration(statePaths);
  await ensureBridge(statePaths, state);
  return operation(state);
}

function assertSuccessfulBridgeResult(result) {
  if (result?.ok === true) return result;
  const error = result?.error ?? {};
  throw officeLiveError(error.code ?? "bridge-failure", error.message ?? "OfficeKit PowerPoint bridge failed.", { retryable: Boolean(error.retryable), maybeApplied: Boolean(error.maybeApplied), details: error.details });
}
