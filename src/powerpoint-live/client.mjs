import { spawn } from "node:child_process";
import process from "node:process";
import { fileURLToPath } from "node:url";

import { bridgeRequest as excelBridgeRequest } from "../excel-live/client.mjs";
import { officeLiveError } from "../live/errors.mjs";

const BRIDGE_SERVER = fileURLToPath(new URL("./bridge-server.mjs", import.meta.url));

export async function ensurePowerPointBridge(paths, state, { timeoutMs = 10_000 } = {}) {
  const existing = await probePowerPointBridge(state).catch(() => null);
  if (existing?.ok === true) return existing;
  const child = spawn(process.execPath, [BRIDGE_SERVER], {
    detached: true,
    stdio: "ignore",
    windowsHide: true,
    env: { ...process.env, OFFICEKIT_POWERPOINT_HOME: paths.root },
  });
  child.unref();
  const deadline = Date.now() + timeoutMs;
  let lastError;
  while (Date.now() < deadline) {
    await delay(120);
    try {
      const result = await probePowerPointBridge(state);
      if (result?.ok === true) return result;
    } catch (error) {
      lastError = error;
    }
  }
  throw officeLiveError("bridge-start-failed", `OfficeKit PowerPoint bridge did not become ready: ${lastError?.message ?? "unknown startup error"}`, { retryable: true });
}

export async function probePowerPointBridge(state) {
  return bridgeRequest(state, "GET", "/v1/cli/health");
}

export async function bridgeRequest(state, method, pathname, body) {
  try {
    return await excelBridgeRequest(state, method, pathname, body);
  } catch (error) {
    if (error?.code) throw officeLiveError(error.code, String(error.message), { retryable: Boolean(error.retryable), maybeApplied: Boolean(error.maybeApplied), details: error.details });
    throw error;
  }
}

export function newIdempotencyKey() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}-powerpoint`;
}

function delay(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}
