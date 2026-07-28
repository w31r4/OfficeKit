import { spawn } from "node:child_process";
import { randomUUID } from "node:crypto";
import https from "node:https";
import process from "node:process";
import { fileURLToPath } from "node:url";

import { excelLiveError } from "./errors.mjs";

const BRIDGE_SERVER = fileURLToPath(new URL("./bridge-server.mjs", import.meta.url));
const MAX_RESPONSE_BYTES = 10_000_000;

export async function ensureExcelBridge(paths, state, { timeoutMs = 10_000 } = {}) {
  const existing = await probeExcelBridge(state).catch(() => null);
  if (existing?.ok === true) return existing;
  const child = spawn(process.execPath, [BRIDGE_SERVER], {
    detached: true,
    stdio: "ignore",
    windowsHide: true,
    env: {
      ...process.env,
      OFFICEKIT_EXCEL_HOME: paths.root,
    },
  });
  child.unref();
  const deadline = Date.now() + timeoutMs;
  let lastError;
  while (Date.now() < deadline) {
    await delay(120);
    try {
      const result = await probeExcelBridge(state);
      if (result?.ok === true) return result;
    } catch (error) {
      lastError = error;
    }
  }
  throw excelLiveError(
    "bridge-start-failed",
    `OfficeKit Excel bridge did not become ready: ${lastError?.message ?? "unknown startup error"}`,
  );
}

export async function probeExcelBridge(state) {
  return bridgeRequest(state, "GET", "/v1/cli/health");
}

export async function bridgeRequest(state, method, pathname, body) {
  if (state?.config?.certificate == null) {
    throw excelLiveError("not-installed", "Excel certificates are missing. Run officekit excel install.");
  }
  const serialized = body === undefined ? null : JSON.stringify(body);
  return new Promise((resolve, reject) => {
    const request = https.request({
      host: "localhost",
      port: state.config.port,
      method,
      path: pathname,
      // Do not reuse a browser-side TLS socket from the process-wide agent.
      // Pinning is deliberately performed for every CLI request.
      agent: false,
      rejectUnauthorized: false,
      servername: "localhost",
      headers: {
        authorization: `Bearer ${state.secret}`,
        ...(serialized == null ? {} : {
          "content-type": "application/json",
          "content-length": Buffer.byteLength(serialized),
        }),
      },
    }, (response) => {
      const peer = response.socket.getPeerCertificate?.();
      if (peer?.fingerprint256 !== state.config.certificate.leafFingerprintSha256) {
        response.resume();
        reject(excelLiveError("bridge-identity-mismatch", "Excel bridge certificate does not match the local OfficeKit installation."));
        return;
      }
      const chunks = [];
      let received = 0;
      response.on("data", (chunk) => {
        const bytes = Buffer.from(chunk);
        received += bytes.length;
        if (received > MAX_RESPONSE_BYTES) {
          response.destroy(excelLiveError("response-too-large", "Excel bridge returned too much data."));
          return;
        }
        chunks.push(bytes);
      });
      response.once("error", reject);
      response.once("end", () => {
        let parsed;
        try {
          parsed = JSON.parse(Buffer.concat(chunks, received).toString("utf8"));
        } catch (error) {
          reject(excelLiveError("bridge-invalid-response", `Excel bridge returned invalid JSON: ${error.message}`));
          return;
        }
        resolve(parsed);
      });
    });
    request.once("error", (error) => {
      reject(excelLiveError("bridge-unavailable", `OfficeKit Excel bridge is unavailable: ${error.message}`, { retryable: true }));
    });
    request.setTimeout(35_000, () => {
      request.destroy(excelLiveError("bridge-timeout", "OfficeKit Excel bridge did not respond.", { retryable: true }));
    });
    if (serialized != null) request.end(serialized);
    else request.end();
  });
}

export function newIdempotencyKey() {
  return randomUUID();
}

function delay(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}
