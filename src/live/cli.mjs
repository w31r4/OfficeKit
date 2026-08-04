import { lstat, readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { createInterface } from "node:readline/promises";

import {
  ensureExcelCertificates,
  persistCertificateMetadata,
  probeExcelRootTrust,
  trustExcelRootCertificate,
  trustInstructions,
  untrustExcelRootCertificate,
} from "../excel-live/certificates.mjs";
import { runExcelCommand } from "../excel-live/cli.mjs";
import { officeLiveError } from "./errors.mjs";
import { validatePowerPointRequest } from "./adapters/powerpoint.mjs";
import { bridgeRequest, ensurePowerPointBridge } from "../powerpoint-live/client.mjs";
import { writePowerPointManifest, powerpointBridgeOrigin } from "../powerpoint-live/manifest.mjs";
import {
  initializePowerPointConfiguration,
  readPowerPointConfiguration,
  removePowerPointState,
  resolvePowerPointStatePaths,
  updatePowerPointConfiguration,
  POWERPOINT_BRIDGE_PORT,
} from "../powerpoint-live/state.mjs";

const MAX_TASK_FILE_BYTES = 1_000_000;

export const LIVE_USAGE = `Usage:
  officekit live install --app powerpoint [--yes] [--json]
  officekit live doctor --app powerpoint [--json]
  officekit live sessions --app powerpoint --json
  officekit live execute <request.json> [--json]
  officekit live disconnect <session-id> [--json]
  officekit live uninstall --app powerpoint --yes [--json]

Control a PowerPoint presentation already open in Microsoft PowerPoint through
the local OfficeKit Add-in. Existing officekit excel commands remain supported.
`;

export async function runLiveCommand(argv, options = {}) {
  const parsed = parseLiveArguments(argv);
  if (parsed.help || parsed.command == null) {
    options.output?.write(`${LIVE_USAGE}\n`);
    return;
  }
  if (parsed.app === "excel") {
    return runExcelCommand(parsed.forwarded, options);
  }
  // Execution and disconnect are selected from the opaque session identity.
  // PowerPoint IDs carry the adapter prefix; legacy Excel IDs retain their
  // historical UUID form. This keeps the unqualified command deterministic
  // without probing both hosts or silently retrying on another bridge.
  if (parsed.app == null && parsed.command === "execute") {
    const host = await requestFileHost(parsed.argument);
    if (host === "powerpoint") return runPowerPointCommand({ ...parsed, app: host }, options);
    return runExcelCommand([parsed.command, parsed.argument, ...(parsed.json ? ["--json"] : [])], options);
  }
  if (parsed.app == null && parsed.command === "disconnect") {
    if (String(parsed.argument).startsWith("powerpoint-")) return runPowerPointCommand({ ...parsed, app: "powerpoint" }, options);
    return runExcelCommand([parsed.command, parsed.argument, ...(parsed.json ? ["--json"] : [])], options);
  }
  if (parsed.app !== "powerpoint") throw officeLiveError("unsupported-host", "officekit live currently supports --app powerpoint (or --app excel for the compatibility command).");
  return runPowerPointCommand(parsed, options);
}

export async function runPowerPointCommand(
  parsed,
  {
    input = process.stdin,
    output = process.stdout,
    platform = process.platform,
    statePaths = resolvePowerPointStatePaths(),
    packageVersion,
    trust = trustExcelRootCertificate,
    untrust = untrustExcelRootCertificate,
    probeTrust = probeExcelRootTrust,
    ensureBridge = ensurePowerPointBridge,
  } = {},
) {
  if (parsed.command === "install") {
    assertSupportedDesktopPlatform(platform);
    const result = await installPowerPoint({
      statePaths,
      packageVersion: packageVersion ?? await readPackageVersion(),
      confirmed: await confirmation(parsed, { input, output, prompt: "Trust the local OfficeKit certificate in your user profile? [y/N] " }),
      trust,
      ensureBridge,
    });
    writeResult(output, parsed.json, result, formatInstall);
    return result;
  }
  if (parsed.command === "doctor") {
    assertSupportedDesktopPlatform(platform);
    const result = await doctorPowerPoint({ statePaths, probeTrust, ensureBridge, platform });
    writeResult(output, parsed.json, result, formatDoctor);
    return result;
  }
  if (parsed.command === "sessions") {
    assertSupportedDesktopPlatform(platform);
    const result = await withBridge(statePaths, ensureBridge, (state) => bridgeRequest(state, "GET", "/v1/cli/sessions"));
    assertSuccessfulBridgeResult(result);
    writeResult(output, parsed.json, result, formatSessions);
    return result;
  }
  if (parsed.command === "execute") {
    assertSupportedDesktopPlatform(platform);
    const request = await readPowerPointRequestFile(parsed.argument);
    const result = await withBridge(statePaths, ensureBridge, (state) => bridgeRequest(state, "POST", "/v1/cli/execute", { request }));
    assertSuccessfulBridgeResult(result);
    writeResult(output, parsed.json, result, formatExecution);
    return result;
  }
  if (parsed.command === "disconnect") {
    assertSupportedDesktopPlatform(platform);
    const result = await withBridge(statePaths, ensureBridge, (state) => bridgeRequest(state, "POST", "/v1/cli/disconnect", { sessionId: parsed.argument }));
    assertSuccessfulBridgeResult(result);
    writeResult(output, parsed.json, result, (value) => `Disconnected ${value.result.disconnected}.\n`);
    return result;
  }
  if (parsed.command === "uninstall") {
    assertSupportedDesktopPlatform(platform);
    const result = await uninstallPowerPoint({ statePaths, confirmed: await confirmation(parsed, { input, output, prompt: "Remove local OfficeKit PowerPoint state and certificate trust? [y/N] " }), untrust });
    writeResult(output, parsed.json, result, formatUninstall);
    return result;
  }
  throw officeLiveError("unknown-command", `Unknown live command: ${parsed.command}. Run officekit live --help.`);
}

export async function installPowerPoint({ statePaths, packageVersion, confirmed, trust = trustExcelRootCertificate, ensureBridge = ensurePowerPointBridge }) {
  if (!confirmed) throw officeLiveError("confirmation-required", "PowerPoint certificate trust requires explicit confirmation. Rerun with --yes or confirm interactively.");
  let state = await initializePowerPointConfiguration(statePaths, { port: POWERPOINT_BRIDGE_PORT });
  const certificate = await ensureExcelCertificates(statePaths, state.config);
  if (!certificateMetadataMatches(state.config.certificate, certificate.certificate)) {
    const config = await persistCertificateMetadata(statePaths, certificate.certificate);
    state = { config, secret: state.secret };
  }
  if (!state.config.trusted) {
    const config = await trust(statePaths, state.config);
    state = { config, secret: state.secret };
  }
  await writePowerPointManifest(statePaths, state.config, packageVersion);
  await ensureBridge(statePaths, state);
  return {
    protocol: 1,
    ok: true,
    result: {
      installed: true,
      manifest: statePaths.manifest,
      origin: powerpointBridgeOrigin(state.config.port),
      trusted: state.config.trusted,
      uploadSteps: [
        "Open a presentation in Microsoft PowerPoint desktop.",
        "Choose Home > Add-ins > My Add-ins > Upload My Add-in.",
        `Select ${statePaths.manifest}.`,
        "Open OfficeKit from the Home ribbon and click Connect OfficeKit.",
      ],
    },
  };
}

export async function doctorPowerPoint({ statePaths, probeTrust = probeExcelRootTrust, ensureBridge = ensurePowerPointBridge, platform = process.platform }) {
  const state = await readPowerPointConfiguration(statePaths);
  const [trust, manifest] = await Promise.all([probeTrust(statePaths, state.config), existsRegularFile(statePaths.manifest)]);
  await ensureBridge(statePaths, state);
  const bridge = await bridgeRequest(state, "GET", "/v1/cli/doctor");
  assertSuccessfulBridgeResult(bridge);
  return {
    ...bridge,
    result: {
      ...bridge.result,
      installation: { manifest: statePaths.manifest, manifestExists: manifest, configuredTrust: state.config.trusted, observedTrust: trust, trustInstructions: trustInstructions(statePaths) },
      host: { platform, status: "ready", sessions: bridge.result.sessions ?? [] },
    },
  };
}

export async function uninstallPowerPoint({ statePaths, confirmed, untrust = untrustExcelRootCertificate }) {
  if (!confirmed) throw officeLiveError("confirmation-required", "PowerPoint uninstall requires explicit confirmation. Rerun with --yes or confirm interactively.");
  try {
    const state = await readPowerPointConfiguration(statePaths);
    await bridgeRequest(state, "POST", "/v1/cli/shutdown", {}).catch(() => {});
    if (state.config.trusted) await untrust(statePaths, state.config);
  } catch (error) {
    if (error?.code !== "not-installed") throw error;
  }
  return { protocol: 1, ok: true, result: { removed: await removePowerPointState(statePaths), manualStep: "In PowerPoint, remove OfficeKit from Home > Add-ins > My Add-ins if it is still listed." } };
}

function parseLiveArguments(argv) {
  const values = [...argv];
  let app = null;
  let command = null;
  let argument = null;
  let yes = false;
  let json = false;
  let help = false;
  const forwarded = [];
  while (values.length > 0) {
    const value = values.shift();
    if (value === "--app") app = values.shift();
    else if (value.startsWith("--app=")) app = value.slice("--app=".length);
    else if (value === "--yes" || value === "-y") yes = true;
    else if (value === "--json") json = true;
    else if (value === "--help" || value === "-h") help = true;
    else if (command == null) { command = value; }
    else if (argument == null) { argument = value; }
    else if (app === "excel") forwarded.push(value);
    else throw officeLiveError("invalid-command", `Unexpected live argument: ${value}.`);
  }
  if (app === "excel") {
    if (command != null) forwarded.unshift(command);
    if (argument != null) forwarded.splice(command == null ? 0 : 1, 0, argument);
    if (yes) forwarded.push("--yes");
    if (json) forwarded.push("--json");
    if (help) forwarded.push("--help");
  }
  if (["execute", "disconnect"].includes(command) && argument == null && !help) throw officeLiveError("invalid-command", `officekit live ${command} requires an argument.`);
  return { app, command, argument, yes, json, help, forwarded };
}

async function readPowerPointRequestFile(requested) {
  const isWindowsPath = typeof requested === "string" && /^[A-Za-z]:[\\/]/u.test(requested);
  if (typeof requested !== "string" || requested.length === 0 || requested === "-" || requested.includes("\0") || (!isWindowsPath && /^[a-z][a-z0-9+.-]*:/iu.test(requested))) throw officeLiveError("invalid-request-file", "officekit live execute accepts one local JSON request file, not stdin or a URL.");
  const target = path.resolve(requested);
  const stat = await lstat(target).catch((error) => error?.code === "ENOENT" ? null : Promise.reject(error));
  if (stat == null || stat.isSymbolicLink() || !stat.isFile()) throw officeLiveError("invalid-request-file", `PowerPoint request must be a regular non-symlink file: ${target}`);
  if (stat.size > MAX_TASK_FILE_BYTES) throw officeLiveError("request-too-large", `PowerPoint request exceeds ${MAX_TASK_FILE_BYTES} bytes.`);
  try { return validatePowerPointRequest(JSON.parse(await readFile(target, "utf8"))); } catch (error) { if (error?.code) throw error; throw officeLiveError("invalid-request-file", `PowerPoint request file is not valid JSON: ${error.message}`); }
}

async function requestFileHost(requested) {
  if (typeof requested !== "string" || requested.length === 0 || requested === "-" || requested.includes("\0")) {
    throw officeLiveError("invalid-request-file", "officekit live execute accepts one local JSON request file, not stdin or a URL.");
  }
  const isWindowsPath = /^[A-Za-z]:[\\/]/u.test(requested);
  if (!isWindowsPath && /^[a-z][a-z0-9+.-]*:/iu.test(requested)) {
    throw officeLiveError("invalid-request-file", "officekit live execute accepts one local JSON request file, not stdin or a URL.");
  }
  const target = path.resolve(requested);
  const stat = await lstat(target).catch((error) => error?.code === "ENOENT" ? null : Promise.reject(error));
  if (stat == null || stat.isSymbolicLink() || !stat.isFile()) throw officeLiveError("invalid-request-file", `Live request must be a regular non-symlink file: ${target}`);
  if (stat.size > MAX_TASK_FILE_BYTES) throw officeLiveError("request-too-large", `Live request exceeds ${MAX_TASK_FILE_BYTES} bytes.`);
  let value;
  try { value = JSON.parse(await readFile(target, "utf8")); } catch (error) { throw officeLiveError("invalid-request-file", `Live request file is not valid JSON: ${error.message}`); }
  return typeof value?.sessionId === "string" && value.sessionId.startsWith("powerpoint-") ? "powerpoint" : "excel";
}

async function withBridge(statePaths, ensureBridge, operation) { const state = await readPowerPointConfiguration(statePaths); await ensureBridge(statePaths, state); return operation(state); }
function assertSupportedDesktopPlatform(platform) { if (platform !== "win32" && platform !== "darwin") throw officeLiveError("unsupported-platform", "OfficeKit PowerPoint Live currently supports desktop PowerPoint on Windows and macOS; Windows is the first real acceptance platform."); }
function assertSuccessfulBridgeResult(result) { if (result?.ok === true) return; const error = result?.error ?? {}; throw officeLiveError(error.code ?? "bridge-failure", error.message ?? "OfficeKit PowerPoint bridge failed.", { retryable: Boolean(error.retryable), maybeApplied: Boolean(error.maybeApplied), details: error.details }); }
function certificateMetadataMatches(left, right) { return left?.rootFingerprintSha256 === right?.rootFingerprintSha256 && left?.leafFingerprintSha256 === right?.leafFingerprintSha256; }
async function readPackageVersion() { const metadata = JSON.parse(await readFile(path.join(path.resolve(import.meta.dirname, "../.."), "package.json"), "utf8")); return metadata.version; }
async function existsRegularFile(target) { try { const stat = await lstat(target); return stat.isFile() && !stat.isSymbolicLink(); } catch (error) { if (error?.code === "ENOENT") return false; throw error; } }
async function confirmation(parsed, { input, output, prompt }) { if (parsed.yes) return true; if (!input.isTTY || !output.isTTY || parsed.json) return false; const terminal = createInterface({ input, output }); try { return /^(?:y|yes)$/iu.test((await terminal.question(prompt)).trim()); } finally { terminal.close(); } }
function writeResult(output, json, value, formatter) { output.write(json ? `${JSON.stringify(value)}\n` : formatter(value)); }
function formatInstall(value) { return `PowerPoint Live installed.\nManifest: ${value.result.manifest}\n${value.result.uploadSteps.map((step, index) => `${index + 1}. ${step}`).join("\n")}\n`; }
function formatDoctor(value) { return `PowerPoint Live bridge: ${value.result.bridge}\nSessions: ${value.result.sessions.length}\nManifest: ${value.result.installation.manifest}\n`; }
function formatSessions(value) { return `${value.result.sessions.map((session) => `${session.id} ${session.presentation?.name ?? "presentation"}`).join("\n")}\n`; }
function formatExecution(value) { return `${JSON.stringify(value, null, 2)}\n`; }
function formatUninstall(value) { return `PowerPoint Live removed: ${value.result.removed}\n`; }
