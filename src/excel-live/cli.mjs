import { access, lstat, readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { createInterface } from "node:readline/promises";
import { fileURLToPath } from "node:url";

import { bridgeRequest, ensureExcelBridge } from "./client.mjs";
import {
  ensureExcelCertificates,
  persistCertificateMetadata,
  probeExcelRootTrust,
  trustExcelRootCertificate,
  trustInstructions,
  untrustExcelRootCertificate,
} from "./certificates.mjs";
import { excelLiveError } from "./errors.mjs";
import { writeExcelManifest, excelBridgeOrigin } from "./manifest.mjs";
import { validateExcelRequest } from "./protocol.mjs";
import {
  initializeExcelConfiguration,
  readExcelConfiguration,
  removeExcelState,
  resolveExcelStatePaths,
} from "./state.mjs";

const PACKAGE_ROOT = fileURLToPath(new URL("../..", import.meta.url));
const MAX_TASK_FILE_BYTES = 1_000_000;

export const EXCEL_LIVE_USAGE = `Usage:
  officekit excel install [--yes] [--json]
  officekit excel doctor [--json]
  officekit excel sessions --json
  officekit excel execute <request.json> [--json]
  officekit excel disconnect <session-id> [--json]
  officekit excel uninstall --yes [--json]

Control a workbook already open in Microsoft Excel through the local OfficeKit Add-in.
Run officekit excel install once, upload the printed manifest in Excel, then open OfficeKit from the Home ribbon.
`;

export async function runExcelCommand(
  argv,
  {
    input = process.stdin,
    output = process.stdout,
    platform = process.platform,
    statePaths = resolveExcelStatePaths(),
    packageVersion,
    trust = trustExcelRootCertificate,
    untrust = untrustExcelRootCertificate,
    probeTrust = probeExcelRootTrust,
    ensureBridge = ensureExcelBridge,
  } = {},
) {
  const parsed = parseExcelArguments(argv);
  if (parsed.help || parsed.command == null) {
    output.write(`${EXCEL_LIVE_USAGE}\n`);
    return;
  }
  if (parsed.command === "install") {
    assertSupportedDesktopPlatform(platform);
    const result = await installExcel({
      statePaths,
      packageVersion: packageVersion ?? await readPackageVersion(),
      confirmed: await confirmation(parsed, {
        input,
        output,
        prompt: "Trust the local OfficeKit certificate in your user profile? [y/N] ",
      }),
      trust,
      ensureBridge,
    });
    writeResult(output, parsed.json, result, formatInstall);
    return result;
  }
  if (parsed.command === "doctor") {
    assertSupportedDesktopPlatform(platform);
    const result = await doctorExcel({ statePaths, probeTrust, ensureBridge, platform });
    writeResult(output, parsed.json, result, formatDoctor);
    return result;
  }
  if (parsed.command === "sessions") {
    assertSupportedDesktopPlatform(platform);
    const result = await withBridge(statePaths, ensureBridge, (state) =>
      bridgeRequest(state, "GET", "/v1/cli/sessions"));
    assertSuccessfulBridgeResult(result);
    writeResult(output, parsed.json, result, formatSessions);
    return result;
  }
  if (parsed.command === "execute") {
    assertSupportedDesktopPlatform(platform);
    const request = await readExcelRequestFile(parsed.argument);
    const result = await withBridge(statePaths, ensureBridge, (state) =>
      bridgeRequest(state, "POST", "/v1/cli/execute", { request }));
    assertSuccessfulBridgeResult(result);
    writeResult(output, parsed.json, result, formatExecution);
    return result;
  }
  if (parsed.command === "disconnect") {
    assertSupportedDesktopPlatform(platform);
    const result = await withBridge(statePaths, ensureBridge, (state) =>
      bridgeRequest(state, "POST", "/v1/cli/disconnect", { sessionId: parsed.argument }));
    assertSuccessfulBridgeResult(result);
    writeResult(output, parsed.json, result, (value) => `Disconnected ${value.result.disconnected}.\n`);
    return result;
  }
  if (parsed.command === "uninstall") {
    assertSupportedDesktopPlatform(platform);
    const result = await uninstallExcel({
      statePaths,
      confirmed: await confirmation(parsed, {
        input,
        output,
        prompt: "Remove local OfficeKit Excel state and certificate trust? [y/N] ",
      }),
      untrust,
    });
    writeResult(output, parsed.json, result, formatUninstall);
    return result;
  }
  throw excelLiveError("unknown-command", `Unknown Excel command: ${parsed.command}. Run officekit excel --help.`);
}

export async function installExcel({
  statePaths,
  packageVersion,
  confirmed,
  trust = trustExcelRootCertificate,
  ensureBridge = ensureExcelBridge,
}) {
  if (!confirmed) {
    throw excelLiveError("confirmation-required", "Excel certificate trust requires explicit confirmation. Rerun with --yes or confirm interactively.");
  }
  let state = await initializeExcelConfiguration(statePaths);
  const certificate = await ensureExcelCertificates(statePaths, state.config);
  if (!certificateMetadataMatches(state.config.certificate, certificate.certificate)) {
    const config = await persistCertificateMetadata(statePaths, certificate.certificate);
    state = { config, secret: state.secret };
  }
  if (!state.config.trusted) {
    const config = await trust(statePaths, state.config);
    state = { config, secret: state.secret };
  }
  await writeExcelManifest(statePaths, state.config, packageVersion);
  await ensureBridge(statePaths, state);
  return {
    protocol: 1,
    ok: true,
    result: {
      installed: true,
      manifest: statePaths.manifest,
      origin: excelBridgeOrigin(state.config.port),
      trusted: state.config.trusted,
      uploadSteps: [
        "Open a workbook in Microsoft Excel desktop.",
        "Choose Home > Add-ins > My Add-ins > Upload My Add-in.",
        `Select ${statePaths.manifest}.`,
        "Open OfficeKit from the Home ribbon and click Connect OfficeKit.",
      ],
    },
  };
}

export async function doctorExcel({
  statePaths,
  probeTrust = probeExcelRootTrust,
  ensureBridge = ensureExcelBridge,
  platform = process.platform,
}) {
  const state = await readExcelConfiguration(statePaths);
  const [trust, manifest] = await Promise.all([
    probeTrust(statePaths, state.config),
    existsRegularFile(statePaths.manifest),
  ]);
  await ensureBridge(statePaths, state);
  const bridge = await bridgeRequest(state, "GET", "/v1/cli/doctor");
  assertSuccessfulBridgeResult(bridge);
  const sessions = bridge.result.sessions ?? [];
  return {
    ...bridge,
    result: {
      ...bridge.result,
      installation: {
        manifest: statePaths.manifest,
        manifestExists: manifest,
        configuredTrust: state.config.trusted,
        observedTrust: trust,
        trustInstructions: trustInstructions(statePaths),
      },
      host: liveHostDiagnostic(sessions, platform),
    },
  };
}

export async function uninstallExcel({ statePaths, confirmed, untrust = untrustExcelRootCertificate }) {
  if (!confirmed) {
    throw excelLiveError("confirmation-required", "Excel uninstall requires explicit confirmation. Rerun with --yes or confirm interactively.");
  }
  let state = null;
  try {
    state = await readExcelConfiguration(statePaths);
    await bridgeRequest(state, "POST", "/v1/cli/shutdown", {}).catch(() => {});
    if (state.config.trusted) await untrust(statePaths, state.config);
  } catch (error) {
    if (error?.code !== "not-installed") throw error;
  }
  const removed = await removeExcelState(statePaths);
  return {
    protocol: 1,
    ok: true,
    result: {
      removed,
      manualStep: "In Excel, remove OfficeKit from Home > Add-ins > My Add-ins if it is still listed.",
    },
  };
}

async function withBridge(statePaths, ensureBridge, operation) {
  const state = await readExcelConfiguration(statePaths);
  await ensureBridge(statePaths, state);
  return operation(state);
}

async function readExcelRequestFile(requested) {
  if (
    typeof requested !== "string" || requested.length === 0 || requested === "-" ||
    requested.includes("\0") || (/^[a-z][a-z0-9+.-]*:/iu.test(requested) && !(process.platform === "win32" && /^[a-z]:/iu.test(requested)))
  ) {
    throw excelLiveError("invalid-request-file", "officekit excel execute accepts one local JSON request file, not stdin or a URL.");
  }
  const target = path.resolve(requested);
  const stat = await lstat(target).catch((error) => {
    if (error?.code === "ENOENT") return null;
    throw error;
  });
  if (stat == null || stat.isSymbolicLink() || !stat.isFile()) {
    throw excelLiveError("invalid-request-file", `Excel request must be a regular non-symlink file: ${target}`);
  }
  if (stat.size > MAX_TASK_FILE_BYTES) {
    throw excelLiveError("request-too-large", `Excel request exceeds ${MAX_TASK_FILE_BYTES} bytes.`);
  }
  let value;
  try {
    value = JSON.parse(await readFile(target, "utf8"));
  } catch (error) {
    throw excelLiveError("invalid-request-file", `Excel request file is not valid JSON: ${error.message}`);
  }
  return validateExcelRequest(value);
}

function parseExcelArguments(argv) {
  const values = [...argv];
  let command = null;
  let argument = null;
  let yes = false;
  let json = false;
  let help = false;
  for (const value of values) {
    if (value === "--yes" || value === "-y") yes = true;
    else if (value === "--json") json = true;
    else if (value === "--help" || value === "-h") help = true;
    // Keep `-` as a positional value long enough for execute to return its
    // deliberate stdin-rejection error instead of pretending it is an option.
    else if (value.startsWith("-") && value !== "-") throw excelLiveError("invalid-option", `Unknown Excel option: ${value}.`);
    else if (command == null) command = value;
    else if (argument == null) argument = value;
    else throw excelLiveError("invalid-command", `Unexpected Excel argument: ${value}.`);
  }
  if (["execute", "disconnect"].includes(command) && argument == null && !help) {
    throw excelLiveError("invalid-command", `officekit excel ${command} requires an argument.`);
  }
  if (["install", "doctor", "sessions", "uninstall"].includes(command) && argument != null) {
    throw excelLiveError("invalid-command", `officekit excel ${command} does not accept a positional argument.`);
  }
  return { command, argument, yes, json, help };
}

async function confirmation(parsed, { input, output, prompt }) {
  if (parsed.yes) return true;
  if (!input.isTTY || !output.isTTY || parsed.json) return false;
  const terminal = createInterface({ input, output });
  try {
    const answer = await terminal.question(prompt);
    return /^(?:y|yes)$/iu.test(answer.trim());
  } finally {
    terminal.close();
  }
}

function assertSupportedDesktopPlatform(platform) {
  if (platform !== "darwin" && platform !== "win32") {
    throw excelLiveError(
      "unsupported-platform",
      "OfficeKit Excel Live Control currently supports Microsoft Excel desktop on macOS and Windows only.",
    );
  }
}

function assertSuccessfulBridgeResult(result) {
  if (result?.ok === true) return;
  const error = result?.error ?? {};
  throw excelLiveError(
    typeof error.code === "string" ? error.code : "bridge-failure",
    typeof error.message === "string" ? error.message : "OfficeKit Excel bridge failed.",
    { retryable: Boolean(error.retryable), maybeApplied: Boolean(error.maybeApplied), details: error.details },
  );
}

async function readPackageVersion() {
  const metadata = JSON.parse(await readFile(path.join(PACKAGE_ROOT, "package.json"), "utf8"));
  return metadata.version;
}

async function existsRegularFile(target) {
  try {
    await access(target);
    const stat = await lstat(target);
    return stat.isFile() && !stat.isSymbolicLink();
  } catch (error) {
    if (error?.code === "ENOENT") return false;
    throw error;
  }
}

function writeResult(output, json, value, formatter) {
  output.write(json ? `${JSON.stringify(value)}\n` : formatter(value));
}

function formatInstall(value) {
  const result = value.result;
  return [
    "OfficeKit Excel Live Control is ready.",
    ...result.uploadSteps.map((step, index) => `${index + 1}. ${step}`),
    "",
    `Manifest: ${result.manifest}`,
    `Local bridge: ${result.origin}`,
  ].join("\n") + "\n";
}

function formatDoctor(value) {
  const result = value.result;
  const sessions = result.sessions ?? [];
  const host = result.host;
  return [
    `Bridge: ${result.bridge}`,
    `Certificate trust: ${result.installation.observedTrust.trusted ? "ready" : "needs repair"}`,
    `Manifest: ${result.installation.manifestExists ? "ready" : "missing"}`,
    `Excel runtime: ${host?.status ?? "waiting-for-add-in"}`,
    `Live Excel sessions: ${sessions.length}`,
    ...(host?.repairs ?? []),
  ].filter(Boolean).join("\n") + "\n";
}

function formatSessions(value) {
  const sessions = value.result.sessions;
  if (sessions.length === 0) return "No live Excel session. Open OfficeKit in the target workbook and click Connect OfficeKit.\n";
  return `${sessions.map((session) => `${session.id}  ${session.workbook.name} — ${session.workbook.activeSheet}`).join("\n")}\n`;
}

function formatExecution(value) {
  return `${value.ok ? "Excel operation completed." : "Excel operation failed."}\n`;
}

function liveHostDiagnostic(sessions, platform) {
  const required = {
    excelApi: "1.8",
    sharedRuntime: "1.1",
    save: "ExcelApi 1.11",
    desktopWindowForZoom: "ExcelApiDesktop 1.1",
  };
  const reported = sessions.map((session) => ({
    sessionId: session.id,
    host: session.host,
    capabilities: session.capabilities,
    ready: Boolean(session.capabilities?.excelApi18 && session.capabilities?.sharedRuntime),
  }));
  const ready = reported.filter((session) => session.ready);
  const repairs = [];
  if (sessions.length === 0) {
    repairs.push("Open OfficeKit from the target workbook's Home ribbon and click Connect OfficeKit, then run officekit excel sessions --json.");
  } else if (ready.length === 0) {
    repairs.push("The connected Excel runtime does not report ExcelApi 1.8 plus SharedRuntime 1.1. Update desktop Excel and reopen OfficeKit.");
  }
  return {
    platform,
    status: ready.length > 0 ? "ready" : sessions.length > 0 ? "missing-required-capability" : "waiting-for-add-in",
    required,
    sessions: reported,
    repairs,
  };
}

function certificateMetadataMatches(configured, actual) {
  return configured?.rootFingerprintSha256 === actual?.rootFingerprintSha256 &&
    configured?.leafFingerprintSha256 === actual?.leafFingerprintSha256;
}
