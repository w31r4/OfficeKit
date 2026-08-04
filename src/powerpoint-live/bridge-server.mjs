import { readFile } from "node:fs/promises";
import process from "node:process";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { startPowerPointBridge } from "../live/bridge.mjs";
import { readCertificateBundle } from "../excel-live/certificates.mjs";
import { readPowerPointConfiguration, resolvePowerPointStatePaths } from "./state.mjs";

const packageRoot = fileURLToPath(new URL("../..", import.meta.url));
const packageMetadata = JSON.parse(await readFile(path.join(packageRoot, "package.json"), "utf8"));
const paths = resolvePowerPointStatePaths();
const { config, secret } = await readPowerPointConfiguration(paths);
if (config.certificate == null) throw new Error("OfficeKit PowerPoint certificates are missing. Run officekit live install --app powerpoint.");
const certificate = await readCertificateBundle(paths, config.certificate);
const bridge = await startPowerPointBridge({ paths, config, secret, certificate, packageVersion: packageMetadata.version });

let shuttingDown = false;
async function shutdown() {
  if (shuttingDown) return;
  shuttingDown = true;
  await bridge.close().catch(() => {});
}

process.once("SIGINT", () => { void shutdown().finally(() => process.exit(0)); });
process.once("SIGTERM", () => { void shutdown().finally(() => process.exit(0)); });

const idleTimer = setInterval(() => {
  if (!bridge.isIdle()) return;
  void shutdown().finally(() => process.exit(0));
}, 15_000);
idleTimer.unref();
