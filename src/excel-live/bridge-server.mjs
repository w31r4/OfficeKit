import { readFile, rm } from "node:fs/promises";
import process from "node:process";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { startExcelBridge } from "./bridge.mjs";
import { readCertificateBundle } from "./certificates.mjs";
import { readExcelConfiguration, resolveExcelStatePaths, writePrivateText } from "./state.mjs";

const packageRoot = fileURLToPath(new URL("../..", import.meta.url));
const packageMetadata = JSON.parse(await readFile(path.join(packageRoot, "package.json"), "utf8"));
const paths = resolveExcelStatePaths();
const { config, secret } = await readExcelConfiguration(paths);
if (config.certificate == null) throw new Error("OfficeKit Excel certificates are missing. Run officekit excel install.");
const certificate = await readCertificateBundle(paths, config.certificate);
const bridge = await startExcelBridge({
  paths,
  config,
  secret,
  certificate,
  packageVersion: packageMetadata.version,
});
await writePrivateText(paths.pid, `${process.pid}\n`);

let shuttingDown = false;
async function shutdown() {
  if (shuttingDown) return;
  shuttingDown = true;
  await bridge.close().catch(() => {});
  await rm(paths.pid, { force: true }).catch(() => {});
}

process.once("SIGINT", () => { void shutdown().finally(() => process.exit(0)); });
process.once("SIGTERM", () => { void shutdown().finally(() => process.exit(0)); });

const idleTimer = setInterval(() => {
  if (!bridge.isIdle()) return;
  void shutdown().finally(() => process.exit(0));
}, 15_000);
idleTimer.unref();
