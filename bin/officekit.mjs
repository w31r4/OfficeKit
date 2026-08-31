#!/usr/bin/env -S node --max-semi-space-size=1

try {
  const argv = process.argv.slice(2);
  const { execNativePpjBuild } = await import("../src/ppj/native-build-dispatch.mjs");
  if (!await execNativePpjBuild(argv)) {
    const { runOfficeKitCli } = await import("../src/cli/officekit.mjs");
    await runOfficeKitCli(argv);
  }
} catch (error) {
  const message = error instanceof Error ? error.message : String(error);
  if (process.argv.includes("--json")) {
    const failure = error?.code && process.argv[2] !== "image"
      ? (await import("../src/excel-live/protocol.mjs")).createExcelFailure(error)
      : error?.code
        ? { ok: false, error: { code: error.code, message } }
        : { ok: false, error: message };
    process.stderr.write(`${JSON.stringify(failure)}\n`);
  } else if (error?.officeKitShowStack && error instanceof Error) {
    process.stderr.write(`${error.stack ?? error.message}\n`);
  } else {
    process.stderr.write(`OfficeKit: ${message}\n`);
  }
  process.exitCode = 1;
}
