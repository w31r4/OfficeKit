#!/usr/bin/env -S node --max-semi-space-size=1

import { runOfficeKitCli } from "../src/cli/officekit.mjs";

try {
  await runOfficeKitCli(process.argv.slice(2));
} catch (error) {
  const message = error instanceof Error ? error.message : String(error);
  if (process.argv.includes("--json")) {
    const failure = error?.code
      ? (await import("../src/excel-live/protocol.mjs")).createExcelFailure(error)
      : { ok: false, error: message };
    process.stderr.write(`${JSON.stringify(failure)}\n`);
  } else if (error?.officeKitShowStack && error instanceof Error) {
    process.stderr.write(`${error.stack ?? error.message}\n`);
  } else {
    process.stderr.write(`OfficeKit: ${message}\n`);
  }
  process.exitCode = 1;
}
