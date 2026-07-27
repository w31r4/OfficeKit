#!/usr/bin/env node

import { runOfficeKitCli } from "../src/cli/officekit.mjs";

try {
  await runOfficeKitCli(process.argv.slice(2));
} catch (error) {
  const message = error instanceof Error ? error.message : String(error);
  if (process.argv.includes("--json")) {
    process.stderr.write(`${JSON.stringify({ ok: false, error: message })}\n`);
  } else if (error?.officeKitShowStack && error instanceof Error) {
    process.stderr.write(`${error.stack ?? error.message}\n`);
  } else {
    process.stderr.write(`OfficeKit: ${message}\n`);
  }
  process.exitCode = 1;
}
