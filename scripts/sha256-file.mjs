#!/usr/bin/env node
/** Print the SHA-256 of one regular, non-symlink release input. */

import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";

function fail(message) {
  throw new Error(`SHA-256 file verification: ${message}`);
}

async function main() {
  const [filePath] = process.argv.slice(2);
  if (!filePath || process.argv.length !== 3) fail("usage is sha256-file.mjs <file>.");
  const absolute = path.resolve(filePath);
  const stat = await fs.promises.lstat(absolute);
  if (!stat.isFile() || stat.isSymbolicLink()) fail(`input must be a regular non-symlink file: ${absolute}.`);

  const digest = crypto.createHash("sha256");
  for await (const chunk of fs.createReadStream(absolute)) digest.update(chunk);
  process.stdout.write(`${digest.digest("hex")}\n`);
}

main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 2;
});
