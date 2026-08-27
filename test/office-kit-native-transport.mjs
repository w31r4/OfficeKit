import assert from "node:assert/strict";
import { appendFile, chmod, cp, mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";

import {
  loadOfficeKitNativeDescriptor,
  officeKitNativeTarget,
  startOfficeKitNativeClient,
} from "../src/codecs/office-kit-native-client.mjs";

const target = officeKitNativeTarget();
assert.equal(target, `${process.platform}-${process.arch}`);
assert.throws(() => officeKitNativeTarget("freebsd", "x64"), (error) => error?.code === "runtime_unsupported_platform");

const installed = await loadOfficeKitNativeDescriptor();
assert.equal(installed.target, target);
assert.equal(installed.manifest.backend, "native-aot");
assert.equal(installed.manifest.transportVersion, 1);
assert.equal(installed.manifest.protocolVersion, 2);

const temporary = await mkdtemp(path.join(os.tmpdir(), "office-kit-native-transport-"));
try {
  const tamperedRoot = path.join(temporary, "tampered");
  await cp(installed.packageRoot, tamperedRoot, { recursive: true });
  const executableName = process.platform === "win32" ? "officekit-codec.exe" : "officekit-codec";
  const tamperedExecutable = path.join(tamperedRoot, "bin", executableName);
  await appendFile(tamperedExecutable, Buffer.from([0]));
  if (process.platform !== "win32") await chmod(tamperedExecutable, 0o755);
  await assert.rejects(
    loadOfficeKitNativeDescriptor({ packageJsonPath: path.join(tamperedRoot, "package.json") }),
    (error) => error?.code === "runtime_integrity_failure",
  );

  const fakePath = path.join(temporary, "fake-codec.mjs");
  await writeFile(fakePath, `
import process from "node:process";
const mode = process.argv[2];
const handshake = Buffer.alloc(12);
handshake.write("OKIT", 0, "ascii");
handshake.writeUInt32BE(1, 4);
handshake.writeUInt32BE(2, 8);
if (mode === "bad-handshake") handshake.write("FAIL", 0, "ascii");
process.stdout.write(handshake);
if (mode === "bad-handshake") setTimeout(() => process.exit(0), 500);
let pending = Buffer.alloc(0);
process.stdin.on("data", (chunk) => {
  if (mode === "crash") process.exit(77);
  pending = Buffer.concat([pending, chunk]);
  while (pending.length >= 4) {
    const length = pending.readUInt32BE(0);
    if (pending.length < 4 + length) return;
    const request = pending.subarray(4, 4 + length);
    pending = pending.subarray(4 + length);
    const prefix = Buffer.alloc(4);
    if (mode === "truncated") {
      prefix.writeUInt32BE(8, 0);
      process.stdout.write(prefix);
      process.stdout.write(Buffer.from([1, 2]));
      setTimeout(() => process.exit(78), 20);
      return;
    }
    prefix.writeUInt32BE(request.length, 0);
    process.stdout.write(prefix);
    process.stdout.write(request);
  }
});
`, "utf8");

  const fakeDescriptor = {
    target,
    packageRoot: temporary,
    executablePath: process.execPath,
    assemblyName: "fake",
    manifest: {},
  };
  const fakeSpawner = (mode) => (_file, _args, options) => spawn(process.execPath, [fakePath, mode], options);

  await assert.rejects(
    startOfficeKitNativeClient({ descriptor: fakeDescriptor, spawnProcess: fakeSpawner("bad-handshake") }),
    (error) => error?.code === "runtime_protocol_mismatch",
  );

  const echo = await startOfficeKitNativeClient({ descriptor: fakeDescriptor, spawnProcess: fakeSpawner("echo") });
  const [first, second] = await Promise.all([
    echo.invoke(Uint8Array.from([1, 2, 3])),
    echo.invoke(Uint8Array.from([4, 5, 6, 7])),
  ]);
  assert.deepEqual([...first], [1, 2, 3]);
  assert.deepEqual([...second], [4, 5, 6, 7]);
  echo.kill();

  const truncated = await startOfficeKitNativeClient({ descriptor: fakeDescriptor, spawnProcess: fakeSpawner("truncated") });
  await assert.rejects(truncated.invoke(Uint8Array.of(1)), (error) => error?.code === "runtime_terminated");
  truncated.kill();

  const crashed = await startOfficeKitNativeClient({ descriptor: fakeDescriptor, spawnProcess: fakeSpawner("crash") });
  await assert.rejects(crashed.invoke(Uint8Array.of(1)), (error) => error?.code === "runtime_terminated");
  crashed.kill();
} finally {
  await rm(temporary, { recursive: true, force: true });
}

console.log("OfficeKit NativeAOT transport and integrity ok");
