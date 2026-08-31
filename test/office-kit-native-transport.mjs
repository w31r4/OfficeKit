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
import { execNativePpjBuild } from "../src/ppj/native-build-dispatch.mjs";

let execveCall;
await assert.rejects(
  execNativePpjBuild(["ppj", "build", "deck.ppj", "-o", "deck.pptx", "--json"], {
    cwd: "/workspace",
    platform: "darwin",
    env: { OFFICEKIT_TEST: "1" },
    loadDescriptor: async (options) => {
      assert.deepEqual(options, { profile: "ppj", requiredCapability: "directBuild" });
      return {
        executablePath: "/codec",
        manifest: { profiles: { ppj: { directBuild: true } } },
      };
    },
    execve: (...args) => {
      execveCall = args;
      throw new Error("execve sentinel");
    },
  }),
  /execve sentinel/u,
);
assert.deepEqual(execveCall[1], [
  "/codec", "--build", "deck.ppj", "-o", "deck.pptx", "--json", "--cwd", "/workspace",
]);
assert.equal(execveCall[2].DOTNET_GCConserveMemory, "9");
assert.equal(await execNativePpjBuild(["ppj", "build", "deck.ppj", "-o", "deck.pptx"], {
  platform: "darwin",
  execve() { throw new Error("must not exec"); },
  loadDescriptor: async () => null,
}), false);
assert.equal(await execNativePpjBuild(["ppj", "build", "deck.ppj", "-o", "deck.pptx", "--task", "task-1"], {
  platform: "darwin",
  execve() { throw new Error("must not exec"); },
  loadDescriptor() { throw new Error("must not load"); },
}), false);

const target = officeKitNativeTarget();
assert.equal(target, `${process.platform}-${process.arch}`);
assert.throws(() => officeKitNativeTarget("freebsd", "x64"), (error) => error?.code === "runtime_unsupported_platform");

const installed = await loadOfficeKitNativeDescriptor();
const installedPpj = await loadOfficeKitNativeDescriptor({ profile: "ppj" });
assert.equal(installed.target, target);
assert.equal(installed.profile, "office");
assert.equal(installedPpj.profile, "ppj");
assert.equal(installedPpj.packageRoot, installed.packageRoot);
assert.equal(installed.manifest.schemaVersion, 2);
assert.equal(installed.manifest.backend, "native-aot");
assert.equal(installed.manifest.transportVersion, 2);
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

  const ppjExecutableName = process.platform === "win32" ? "officekit-ppj-codec.exe" : "officekit-ppj-codec";
  const tamperedPpjRoot = path.join(temporary, "tampered-ppj");
  await cp(installed.packageRoot, tamperedPpjRoot, { recursive: true });
  await appendFile(path.join(tamperedPpjRoot, "bin", ppjExecutableName), Buffer.from([0]));
  if (process.platform !== "win32") await chmod(path.join(tamperedPpjRoot, "bin", ppjExecutableName), 0o755);
  await assert.rejects(
    loadOfficeKitNativeDescriptor({ packageJsonPath: path.join(tamperedPpjRoot, "package.json"), profile: "ppj" }),
    (error) => error?.code === "runtime_integrity_failure",
  );

  const fakePath = path.join(temporary, "fake-codec.mjs");
  await writeFile(fakePath, `
import process from "node:process";
const mode = process.argv[2];
const handshake = Buffer.alloc(12);
handshake.write("OKIT", 0, "ascii");
handshake.writeUInt32BE(2, 4);
handshake.writeUInt32BE(2, 8);
if (mode === "bad-handshake") handshake.write("FAIL", 0, "ascii");
process.stdout.write(handshake);
if (mode === "bad-handshake") setTimeout(() => process.exit(0), 500);
let pending = Buffer.alloc(0);
process.stdin.on("data", (chunk) => {
  if (mode === "crash") process.exit(77);
  pending = Buffer.concat([pending, chunk]);
  while (pending.length >= 8) {
    const length = pending.readUInt32BE(0);
    const fileLength = pending.readUInt32BE(4);
    if (pending.length < 8 + length + fileLength) return;
    const request = pending.subarray(8, 8 + length);
    const file = pending.subarray(8 + length, 8 + length + fileLength);
    pending = pending.subarray(8 + length + fileLength);
    const prefix = Buffer.alloc(4);
    if (mode === "truncated") {
      prefix.writeUInt32BE(8, 0);
      process.stdout.write(prefix);
      process.stdout.write(Buffer.from([1, 2]));
      setTimeout(() => process.exit(78), 20);
      return;
    }
    prefix.writeUInt32BE(request.length + file.length, 0);
    process.stdout.write(prefix);
    process.stdout.write(request);
    process.stdout.write(file);
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
    echo.invoke(Uint8Array.from([4, 5]), Uint8Array.from([6, 7])),
  ]);
  assert.deepEqual([...first], [1, 2, 3]);
  assert.deepEqual([...second], [4, 5, 6, 7]);
  assert.equal(echo.idle, true);
  await echo.retire();
  assert.equal(echo.closed, true);

  const idle = await startOfficeKitNativeClient({
    descriptor: fakeDescriptor,
    spawnProcess: fakeSpawner("echo"),
    idleRetireMs: 10,
  });
  assert.deepEqual([...await idle.invoke(Uint8Array.of(9))], [9]);
  await Promise.race([
    idle.terminated,
    new Promise((_, reject) => setTimeout(() => reject(new Error("idle codec did not retire")), 500)),
  ]);
  assert.equal(idle.closed, true);

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
