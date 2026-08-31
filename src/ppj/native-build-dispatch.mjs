import process from "node:process";

import { loadOfficeKitNativeDescriptor } from "../codecs/office-kit-native-client.mjs";

export async function execNativePpjBuild(
  argv,
  {
    cwd = process.cwd(),
    execve = process.execve,
    loadDescriptor = loadOfficeKitNativeDescriptor,
    platform = process.platform,
    env = process.env,
  } = {},
) {
  if (!eligible(argv, platform, execve)) return false;
  const descriptor = await loadDescriptor({ profile: "ppj", requiredCapability: "directBuild" });
  if (!descriptor) return false;
  if (descriptor.manifest?.profiles?.ppj?.directBuild !== true) return false;
  execve(descriptor.executablePath, [
    descriptor.executablePath,
    "--build",
    ...argv.slice(2),
    "--cwd",
    cwd,
  ], {
    ...env,
    DOTNET_GCConserveMemory: "9",
  });
  throw new Error("OfficeKit native PPJ build returned without replacing the Node process.");
}

function eligible(argv, platform, execve) {
  if (!Array.isArray(argv) || argv[0] !== "ppj" || argv[1] !== "build") return false;
  if (platform === "win32" || typeof execve !== "function") return false;
  const args = argv.slice(2);
  return !args.some((argument) =>
    argument === "--task" || argument.startsWith("--task=") ||
    argument === "--help" || argument === "-h" ||
    argument === "--cwd" || argument.startsWith("--cwd="));
}
