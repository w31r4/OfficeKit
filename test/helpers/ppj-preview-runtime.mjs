// Test-only preload: select an independently built real codec, including in
// CLI child processes. No response mocking, installed-package replacement or
// production environment-variable fallback is involved.
import path from "node:path";
import { registerHooks } from "node:module";
import { startOfficeKitNativeClient } from "../../src/codecs/office-kit-native-client.mjs";

if (process.env.OFFICEKIT_PREVIEW_TEST_RUNTIME) {
  const packageJsonPath = path.resolve(process.env.OFFICEKIT_PREVIEW_TEST_RUNTIME, "package.json");
  globalThis[Symbol.for("officekit.preview.entry.test")] = options => startOfficeKitNativeClient({ ...options, packageJsonPath });
  const module = `export const OFFICE_KIT_NATIVE_TRANSPORT_VERSION=2;
    export const startOfficeKitNativeClient=options=>globalThis[Symbol.for("officekit.preview.entry.test")](options);`;
  registerHooks({ resolve(specifier, context, next) {
    if (specifier === "./office-kit-native-client.mjs" && context.parentURL?.endsWith("/src/codecs/office-kit-runtime.mjs"))
      return { url: `data:text/javascript,${encodeURIComponent(module)}`, shortCircuit: true };
    return next(specifier, context);
  } });
}
