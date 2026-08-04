import os from "node:os";
import path from "node:path";

import {
  appendAuditRecord,
  initializeExcelConfiguration,
  POWERPOINT_ADDIN_ID,
  readExcelConfiguration,
  removeExcelState,
  resolveExcelStatePaths,
  updateExcelConfiguration,
} from "../excel-live/state.mjs";

export const POWERPOINT_BRIDGE_PORT = 47214;
export { POWERPOINT_ADDIN_ID } from "../excel-live/state.mjs";

export function resolvePowerPointStatePaths({ env = process.env, home = os.homedir() } = {}) {
  const configuredHome = env.OFFICEKIT_POWERPOINT_HOME;
  const root = path.resolve(
    configuredHome && configuredHome.length > 0
      ? configuredHome
      : path.join(env.OFFICE_KIT_HOME || path.join(home, ".office-kit"), "powerpoint"),
  );
  const shared = resolveExcelStatePaths({ env: { ...env, OFFICEKIT_EXCEL_HOME: root }, home });
  return Object.freeze({
    ...shared,
    manifest: path.join(root, "officekit-powerpoint-manifest.xml"),
  });
}

export {
  appendAuditRecord,
  readExcelConfiguration as readPowerPointConfiguration,
  removeExcelState as removePowerPointState,
  updateExcelConfiguration as updatePowerPointConfiguration,
};

export async function initializePowerPointConfiguration(paths, { port = POWERPOINT_BRIDGE_PORT } = {}) {
  const state = await initializeExcelConfiguration(paths, { port });
  if (state.config.addinId === POWERPOINT_ADDIN_ID && state.config.application === "powerpoint") return state;
  const updated = await updateExcelConfiguration(paths, (config) => ({
    ...config,
    addinId: POWERPOINT_ADDIN_ID,
    application: "powerpoint",
  }));
  return { config: updated.config, secret: updated.secret };
}
