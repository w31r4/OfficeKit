import { ExcelBridge, startExcelBridge } from "../excel-live/bridge.mjs";
import { createPowerPointLiveAdapter } from "./adapters/powerpoint.mjs";

/**
 * Shared HTTPS/session transport. The historical ExcelBridge name remains a
 * compatibility export; new host adapters use the same queue, pairing,
 * idempotency, timeout, and audit implementation.
 */
export class LiveBridge extends ExcelBridge {}

export async function startLiveBridge(options = {}) {
  return startExcelBridge(options);
}

export async function startPowerPointBridge({ adapter = createPowerPointLiveAdapter(), ...options } = {}) {
  return startExcelBridge({ ...options, adapter });
}
