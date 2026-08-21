#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const REQUIRED_WORKFLOWS = ["excel-live", "powerpoint-live"];
const REQUIRED_CHECKS = {
  "excel-live": [
    "manifestUploaded",
    "paired",
    "twoWorkbooksIsolated",
    "unsavedReadWrite",
    "explicitSave",
    "disconnectReconnect",
    "sourceProtected",
    "bridgeIdleExit",
  ],
  "powerpoint-live": [
    "manifestUploaded",
    "paired",
    "twoPresentationsIsolated",
    "unsavedReadWrite",
    "selectionRead",
    "slideImageReviewed",
    "explicitSave",
    "disconnectReconnect",
    "unsupportedCapabilityFailClosed",
    "sourceProtected",
    "bridgeIdleExit",
  ],
};

export function validateWindowsLiveEvidence(value, { expectedCommit = undefined } = {}) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("evidence must be a JSON object");
  if (value.schema !== "office-kit.windows-live-evidence.v1") throw new Error("unsupported Windows live evidence schema");
  if (value.method !== "human-observed-windows-office") throw new Error("evidence must come from a human-observed Windows Office host");
  if (!/^20\d\d-\d\d-\d\dT/.test(String(value.checkedAt || ""))) throw new Error("checkedAt must be an ISO timestamp");
  if (!/^win32-(x64|arm64)$/.test(String(value.host?.platform || ""))) throw new Error("host.platform must identify Windows");
  if (value.host?.office?.excel?.installed !== true || value.host?.office?.powerpoint?.installed !== true) {
    throw new Error("both Excel and PowerPoint must be installed on the observed host");
  }
  if (!String(value.host?.office?.excel?.version || "") || !String(value.host?.office?.powerpoint?.version || "")) {
    throw new Error("Excel and PowerPoint versions are required");
  }
  if (!/^[0-9a-f]{40}$/.test(String(value.commit || ""))) throw new Error("commit must be a 40-character SHA-1");
  if (expectedCommit && value.commit !== expectedCommit) throw new Error("evidence commit does not match the checked-out commit");
  if (!Array.isArray(value.workflows)) throw new Error("workflows must be an array");
  const byName = new Map(value.workflows.map((workflow) => [workflow?.name, workflow]));
  for (const name of REQUIRED_WORKFLOWS) {
    const workflow = byName.get(name);
    if (!workflow || workflow.result !== "passed") throw new Error(`${name} must have a passed human result`);
    if (workflow.automationSource === "mock" || workflow.automationSource === "macos") {
      throw new Error(`${name} cannot use mock or macOS evidence`);
    }
    if (!String(workflow.observedAt || "").startsWith(String(value.checkedAt).slice(0, 10))) {
      throw new Error(`${name}.observedAt must be on the evidence date`);
    }
    for (const check of REQUIRED_CHECKS[name]) {
      if (workflow.checks?.[check] !== true) {
        throw new Error(`${name}.checks.${check} must be true in human-observed evidence`);
      }
    }
  }
  return {
    schema: value.schema,
    checkedAt: value.checkedAt,
    commit: value.commit,
    platform: value.host.platform,
    excelVersion: value.host.office.excel.version,
    powerpointVersion: value.host.office.powerpoint.version,
    workflows: REQUIRED_WORKFLOWS,
  };
}

const entry = process.argv[1] ? path.resolve(process.argv[1]) : "";
if (entry === path.resolve(fileURLToPath(import.meta.url))) {
  const evidencePath = process.argv[2];
  const expectedCommit = process.argv[3];
  if (!evidencePath) {
    console.error("usage: node scripts/validate-windows-live-evidence.mjs <evidence.json>");
    process.exit(2);
  }
  try {
    const value = JSON.parse(await fs.readFile(path.resolve(evidencePath), "utf8"));
    console.log(JSON.stringify(validateWindowsLiveEvidence(value, { expectedCommit }), null, 2));
  } catch (error) {
    console.error(`Windows live evidence rejected: ${error.message}`);
    process.exit(1);
  }
}
