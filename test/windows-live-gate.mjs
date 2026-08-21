import assert from "node:assert/strict";
import { validateWindowsLiveEvidence } from "../scripts/validate-windows-live-evidence.mjs";

const evidence = {
  schema: "office-kit.windows-live-evidence.v1",
  method: "human-observed-windows-office",
  checkedAt: "2026-08-04T12:00:00Z",
  commit: "0123456789abcdef0123456789abcdef01234567",
  host: {
    platform: "win32-x64",
    office: {
      excel: { installed: true, version: "Microsoft Excel 16.0" },
      powerpoint: { installed: true, version: "Microsoft PowerPoint 16.0" },
    },
  },
  workflows: [
    {
      name: "excel-live",
      result: "passed",
      observedAt: "2026-08-04T12:01:00Z",
      automationSource: "windows-office",
      checks: {
        manifestUploaded: true,
        paired: true,
        twoWorkbooksIsolated: true,
        unsavedReadWrite: true,
        explicitSave: true,
        disconnectReconnect: true,
        sourceProtected: true,
        bridgeIdleExit: true,
      },
    },
    {
      name: "powerpoint-live",
      result: "passed",
      observedAt: "2026-08-04T12:02:00Z",
      automationSource: "windows-office",
      checks: {
        manifestUploaded: true,
        paired: true,
        twoPresentationsIsolated: true,
        unsavedReadWrite: true,
        selectionRead: true,
        slideImageReviewed: true,
        explicitSave: true,
        disconnectReconnect: true,
        unsupportedCapabilityFailClosed: true,
        sourceProtected: true,
        bridgeIdleExit: true,
      },
    },
  ],
};

assert.deepEqual(validateWindowsLiveEvidence(evidence), {
  schema: "office-kit.windows-live-evidence.v1",
  checkedAt: "2026-08-04T12:00:00Z",
  commit: "0123456789abcdef0123456789abcdef01234567",
  platform: "win32-x64",
  excelVersion: "Microsoft Excel 16.0",
  powerpointVersion: "Microsoft PowerPoint 16.0",
  workflows: ["excel-live", "powerpoint-live"],
});
assert.equal(validateWindowsLiveEvidence(evidence, { expectedCommit: evidence.commit }).commit, evidence.commit);
assert.throws(
  () => validateWindowsLiveEvidence(evidence, { expectedCommit: "fedcba9876543210fedcba9876543210fedcba98" }),
  /checked-out commit/,
);

for (const mutation of [
  (value) => { value.method = "mock"; },
  (value) => { value.host.platform = "darwin-arm64"; },
  (value) => { value.host.office.excel.installed = false; },
  (value) => { value.workflows[0].automationSource = "mock"; },
  (value) => { value.workflows[1].result = "skipped"; },
  (value) => { value.workflows[1].checks.slideImageReviewed = false; },
]) {
  const invalid = structuredClone(evidence);
  mutation(invalid);
  assert.throws(() => validateWindowsLiveEvidence(invalid), /evidence|Windows|Excel|mock|passed|human/i);
}

console.log("Windows Office live evidence gate ok");
