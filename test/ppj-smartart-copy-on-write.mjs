import assert from "node:assert/strict";

import { createHash } from "node:crypto";

import { copyOnWriteSmartArtDefinition } from "../src/ppj/smartart-definition.mjs";

const sha256 = (value) => createHash("sha256").update(value).digest("hex");

const original = Buffer.from(JSON.stringify({
  schema: "office-kit/smartart-definition/v1",
  layout: { id: "shared-process", profile: "process" },
  style: { id: "basic" },
  colors: { id: "accent" },
}));
const originalHash = sha256(original);
const workspace = {
  root: {
    schema: "office-kit/ppj/v1",
    assets: [{
      id: "shared-definition",
      uri: `deck.assets/smartart/${originalHash}.json`,
      mimeType: "application/vnd.officekit.smartart-definition+json",
      sha256: originalHash,
      rights: { status: "internal" },
      accessibility: { decorative: true },
    }],
    pages: [{
      id: "page-1",
      elements: [
        { id: "diagram-a", type: "smartArt", mode: "authored", definitionAsset: "shared-definition", nodes: [] },
        { id: "diagram-b", type: "smartArt", mode: "authored", definitionAsset: "shared-definition", nodes: [] },
      ],
    }],
  },
};
const edited = Buffer.from(JSON.stringify({
  schema: "office-kit/smartart-definition/v1",
  layout: {
    id: "two-column-process",
    profile: "process",
    operators: [{ id: "columns", kind: "rule", arguments: { columns: 2 } }],
  },
  style: { id: "basic" },
  colors: { id: "accent" },
}));

const revision = copyOnWriteSmartArtDefinition(workspace, { elementId: "diagram-b", definition: edited });
assert.equal(workspace.root.pages[0].elements[1].definitionAsset, "shared-definition", "the input workspace stays immutable");
assert.equal(revision.root.pages[0].elements[0].definitionAsset, "shared-definition", "the unselected instance keeps the shared definition");
assert.equal(revision.root.pages[0].elements[1].definitionAsset, revision.definitionAssetId, "only the selected instance is repointed");
assert.notEqual(revision.definitionAssetId, "shared-definition");
assert.equal(revision.root.assets.length, 2);
assert.equal(revision.asset.sha256, sha256(edited));
assert.equal(revision.asset.reused, false);

const reused = copyOnWriteSmartArtDefinition(workspace, { elementId: "diagram-a", definition: original });
assert.equal(reused.definitionAssetId, "shared-definition", "identical content reuses the existing immutable asset");
assert.equal(reused.root.assets.length, 1);
assert.equal(reused.asset.reused, true);
