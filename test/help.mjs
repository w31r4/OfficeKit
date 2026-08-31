import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import * as rootApi from "office-kit";

import { PUBLIC_HELP_CATALOG } from "../src/help/index.mjs";

for (const name of ["Presentation", "PresentationFile", "Slide", "Shape"]) {
  assert.equal(name in rootApi, false, `root must not expose legacy Presentation binding ${name}`);
}
const presentationHelp = PUBLIC_HELP_CATALOG.filter((entry) => entry.artifactKind === "presentation");
assert.ok(presentationHelp.length > 0);
assert.ok(presentationHelp.every((entry) => entry.name.startsWith("officekit ppj ")));
assert.match(rootApi.helpArtifact("presentation", "ppj build").ndjson, /source-bound PPJ diff/);
assert.equal(rootApi.helpArtifact("presentation", "PresentationFile.importPptx").ndjson, "");

const apiDocs = await readFile(new URL("../docs/api.md", import.meta.url), "utf8");
assert.match(apiDocs, /#### `officekit ppj build`/);
assert.doesNotMatch(apiDocs, /#### `PresentationFile\.importPptx`/);

console.log("public Help surface smoke ok");
