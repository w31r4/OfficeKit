import assert from "node:assert/strict";
import { createHash } from "node:crypto";

import { presentationEnvelope } from "../src/codecs/office-kit-presentation.mjs";
import { Presentation } from "../src/presentation/index.mjs";

const sourceDataUrl = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACXBIWXMAAAPoAAAD6AG1e1JrAAAADUlEQVR4nGNgYGBgAAAABQABpfZFQAAAAABJRU5ErkJggg==";
const replacementDataUrl = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2nGQAAAAASUVORK5CYII=";
const sourceBytes = Buffer.from(sourceDataUrl.split(",")[1], "base64");
const sourceSha256 = createHash("sha256").update(sourceBytes).digest("hex");
const sourceAsset = {
  id: `asset/presentation/picture-bullet/${sourceSha256}`,
  fileName: "source.png",
  contentType: "image/png",
  data: sourceBytes,
  sha256: sourceSha256,
};

let resolutions = 0;
const presentation = Presentation.create();
const slide = presentation.slides.add({ name: "Lazy image" });
const importedDataUrlSource = Object.freeze({
  asset: sourceAsset,
  resolve() {
    resolutions += 1;
    return sourceDataUrl;
  },
});
const image = slide.images.add({
  name: "Imported image",
  position: { left: 0, top: 0, width: 100, height: 100 },
  fit: "stretch",
  _officeKitDataUrlSource: importedDataUrlSource,
});
const group = slide.groups.add({
  name: "Imported group",
  position: { left: 120, top: 0, width: 100, height: 100 },
  childFrame: { left: 0, top: 0, width: 100, height: 100 },
  children: [{
    kind: "image",
    name: "Imported grouped image",
    position: { left: 0, top: 0, width: 100, height: 100 },
    fit: "stretch",
    _officeKitDataUrlSource: importedDataUrlSource,
  }],
});
const groupedImage = group.images.items[0];

const unchanged = presentationEnvelope(presentation, 2);
assert.equal(resolutions, 0, "unchanged export must reuse imported bytes without materializing base64");
assert.equal(unchanged.assets.length, 1);
assert.equal(unchanged.assets[0].sha256, sourceSha256);
assert.equal(image.dataUrl, sourceDataUrl);
assert.equal(image.dataUrl, sourceDataUrl);
assert.equal(groupedImage.dataUrl, sourceDataUrl);
assert.equal(groupedImage.dataUrl, sourceDataUrl);
assert.equal(resolutions, 2, "each public dataUrl getter must materialize at most once");

image.dataUrl = replacementDataUrl;
const changed = presentationEnvelope(presentation, 2);
assert.equal(changed.assets.length, 2);
assert.ok(changed.assets.some((asset) => asset.sha256 === sourceSha256));
assert.ok(changed.assets.some((asset) => asset.sha256 !== sourceSha256));
assert.equal(image.dataUrl, replacementDataUrl);

process.stdout.write("presentation lazy image assets: ok\n");
