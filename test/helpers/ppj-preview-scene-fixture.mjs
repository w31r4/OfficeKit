import { create, toBinary } from "@bufbuild/protobuf";
import { PresentationPreviewSceneSchema, PresentationElementSchema } from "../../src/generated/office_kit/artifact/v1/office_artifact_pb.js";
import { sha256 } from "../../src/ppj/workspace.mjs";

export const emuFrame = (x = 10, y = 40, width = 200, height = 100) => ({
  leftEmu: BigInt(x * 12700), topEmu: BigInt(y * 12700),
  widthEmu: BigInt(width * 12700), heightEmu: BigInt(height * 12700),
});
export const nativeElement = (id, kind, value) => create(PresentationElementSchema, { id, content: { case: kind, value } });

// Synthetic transport evidence only. Callers supply independent native state
// and every semantic owner explicitly; this is not a second PPJ compiler.
export function previewSceneFixture(program, pages, owners, assets = [], canvas = { width: 400, height: 240 }) {
  const programJson = Buffer.from(JSON.stringify(program)), file = Buffer.from("synthetic candidate, not a PPTX");
  const payloads = assets.map(asset => ({ ...asset, sha256: sha256(asset.data) }));
  const scene = create(PresentationPreviewSceneSchema, {
    version: 1, origin: 1, programSha256: sha256(programJson), candidateSha256: sha256(file),
    presentation: { slideWidthEmu: BigInt(canvas.width * 12700), slideHeightEmu: BigInt(canvas.height * 12700), slides: pages },
    assets: payloads.map(asset => ({ nativeId: asset.id, contentType: asset.mimeType, sha256: asset.sha256 })),
  });
  const bindingSchema = PresentationPreviewSceneSchema.fields.find(field => field.localName === "bindings").message;
  let ownerIndex = 0;
  function bind(elements, prefix, pageId) {
    elements.forEach((element, zOrder) => {
      const scenePath = `${prefix}[${zOrder}]`, programPath = owners[ownerIndex++];
      if (!programPath) throw new Error("Fixture needs an explicit owner for every native node");
      scene.bindings.push(create(bindingSchema, { pageId, semanticId: element.id, nativeId: element.id,
        scenePath, programPath, attribution: 1, zOrder }));
      if (element.content.case === "group") bind(element.content.value.children, `${scenePath}.group.children`, pageId);
    });
  }
  scene.presentation.slides.forEach((page, index) => bind(page.elements, `$.presentation.slides[${index}].elements`, page.id));
  if (ownerIndex !== owners.length) throw new Error("Unused fixture owner");
  scene.sha256 = sha256(toBinary(PresentationPreviewSceneSchema, scene));
  return { programJson, programSha256: scene.programSha256, file, outputSha256: scene.candidateSha256,
    sourceBound: false, assets: payloads, previewScene: scene };
}
