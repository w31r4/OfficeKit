import {
  ArtifactFamily,
  CodecOperation,
} from "../generated/office_kit/artifact/v1/office_artifact_pb.js";
import { FileBlob } from "../shared/file-blob.mjs";
import { presentationEnvelope, presentationFromEnvelope } from "./office-kit-presentation.mjs";
import {
  assertCodecOptions,
  codecLimits,
  inputBytes,
  invokeOfficeKit,
  OFFICE_KIT_PROTOCOL_VERSION,
} from "./office-kit-runtime.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

export async function exportPptxWithOfficeKit(presentation, options = {}) {
  assertCodecOptions(options, new Set(["limits"]), "exportPptxWithOfficeKit");
  const response = await invokeOfficeKit({
    protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
    operation: CodecOperation.EXPORT_PPTX,
    family: ArtifactFamily.PRESENTATION,
    artifact: presentationEnvelope(presentation, OFFICE_KIT_PROTOCOL_VERSION),
    limits: codecLimits(options.limits),
  });
  return new FileBlob(response.file, {
    type: PPTX_MIME,
    metadata: { artifactKind: "presentation", codec: "office-kit", diagnostics: response.diagnostics },
  });
}

export async function importPptxWithOfficeKit(input, options = {}) {
  assertCodecOptions(options, new Set(["limits"]), "importPptxWithOfficeKit");
  const response = await invokeOfficeKit({
    protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
    operation: CodecOperation.IMPORT_PPTX,
    family: ArtifactFamily.PRESENTATION,
    file: await inputBytes(input),
    limits: codecLimits(options.limits),
  });
  return presentationFromEnvelope(response.artifact);
}
