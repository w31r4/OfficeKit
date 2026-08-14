import {
  ArtifactFamily,
  CodecOperation,
} from "../generated/office_kit/artifact/v1/office_artifact_pb.js";
import { createHash } from "node:crypto";
import { FileBlob } from "../shared/file-blob.mjs";
import { OfficeKitCodecError } from "./office-kit-error.mjs";
import { compilePresentationEditPlan, presentationEnvelope, presentationFromEnvelope } from "./office-kit-presentation.mjs";
import {
  assertCodecOptions,
  codecLimits,
  inputBytes,
  invokeOfficeKit,
  OFFICE_KIT_PROTOCOL_VERSION,
} from "./office-kit-runtime.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

function presentationEditPlanMetadata(editPlan, result) {
  const footprintById = new Map((result.operations || []).map((operation) => [operation.operationId, operation]));
  return {
    schema: editPlan.schema,
    sourceRevisionSha256: editPlan.sourceRevisionSha256,
    outputSha256: result.outputSha256,
    changedParts: [...result.changedParts],
    operations: editPlan.operations.map((operation) => {
      const footprint = footprintById.get(operation.operationId);
      if (!footprint) throw new OfficeKitCodecError(`OfficeKit PPTX Edit Plan result omitted operation ${operation.operationId}.`, [], { code: "presentation_edit_plan_result_mismatch" });
      if (footprint.slideId !== operation.slideId || footprint.slidePartPath !== operation.slidePartPath ||
          footprint.targetId !== operation.targetId || footprint.shapeTreeIndex !== operation.shapeTreeIndex ||
          footprint.textLeafIndex !== operation.textLeafIndex ||
          JSON.stringify(footprint.shapeTreePath) !== JSON.stringify(operation.shapeTreePath)) {
        throw new OfficeKitCodecError(`OfficeKit PPTX Edit Plan result changed the binding for operation ${operation.operationId}.`, [], { code: "presentation_edit_plan_result_mismatch" });
      }
      return {
        operationId: operation.operationId,
        slideId: operation.slideId,
        slidePartPath: operation.slidePartPath,
        expectedSlideSha256: operation.expectedSlideSha256,
        targetId: operation.targetId,
        shapeTreeIndex: operation.shapeTreeIndex,
        shapeTreePath: [...operation.shapeTreePath],
        expectedElementSha256: operation.expectedElementSha256,
        expectedSemanticSha256: operation.expectedSemanticSha256,
        textLeafIndex: operation.textLeafIndex,
        expectedTextSha256: operation.expectedTextSha256,
        expectedValue: operation.expectedValue,
        value: operation.value,
        footprint: {
          sourceElementSha256: footprint.sourceElementSha256,
          outputElementSha256: footprint.outputElementSha256,
          oldValueSha256: footprint.oldValueSha256,
          newValueSha256: footprint.newValueSha256,
          sourceStartOffset: String(footprint.sourceStartOffset),
          sourceEndOffset: String(footprint.sourceEndOffset),
          outputEndOffset: String(footprint.outputEndOffset),
          shapeTreePath: [...footprint.shapeTreePath],
        },
      };
    }),
  };
}

export async function exportPptxWithOfficeKit(presentation, options = {}) {
  assertCodecOptions(options, new Set(["limits"]), "exportPptxWithOfficeKit");
  const editPlan = compilePresentationEditPlan(presentation, OFFICE_KIT_PROTOCOL_VERSION);
  if (editPlan) {
    if (editPlan.operations.length === 0) {
      return new FileBlob(editPlan.sourceBytes, {
        type: PPTX_MIME,
        metadata: {
          artifactKind: "presentation",
          codec: "office-kit",
          editPlan: { schema: editPlan.schema, sourceRevisionSha256: editPlan.sourceRevisionSha256, operations: [], changedParts: [] },
        },
      });
    }
    const response = await invokeOfficeKit({
      protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
      operation: CodecOperation.APPLY_PPTX_EDIT_PLAN,
      family: ArtifactFamily.PRESENTATION,
      file: editPlan.sourceBytes,
      presentationEditPlan: editPlan.wire,
      limits: codecLimits(options.limits),
    });
    const result = response.presentationEditPlan;
    const outputSha256 = createHash("sha256").update(response.file).digest("hex");
    if (!result || result.sourceSha256 !== editPlan.sourceRevisionSha256 || result.outputSha256 !== outputSha256 ||
        result.operations.length !== editPlan.operations.length || result.changedParts.length === 0) {
      throw new OfficeKitCodecError("OfficeKit PPTX Edit Plan result did not match the compiled request and output bytes.", response.diagnostics, { code: "presentation_edit_plan_result_mismatch" });
    }
    return new FileBlob(response.file, {
      type: PPTX_MIME,
      metadata: {
        artifactKind: "presentation",
        codec: "office-kit",
        diagnostics: response.diagnostics,
        editPlan: presentationEditPlanMetadata(editPlan, result),
      },
    });
  }
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
