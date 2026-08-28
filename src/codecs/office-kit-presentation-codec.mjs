import {
  ArtifactFamily,
  CodecOperation,
} from "../generated/office_kit/artifact/v1/office_artifact_pb.js";
import { createHash } from "node:crypto";
import { FileBlob } from "../shared/file-blob.mjs";
import { OfficeKitCodecError } from "./office-kit-error.mjs";
import { compilePresentationEditPlan, presentationEnvelope, presentationFromEnvelope, presentationRequiresNativeLeafEditPlan } from "./office-kit-presentation.mjs";
import {
  assertCodecOptions,
  codecLimits,
  inputBytes,
  invokeOfficeKit,
  OFFICE_KIT_PROTOCOL_VERSION,
} from "./office-kit-runtime.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const PPTX_IMAGE_EXTENSIONS_BY_CONTENT_TYPE = new Map([
  ["image/png", new Set(["png"])],
  ["image/jpeg", new Set(["jpg", "jpeg"])],
  ["image/gif", new Set(["gif"])],
  ["image/svg+xml", new Set(["svg"])],
]);

function restoreThinPresentationImport(artifact, sourceBytes) {
  const snapshot = artifact?.opaqueOpc?.sourcePackage;
  if (snapshot?.data?.length) return undefined;
  const sourceHash = createHash("sha256").update(sourceBytes).digest("hex");
  const snapshotHash = String(snapshot?.sha256 || "").toLowerCase();
  const identityHash = String(artifact?.source?.packageSha256 || "").toLowerCase();
  if (!snapshot || snapshotHash !== sourceHash || identityHash !== sourceHash) {
    throw new OfficeKitCodecError("OfficeKit thin PPTX response does not match the exact request source package.", [], { code: "source_package_hash_mismatch" });
  }
  snapshot.data = sourceBytes;
}

function presentationSlideRelationshipPartPath(slidePartPath) {
  const match = String(slidePartPath).match(/^ppt\/slides\/(slide[1-9][0-9]*[.]xml)$/u);
  return match ? `ppt/slides/_rels/${match[1]}.rels` : undefined;
}

function presentationImageEditPlanMetadata(operation, editPlan, result) {
  const replacement = operation.imageReplacement;
  const assets = editPlan.wire.assets.filter((asset) => asset.id === replacement?.assetId);
  const asset = assets.length === 1 ? assets[0] : undefined;
  const relationshipPartPath = presentationSlideRelationshipPartPath(operation.slidePartPath);
  const contentType = String(asset?.contentType || "").toLowerCase();
  const extensions = PPTX_IMAGE_EXTENSIONS_BY_CONTENT_TYPE.get(contentType);
  const sha256 = String(asset?.sha256 || "").toLowerCase();
  const bytes = asset?.data;
  if (!replacement || replacement.assetId !== operation.value || !asset || !relationshipPartPath ||
      !(bytes instanceof Uint8Array) || !/^[0-9a-f]{64}$/u.test(sha256) ||
      asset.id !== `asset/presentation/picture-bullet/${sha256}` ||
      createHash("sha256").update(bytes).digest("hex") !== sha256 || !extensions ||
      !result.changedParts.includes(relationshipPartPath)) {
    throw new OfficeKitCodecError(`OfficeKit PPTX Edit Plan result changed the image binding for operation ${operation.operationId}.`, [], { code: "presentation_edit_plan_result_mismatch" });
  }
  const mediaPrefix = `ppt/media/office-kit-${sha256.slice(0, 24)}.`;
  const mediaParts = result.changedParts.filter((partPath) => partPath.startsWith(mediaPrefix));
  const mediaPartPath = mediaParts.length === 1 ? mediaParts[0] : undefined;
  const mediaExtension = mediaPartPath?.slice(mediaPrefix.length);
  if (mediaParts.length > 1 || (mediaExtension && !extensions.has(mediaExtension))) {
    throw new OfficeKitCodecError(`OfficeKit PPTX Edit Plan result changed the image package footprint for operation ${operation.operationId}.`, [], { code: "presentation_edit_plan_result_mismatch" });
  }
  const crop = replacement.crop;
  return {
    assetId: asset.id,
    sha256,
    contentType,
    byteLength: bytes.byteLength,
    relationshipPartPath,
    mediaPartPath: mediaPartPath ?? null,
    crop: crop ? {
      leftThousandthPercent: crop.leftThousandthPercent,
      topThousandthPercent: crop.topThousandthPercent,
      rightThousandthPercent: crop.rightThousandthPercent,
      bottomThousandthPercent: crop.bottomThousandthPercent,
    } : null,
  };
}

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
          footprint.textLeafIndex !== operation.textLeafIndex || footprint.leafKind !== operation.leafKind ||
          footprint.mutationPartPath !== (operation.targetPartPath || operation.slidePartPath) ||
          JSON.stringify(footprint.shapeTreePath) !== JSON.stringify(operation.shapeTreePath)) {
        throw new OfficeKitCodecError(`OfficeKit PPTX Edit Plan result changed the binding for operation ${operation.operationId}.`, [], { code: "presentation_edit_plan_result_mismatch" });
      }
      const nestedFootprints = footprint.nestedFootprints || [];
      if (operation.leafKind === "chartDataValue") {
        if (nestedFootprints.length !== 1 || nestedFootprints[0].containerPartPath !== operation.embeddedPackagePartPath ||
            nestedFootprints[0].partPath !== operation.embeddedWorksheetPartPath) {
          throw new OfficeKitCodecError(`OfficeKit PPTX Edit Plan result changed the embedded-workbook binding for operation ${operation.operationId}.`, [], { code: "presentation_edit_plan_result_mismatch" });
        }
      } else if (nestedFootprints.length !== 0) {
        throw new OfficeKitCodecError(`OfficeKit PPTX Edit Plan result added an undeclared nested footprint to operation ${operation.operationId}.`, [], { code: "presentation_edit_plan_result_mismatch" });
      }
      return {
        operationId: operation.operationId,
        slideId: operation.slideId,
        slidePartPath: operation.slidePartPath,
        expectedSlideSha256: operation.expectedSlideSha256,
        targetId: operation.targetId,
        shapeTreeIndex: operation.shapeTreeIndex,
        shapeTreePath: [...operation.shapeTreePath],
        leafKind: operation.leafKind,
        expectedElementSha256: operation.expectedElementSha256,
        expectedSemanticSha256: operation.expectedSemanticSha256,
        textLeafIndex: operation.textLeafIndex,
        expectedTextSha256: operation.expectedTextSha256,
        expectedValue: operation.expectedValue,
        value: operation.value,
        ...(operation.targetPartPath ? {
          targetPartPath: operation.targetPartPath,
          expectedTargetPartSha256: operation.expectedTargetPartSha256,
          relationshipId: operation.relationshipId,
        } : {}),
        ...(operation.embeddedPackagePartPath ? {
          embeddedPackagePartPath: operation.embeddedPackagePartPath,
          expectedEmbeddedPackageSha256: operation.expectedEmbeddedPackageSha256,
          embeddedPackageRelationshipId: operation.embeddedPackageRelationshipId,
          embeddedWorksheetPartPath: operation.embeddedWorksheetPartPath,
          expectedEmbeddedWorksheetSha256: operation.expectedEmbeddedWorksheetSha256,
          embeddedCellReference: operation.embeddedCellReference,
          chartSeriesIndex: operation.chartSeriesIndex,
          chartPointIndex: operation.chartPointIndex,
          chartFormula: operation.chartFormula,
        } : {}),
        ...(operation.leafKind === "diagramText" ? {
          diagramModelId: operation.diagramModelId,
          diagramRunIndex: operation.diagramRunIndex,
        } : {}),
        ...(operation.leafKind === "imageAsset" ? {
          imageReplacement: presentationImageEditPlanMetadata(operation, editPlan, result),
        } : {}),
        ...(operation.leafKind === "deleteElement" ? {
          elementDeletion: { expectedNativeId: operation.elementDeletion?.expectedNativeId },
        } : {}),
        footprint: {
          mutationPartPath: footprint.mutationPartPath,
          sourceElementSha256: footprint.sourceElementSha256,
          outputElementSha256: footprint.outputElementSha256,
          oldValueSha256: footprint.oldValueSha256,
          newValueSha256: footprint.newValueSha256,
          sourceStartOffset: String(footprint.sourceStartOffset),
          sourceEndOffset: String(footprint.sourceEndOffset),
          outputEndOffset: String(footprint.outputEndOffset),
          shapeTreePath: [...footprint.shapeTreePath],
          leafKind: footprint.leafKind,
          nestedFootprints: nestedFootprints.map((nested) => ({
            containerPartPath: nested.containerPartPath,
            partPath: nested.partPath,
            oldValueSha256: nested.oldValueSha256,
            newValueSha256: nested.newValueSha256,
            sourceStartOffset: String(nested.sourceStartOffset),
            sourceEndOffset: String(nested.sourceEndOffset),
            outputEndOffset: String(nested.outputEndOffset),
          })),
        },
      };
    }),
  };
}

export async function exportPptxWithOfficeKit(presentation, options = {}) {
  assertCodecOptions(options, new Set(["limits"]), "exportPptxWithOfficeKit");
  const editPlan = compilePresentationEditPlan(presentation, OFFICE_KIT_PROTOCOL_VERSION);
  if (!editPlan && presentationRequiresNativeLeafEditPlan(presentation)) {
    throw new OfficeKitCodecError("Presentation native-leaf edits must compile to a bounded source-package Edit Plan; dependent or unsupported changes fail closed.", [], { code: "unsupported_presentation_native_leaf_edit" });
  }
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
    }, { fileSidecar: true });
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
  const sourceBytes = Uint8Array.from(await inputBytes(input));
  const response = await invokeOfficeKit({
    protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
    operation: CodecOperation.IMPORT_PPTX,
    family: ArtifactFamily.PRESENTATION,
    file: sourceBytes,
    limits: codecLimits(options.limits),
    thinPresentationImportResponse: true,
  }, { fileSidecar: true });
  restoreThinPresentationImport(response.artifact, sourceBytes);
  return presentationFromEnvelope(response.artifact);
}
