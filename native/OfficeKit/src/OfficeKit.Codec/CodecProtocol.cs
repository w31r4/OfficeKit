using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using System.Runtime.InteropServices;

namespace OfficeKit.Codec;

public static class CodecProtocol
{
    public const uint ProtocolVersion = CodecWireProtocol.ProtocolVersion;
    public const int AbsoluteRequestLimit = CodecWireProtocol.AbsoluteRequestLimit;

    public static byte[] Invoke(byte[] requestBytes)
    {
        return Invoke(ref requestBytes);
    }

    public static byte[] Invoke(ref byte[] requestBytes)
    {
        return InvokeResponse(ref requestBytes).ToByteArray();
    }

    public static CodecResponse InvokeResponse(ref byte[] requestBytes)
    {
        return InvokeResponse(ref requestBytes, null);
    }

    public static CodecResponse InvokeResponse(ref byte[] requestBytes, byte[]? requestFileBytes)
    {
        var response = new CodecResponse { ProtocolVersion = ProtocolVersion };
        try
        {
            if (requestBytes is null || requestBytes.Length == 0)
                throw new CodecException("empty_request", "Codec request bytes must not be empty.");
            if (requestBytes.Length > AbsoluteRequestLimit)
                throw new CodecException("request_budget_exceeded", $"Codec request exceeds the absolute {AbsoluteRequestLimit}-byte wire budget.");

            var request = CodecRequest.Parser.ParseFrom(requestBytes);
            requestBytes = [];
            if (requestFileBytes is { Length: > 0 })
            {
                if (!request.File.IsEmpty)
                    throw new CodecException("ambiguous_file_payload", "Codec request cannot contain both inline file bytes and a native transport file sidecar.");
                // NativeHost owns this array for the duration of InvokeResponse
                // and never mutates it. Keep the public protobuf request model
                // unchanged without copying the raw file through the parser.
                request.File = UnsafeByteOperations.UnsafeWrap(requestFileBytes);
            }
            ValidateRequest(request);
            var limits = EffectiveCodecLimits.From(request.Limits);
            switch (request.Operation)
            {
                case CodecOperation.ImportXlsx:
                {
                    var result = XlsxCodec.Import(RequestFileBytes(request.File), limits);
                    response.Artifact = result.Artifact;
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.ExportXlsx:
                {
                    var result = XlsxCodec.Export(request.Artifact, limits);
                    response.File = ByteString.CopyFrom(result.File);
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.ImportDocx:
                {
                    var result = DocxCodec.Import(RequestFileBytes(request.File), limits);
                    response.Artifact = result.Artifact;
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.ExportDocx:
                {
                    var result = DocxCodec.Export(request.Artifact, limits);
                    response.File = ByteString.CopyFrom(result.File);
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.FinalizeDocxRevisions:
                {
                    var result = DocxRevisionFinalizationCodec.Finalize(
                        RequestFileBytes(request.File),
                        request.RevisionFinalization,
                        limits);
                    response.File = ByteString.CopyFrom(result.File);
                    response.RevisionFinalization = result.Result;
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.AddDocxTrackedReplacement:
                {
                    var result = DocxTrackedReplacementCodec.Add(
                        RequestFileBytes(request.File),
                        request.TrackedReplacement,
                        limits);
                    response.File = ByteString.CopyFrom(result.File);
                    response.TrackedReplacement = result.Result;
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.ImportPptx:
                {
                    var result = PptxCodec.Import(RequestFileBytes(request.File), limits);
                    response.Artifact = result.Artifact;
                    if (request.ThinPresentationImportResponse)
                        ThinPresentationImportResponse(response.Artifact);
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.ExportPptx:
                {
                    var result = PptxCodec.Export(request.Artifact, limits);
                    response.File = ByteString.CopyFrom(result.File);
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.ApplyPptxEditPlan:
                {
                    var result = PptxEditPlanCodec.Apply(
                        RequestFileBytes(request.File),
                        request.PresentationEditPlan,
                        limits);
                    response.File = ByteString.CopyFrom(result.File);
                    response.PresentationEditPlan = result.Result;
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.ProjectPptxToPpj:
                {
                    var result = PpjPresentationProjector.Project(
                        RequestFileBytes(request.File),
                        request.PresentationProgram,
                        limits);
                    response.PresentationProgram = result.Program;
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.CompilePpjToPptx:
                {
                    var result = PpjPresentationCompiler.Compile(
                        request.PresentationProgram,
                        RequestFileBytes(request.File),
                        limits);
                    response.File = ByteString.CopyFrom(result.File);
                    response.PresentationProgram = result.Program;
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                default:
                    throw new CodecException("unsupported_operation", $"Codec operation {request.Operation} is not implemented.");
            }
            response.Ok = true;
        }
        catch (CodecException exception)
        {
            response.Diagnostics.Add(Error(exception.Code, exception.Message, exception.SourcePath));
        }
        catch (InvalidProtocolBufferException)
        {
            response.Diagnostics.Add(Error("invalid_wire_payload", "Codec request is not valid OfficeKit protobuf data."));
        }
        catch (Exception)
        {
            response.Diagnostics.Add(Error("codec_failure", "OpenXML codec failed while processing the request."));
        }
        return response;
    }

    private static byte[] RequestFileBytes(ByteString file)
    {
        if (MemoryMarshal.TryGetArray(file.Memory, out var segment) && segment.Array is not null &&
            segment.Offset == 0 && segment.Count == segment.Array.Length)
            return segment.Array;
        return file.ToByteArray();
    }

    private static void ValidateRequest(CodecRequest request)
    {
        if (request.ProtocolVersion != ProtocolVersion)
            throw new CodecException("unsupported_protocol_version", $"Protocol version {request.ProtocolVersion} is unsupported; expected {ProtocolVersion}.");
        var expectedFamily = request.Operation switch
        {
            CodecOperation.ImportXlsx or CodecOperation.ExportXlsx => ArtifactFamily.Workbook,
            CodecOperation.ImportDocx or CodecOperation.ExportDocx or CodecOperation.FinalizeDocxRevisions or CodecOperation.AddDocxTrackedReplacement => ArtifactFamily.Document,
            CodecOperation.ImportPptx or CodecOperation.ExportPptx or CodecOperation.ApplyPptxEditPlan or CodecOperation.ProjectPptxToPpj or CodecOperation.CompilePpjToPptx => ArtifactFamily.Presentation,
            _ => throw new CodecException("unsupported_operation", $"Codec operation {request.Operation} is not implemented."),
        };
        if (request.Family != expectedFamily)
            throw new CodecException("artifact_family_mismatch", $"Codec operation {request.Operation} requires artifact family {expectedFamily}, not {request.Family}.");
        if (request.Operation is CodecOperation.ImportXlsx or CodecOperation.ImportDocx or CodecOperation.ImportPptx or CodecOperation.FinalizeDocxRevisions or CodecOperation.AddDocxTrackedReplacement or CodecOperation.ApplyPptxEditPlan or CodecOperation.ProjectPptxToPpj && request.File.IsEmpty)
        {
            var message = request.Operation switch
            {
                CodecOperation.FinalizeDocxRevisions => "DOCX revision finalization requires non-empty file bytes.",
                CodecOperation.AddDocxTrackedReplacement => "DOCX tracked replacement requires non-empty file bytes.",
                CodecOperation.ApplyPptxEditPlan => "PPTX edit plan requires non-empty file bytes.",
                CodecOperation.ProjectPptxToPpj => "PPTX projection requires non-empty file bytes.",
                _ => $"{expectedFamily} import requires non-empty file bytes.",
            };
            throw new CodecException("empty_input", message);
        }
        if (request.Operation is CodecOperation.ExportXlsx or CodecOperation.ExportDocx or CodecOperation.ExportPptx && request.Artifact is null)
            throw new CodecException("missing_artifact", $"{expectedFamily} export requires an artifact envelope.");
        if (request.Operation == CodecOperation.FinalizeDocxRevisions && request.RevisionFinalization is null)
            throw new CodecException("missing_revision_finalization", "DOCX revision finalization requires revision_finalization options.");
        if (request.Operation == CodecOperation.AddDocxTrackedReplacement && request.TrackedReplacement is null)
            throw new CodecException("missing_tracked_replacement", "DOCX tracked replacement requires tracked_replacement options.");
        if (request.Operation == CodecOperation.ApplyPptxEditPlan && request.PresentationEditPlan is null)
            throw new CodecException("missing_presentation_edit_plan", "PPTX edit plan requires presentation_edit_plan options.");
        if (request.Operation is CodecOperation.ProjectPptxToPpj or CodecOperation.CompilePpjToPptx && request.PresentationProgram is null)
            throw new CodecException("missing_presentation_program", "PPJ operations require presentation_program options.");
        if (request.Operation == CodecOperation.CompilePpjToPptx && request.PresentationProgram.ProgramJson.IsEmpty)
            throw new CodecException("empty_ppj", "PPJ compilation requires non-empty program_json bytes.");
        if (request.ThinPresentationImportResponse && request.Operation != CodecOperation.ImportPptx)
            throw new CodecException("invalid_request", "thin_presentation_import_response is valid only for PPTX import.");
    }

    private static void ThinPresentationImportResponse(ArtifactEnvelope artifact)
    {
        var snapshot = artifact.OpaqueOpc?.SourcePackage;
        if (snapshot is not { Data.IsEmpty: false })
            throw new CodecException("missing_source_package", "Thin PPTX import requires a validated source package snapshot.");
        snapshot.Data = ByteString.Empty;
    }

    internal static Diagnostic Error(string code, string message, string? sourcePath = null) =>
        CodecDiagnostics.Error(code, message, sourcePath);

    internal static Diagnostic Warning(string code, string message, string? sourcePath = null, string? sourceIdentity = null) =>
        CodecDiagnostics.Warning(code, message, sourcePath, sourceIdentity);
}
