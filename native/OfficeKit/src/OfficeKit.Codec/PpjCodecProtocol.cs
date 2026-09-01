using System.Runtime.InteropServices;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

public static class PpjCodecProtocol
{
    public static CodecResponse InvokeResponse(ref byte[] requestBytes, byte[]? requestFileBytes)
    {
        var response = new CodecResponse { ProtocolVersion = CodecWireProtocol.ProtocolVersion };
        try
        {
            var request = ParseRequest(ref requestBytes, requestFileBytes);
            ValidateRequest(request);
            var limits = EffectiveCodecLimits.From(request.Limits);
            switch (request.Operation)
            {
                case CodecOperation.ProjectPptxToPpj:
                {
                    using var result = PpjPresentationProjector.Project(
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
                    throw Unsupported(request.Operation);
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

    private static CodecRequest ParseRequest(ref byte[] requestBytes, byte[]? requestFileBytes)
    {
        if (requestBytes is null || requestBytes.Length == 0)
            throw new CodecException("empty_request", "Codec request bytes must not be empty.");
        if (requestBytes.Length > CodecWireProtocol.AbsoluteRequestLimit)
            throw new CodecException("request_budget_exceeded", $"Codec request exceeds the absolute {CodecWireProtocol.AbsoluteRequestLimit}-byte wire budget.");
        var request = CodecRequest.Parser.ParseFrom(requestBytes);
        requestBytes = [];
        if (requestFileBytes is not { Length: > 0 }) return request;
        if (!request.File.IsEmpty)
            throw new CodecException("ambiguous_file_payload", "Codec request cannot contain both inline file bytes and a native transport file sidecar.");
        request.File = UnsafeByteOperations.UnsafeWrap(requestFileBytes);
        return request;
    }

    private static void ValidateRequest(CodecRequest request)
    {
        if (request.ProtocolVersion != CodecWireProtocol.ProtocolVersion)
            throw new CodecException("unsupported_protocol_version", $"Protocol version {request.ProtocolVersion} is unsupported; expected {CodecWireProtocol.ProtocolVersion}.");
        if (request.Operation is not (CodecOperation.ProjectPptxToPpj or CodecOperation.CompilePpjToPptx))
            throw Unsupported(request.Operation);
        if (request.Family != ArtifactFamily.Presentation)
            throw new CodecException("artifact_family_mismatch", $"Codec operation {request.Operation} requires artifact family {ArtifactFamily.Presentation}, not {request.Family}.");
        if (request.Operation == CodecOperation.ProjectPptxToPpj && request.File.IsEmpty)
            throw new CodecException("empty_input", "PPTX projection requires non-empty file bytes.");
        if (request.PresentationProgram is null)
            throw new CodecException("missing_presentation_program", "PPJ operations require presentation_program options.");
        if (request.Operation == CodecOperation.CompilePpjToPptx && request.PresentationProgram.ProgramJson.IsEmpty)
            throw new CodecException("empty_ppj", "PPJ compilation requires non-empty program_json bytes.");
        if (request.ThinPresentationImportResponse)
            throw new CodecException("invalid_request", "thin_presentation_import_response is valid only for the legacy PPTX import operation.");
    }

    private static byte[] RequestFileBytes(ByteString file)
    {
        if (MemoryMarshal.TryGetArray(file.Memory, out var segment) && segment.Array is not null &&
            segment.Offset == 0 && segment.Count == segment.Array.Length)
            return segment.Array;
        return file.ToByteArray();
    }

    private static CodecException Unsupported(CodecOperation operation) =>
        new("unsupported_operation", $"Codec operation {operation} is not implemented by the PPJ profile.");

    private static Diagnostic Error(string code, string message, string? sourcePath = null) => new()
    {
        Severity = DiagnosticSeverity.Error,
        Code = code,
        Message = message,
        SourcePath = sourcePath ?? string.Empty,
    };
}
