using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

public static class CodecWireProtocol
{
    public const uint ProtocolVersion = 2;
    public const int AbsoluteRequestLimit = 128 * 1024 * 1024;
}

internal static class CodecDiagnostics
{
    internal static Diagnostic Error(string code, string message, string? sourcePath = null) => new()
    {
        Severity = DiagnosticSeverity.Error,
        Code = code,
        Message = message,
        SourcePath = sourcePath ?? string.Empty,
    };

    internal static Diagnostic Warning(
        string code,
        string message,
        string? sourcePath = null,
        string? sourceIdentity = null) => new()
    {
        Severity = DiagnosticSeverity.Warning,
        Code = code,
        Message = message,
        SourcePath = sourcePath ?? string.Empty,
        SourceIdentity = sourceIdentity ?? string.Empty,
    };
}
