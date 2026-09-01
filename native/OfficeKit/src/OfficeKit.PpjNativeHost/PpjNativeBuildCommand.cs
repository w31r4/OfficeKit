using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;
using OfficeKit.Codec;

namespace OfficeKit.PpjNativeHost;

internal static class PpjNativeBuildCommand
{
    internal static int Run(string[] args)
    {
        var json = args.Contains("--json", StringComparer.Ordinal);
        try
        {
            if (OperatingSystem.IsWindows())
                throw new CliException("The direct PPJ build path is unavailable on Windows.");
            var options = Parse(args);
            var input = ResolveRegularFile(Path.GetFullPath(options.Input, options.Cwd), "PPJ input");
            if (!Path.GetExtension(input).Equals(".ppj", StringComparison.OrdinalIgnoreCase))
                throw new CliException($"PPJ input must be a .ppj file: {input}");

            var programBytes = File.ReadAllBytes(input);
            if (programBytes.Length is 0 or > PpjProgramValidator.MaxSourceBytes)
                throw new CliException(
                    $"PPJ input must contain 1 through {PpjProgramValidator.MaxSourceBytes} bytes: {input}");

            using var validation = PpjProgramValidator.Validate(programBytes);
            if (!validation.IsValid)
            {
                var first = validation.Diagnostics[0];
                throw new CodecException(first.Code, first.Message, first.Path);
            }

            var directory = Path.GetDirectoryName(input)!;
            var request = new PresentationProgramRequest
            {
                ProgramJson = UnsafeByteOperations.UnsafeWrap(programBytes),
                IncludeNodeMap = true,
            };
            var assetPaths = validation.Program!.Assets.ToDictionary(
                declaration => declaration.Id,
                declaration => ResolveWorkspaceResource(directory, declaration.Uri, $"PPJ asset {declaration.Id}"),
                StringComparer.Ordinal);
            foreach (var declaration in validation.Program!.Assets)
            {
                if (validation.Program.Source is null)
                    request.Assets.Add(LoadAsset(declaration, assetPaths[declaration.Id]));
            }

            var sourcePath = validation.Program.Source is null
                ? null
                : ResolveWorkspaceResource(directory, validation.Program.Source.Uri, "PPJ source package");
            var sourceBytes = sourcePath is null ? [] : File.ReadAllBytes(sourcePath);
            var deferredAssetBytes = validation.Program.Source is null
                ? 0
                : assetPaths.Values.Sum(path => new FileInfo(path).Length);
            var deferredAssetCount = validation.Program.Source is null ? 0 : assetPaths.Count;
            EnsureTransportBudget(request, sourceBytes.LongLength, deferredAssetBytes, deferredAssetCount);
            var destination = Path.GetFullPath(options.Output, options.Cwd);
            if (!Path.GetExtension(destination).Equals(".pptx", StringComparison.OrdinalIgnoreCase))
                throw new CliException($"PPJ build output must be a .pptx file: {destination}");
            if (PathEquals(destination, input) || sourcePath is not null && PathEquals(destination, sourcePath))
                throw new CliException("PPJ build output must not overwrite the PPJ or its bound source PPTX.");
            if (PathExists(destination))
                throw new CliException($"PPTX output already exists: {destination}");

            var compiled = PpjPresentationCompiler.CompileValidated(
                request,
                sourceBytes,
                EffectiveCodecLimits.From(null),
                validation,
                declaration => LoadAsset(declaration, assetPaths[declaration.Id]));
            if (compiled.File.Length == 0 || !Sha256(compiled.File).Equals(compiled.Program.OutputSha256, StringComparison.Ordinal))
                throw new CliException("OfficeKit native compiler returned a PPTX with an invalid content hash.");

            WriteExclusiveFile(destination, compiled.File);
            WriteSuccess(options.Json, input, destination, compiled);
            return 0;
        }
        catch (CodecException exception)
        {
            WriteCodecFailure(json, exception.Code, exception.Message);
            return 1;
        }
        catch (CliException exception)
        {
            WriteCliFailure(json, exception.Message);
            return 1;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException)
        {
            WriteCliFailure(json, exception.Message);
            return 1;
        }
    }

    private static BuildOptions Parse(string[] args)
    {
        string? input = null;
        string? output = null;
        string? cwd = null;
        var json = false;
        for (var index = 0; index < args.Length; index++)
        {
            var argument = args[index];
            if (argument == "--json") json = true;
            else if (argument is "-o" or "--output") output = RequiredValue(args, ref index, argument);
            else if (argument.StartsWith("--output=", StringComparison.Ordinal)) output = argument["--output=".Length..];
            else if (argument == "--cwd") cwd = RequiredValue(args, ref index, argument);
            else if (argument.StartsWith("-", StringComparison.Ordinal))
                throw new CliException($"Unknown PPJ build option \"{argument}\".");
            else if (input is null) input = argument;
            else throw new CliException("PPJ build requires one deck.ppj and -o <deck.pptx>.");
        }
        if (input is null || output is null || cwd is null)
            throw new CliException("PPJ build requires one deck.ppj and -o <deck.pptx>.");
        return new(input, output, Path.GetFullPath(cwd), json);
    }

    private static string RequiredValue(string[] args, ref int index, string option)
    {
        index++;
        if (index >= args.Length || args[index].StartsWith("-", StringComparison.Ordinal))
            throw new CliException($"{option} requires a value.");
        return args[index];
    }

    private static string ResolveWorkspaceResource(string root, string uri, string label)
    {
        if (string.IsNullOrEmpty(uri) || uri.Contains('\\') || uri.Contains('\0') || uri.StartsWith('/') ||
            Uri.TryCreate(uri, UriKind.Absolute, out _) || uri.Split('/').Any(segment => segment == ".."))
            throw new CliException($"{label} URI must stay relative to the PPJ directory: {uri}");
        var target = Path.GetFullPath(Path.Combine(root, uri.Replace('/', Path.DirectorySeparatorChar)));
        if (!IsInside(root, target))
            throw new CliException($"{label} URI escapes the PPJ directory: {uri}");
        var resolved = ResolveRegularFile(target, label);
        if (!IsInside(root, resolved))
            throw new CliException($"{label} resolves outside the PPJ directory: {uri}");
        return resolved;
    }

    private static string ResolveRegularFile(string target, string label)
    {
        var info = new FileInfo(target);
        if (info.LinkTarget is not null)
            throw new CliException($"{label} must be a non-symlink regular file: {target}");
        if (Directory.Exists(target))
            throw new CliException($"{label} must be a non-symlink regular file: {target}");
        if (!File.Exists(target))
            throw new CliException($"{label} does not exist: {target}");
        if ((File.GetAttributes(target) & FileAttributes.ReparsePoint) != 0)
            throw new CliException($"{label} must be a non-symlink regular file: {target}");
        return RealPath(target);
    }

    private static bool IsInside(string root, string target)
    {
        var relative = Path.GetRelativePath(root, target);
        return relative.Length > 0 && relative != "." && !relative.Equals("..", StringComparison.Ordinal) &&
               !relative.StartsWith($"..{Path.DirectorySeparatorChar}", StringComparison.Ordinal) &&
               !Path.IsPathRooted(relative);
    }

    private static string RealPath(string path)
    {
        var pointer = realpath(path, IntPtr.Zero);
        if (pointer == IntPtr.Zero)
            throw new IOException($"Could not resolve real path {path} (errno {Marshal.GetLastPInvokeError()}).");
        try
        {
            return Marshal.PtrToStringUTF8(pointer) ?? throw new IOException($"Could not decode real path {path}.");
        }
        finally
        {
            free(pointer);
        }
    }

    private static void WriteExclusiveFile(string destination, byte[] bytes)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
        var temporary = Path.Combine(
            Path.GetDirectoryName(destination)!,
            $".{Path.GetFileName(destination)}.{Environment.ProcessId}-{Guid.NewGuid():N}.tmp");
        try
        {
            using (var output = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                output.Write(bytes);
            if (!OperatingSystem.IsWindows())
                File.SetUnixFileMode(temporary, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.GroupRead | UnixFileMode.OtherRead);
            File.Move(temporary, destination, overwrite: false);
        }
        catch (IOException) when (PathExists(destination))
        {
            throw new CliException($"PPTX output already exists: {destination}");
        }
        finally
        {
            try { File.Delete(temporary); } catch (IOException) { }
        }
    }

    private static Asset LoadAsset(PpjAssetModel declaration, string assetPath) => new()
    {
        Id = declaration.Id,
        FileName = Path.GetFileName(declaration.Uri),
        ContentType = declaration.MimeType,
        Sha256 = declaration.Sha256,
        Data = UnsafeByteOperations.UnsafeWrap(File.ReadAllBytes(assetPath)),
    };

    private static void EnsureTransportBudget(
        PresentationProgramRequest request,
        long sourceBytes,
        long deferredAssetBytes,
        int deferredAssetCount)
    {
        var wireRequest = new CodecRequest
        {
            ProtocolVersion = CodecWireProtocol.ProtocolVersion,
            Operation = CodecOperation.CompilePpjToPptx,
            Family = ArtifactFamily.Presentation,
            PresentationProgram = request,
        };
        // A deferred protobuf bytes field adds one tag and up to five length
        // bytes at the current transport limit. Keep the direct path under the
        // same aggregate budget without materializing source-owned assets.
        var deferredWireOverhead = checked((long)deferredAssetCount * 6);
        if (checked((long)wireRequest.CalculateSize() + sourceBytes + deferredAssetBytes + deferredWireOverhead) >
            CodecWireProtocol.AbsoluteRequestLimit)
        {
            throw new CodecException(
                "request_budget_exceeded",
                $"OfficeKit native codec request and file sidecar exceed the {CodecWireProtocol.AbsoluteRequestLimit}-byte transport budget.");
        }
    }

    private static void WriteSuccess(bool json, string input, string output, PpjCompileResult compiled)
    {
        if (!json)
        {
            Console.Out.WriteLine($"OfficeKit built {output}");
            Console.Out.WriteLine($"Program   {compiled.Program.ProgramSha256}");
            Console.Out.WriteLine($"PPTX      {compiled.Program.OutputSha256}");
            Console.Out.WriteLine($"Changed   {compiled.Program.ChangedNodeIds.Count} nodes / {compiled.Program.ChangedParts.Count} parts");
            return;
        }

        using var writer = new Utf8JsonWriter(Console.OpenStandardOutput());
        writer.WriteStartObject();
        writer.WriteBoolean("ok", true);
        writer.WriteString("command", "build");
        writer.WriteString("input", input);
        writer.WriteString("output", output);
        writer.WriteString("programSha256", compiled.Program.ProgramSha256);
        writer.WriteString("outputSha256", compiled.Program.OutputSha256);
        writer.WriteBoolean("sourceBound", compiled.Program.SourceBound);
        if (compiled.Program.SourceSha256.Length == 0) writer.WriteNull("sourceSha256");
        else writer.WriteString("sourceSha256", compiled.Program.SourceSha256);
        writer.WriteNumber("expandedElementCount", compiled.Program.ExpandedElementCount);
        WriteStrings(writer, "changedParts", compiled.Program.ChangedParts);
        WriteStrings(writer, "changedNodeIds", compiled.Program.ChangedNodeIds);
        writer.WriteStartArray("diagnostics");
        foreach (var diagnostic in compiled.Diagnostics)
        {
            writer.WriteStartObject();
            writer.WriteNumber("severity", (int)diagnostic.Severity);
            writer.WriteString("code", diagnostic.Code);
            writer.WriteString("message", diagnostic.Message);
            writer.WriteString("sourcePath", diagnostic.SourcePath);
            writer.WriteString("sourceIdentity", diagnostic.SourceIdentity);
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
        writer.WriteNull("task");
        writer.WriteEndObject();
        writer.Flush();
        Console.Out.WriteLine();
    }

    private static void WriteStrings(Utf8JsonWriter writer, string name, IEnumerable<string> values)
    {
        writer.WriteStartArray(name);
        foreach (var value in values) writer.WriteStringValue(value);
        writer.WriteEndArray();
    }

    private static void WriteCliFailure(bool json, string message)
    {
        if (!json)
        {
            Console.Error.WriteLine($"OfficeKit: {message}");
            return;
        }
        using var writer = new Utf8JsonWriter(Console.OpenStandardError());
        writer.WriteStartObject();
        writer.WriteBoolean("ok", false);
        writer.WriteString("error", message);
        writer.WriteEndObject();
        writer.Flush();
        Console.Error.WriteLine();
    }

    private static void WriteCodecFailure(bool json, string code, string message)
    {
        if (!json)
        {
            Console.Error.WriteLine($"OfficeKit: {message}");
            return;
        }
        using var writer = new Utf8JsonWriter(Console.OpenStandardError());
        writer.WriteStartObject();
        writer.WriteNumber("protocol", 1);
        writer.WriteBoolean("ok", false);
        writer.WriteStartObject("error");
        writer.WriteString("code", code);
        writer.WriteString("message", message);
        writer.WriteBoolean("retryable", false);
        writer.WriteBoolean("maybeApplied", false);
        writer.WriteEndObject();
        writer.WriteEndObject();
        writer.Flush();
        Console.Error.WriteLine();
    }

    private static bool PathExists(string path) => File.Exists(path) || Directory.Exists(path);
    private static bool PathEquals(string left, string right) =>
        left.Equals(right, StringComparison.Ordinal);
    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private sealed record BuildOptions(string Input, string Output, string Cwd, bool Json);

    private sealed class CliException(string message) : Exception(message);

    [DllImport("libc", SetLastError = true)]
    private static extern IntPtr realpath([MarshalAs(UnmanagedType.LPUTF8Str)] string path, IntPtr resolvedPath);

    [DllImport("libc")]
    private static extern void free(IntPtr pointer);
}
