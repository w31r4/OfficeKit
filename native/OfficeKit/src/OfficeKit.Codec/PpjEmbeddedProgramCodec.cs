using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

/// <summary>
/// Owns the private OPC snapshot carried only by presentations compiled from
/// an authored PPJ. The snapshot is source code, not a second presentation
/// model: canonical PPJ and its declared assets remain authoritative, while
/// the map binds stable program identities to the native output produced by
/// this compiler revision.
/// </summary>
internal static class PpjEmbeddedProgramCodec
{
    internal const string ProgramPath = "officeKit/program.ppj";
    internal const string ProgramMapPath = "officeKit/program-map.json";
    internal const string ProgramRelationshipsPath = "officeKit/_rels/program.ppj.rels";
    internal const string ProgramContentType = "application/vnd.officekit.ppj+json";
    internal const string ProgramMapContentType = "application/vnd.officekit.ppj-map+json";
    internal const string ProgramRelationshipType = "https://schemas.officekit.dev/relationships/presentation-program";
    internal const string ProgramMapRelationshipType = "https://schemas.officekit.dev/relationships/presentation-program-map";
    internal const string ProgramAssetRelationshipType = "https://schemas.officekit.dev/relationships/presentation-program-asset";

    private const string ContentTypesPath = "[Content_Types].xml";
    private const string RootRelationshipsPath = "_rels/.rels";
    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    private sealed record EmbeddedAsset(
        string Id,
        string Uri,
        string MimeType,
        string Sha256,
        string PartPath,
        byte[] Data);

    internal sealed record Recovery(
        PresentationProgramResult Program,
        IReadOnlyList<Diagnostic> Diagnostics);

    internal static Recovery? TryRecover(
        byte[] pptx,
        PresentationProgramRequest request,
        EffectiveCodecLimits limits)
    {
        _ = PackageGuards.ValidateAndCollectOpaque(pptx, limits, OpcPackageProfile.Pptx, includeSourcePackage: false);
        try
        {
            var parts = ReadParts(pptx);
            if (!HasExactRelationship(
                    parts[RootRelationshipsPath],
                    ProgramRelationshipType,
                    ProgramPath,
                    requireSingle: true))
                return null;
            if (!parts.TryGetValue(ProgramPath, out var programBytes) ||
                !parts.TryGetValue(ProgramMapPath, out var mapBytes) ||
                !parts.TryGetValue(ProgramRelationshipsPath, out var relationshipBytes))
                return null;

            var contentTypes = ContentTypeOverrides(parts[ContentTypesPath]);
            if (!HasContentType(contentTypes, ProgramPath, ProgramContentType) ||
                !HasContentType(contentTypes, ProgramMapPath, ProgramMapContentType) ||
                !HasExactRelationship(
                    relationshipBytes,
                    ProgramMapRelationshipType,
                    "program-map.json",
                    requireSingle: true))
                return null;

            using var validation = PpjProgramValidator.Validate(programBytes);
            if (!validation.IsValid || validation.Program is null || validation.Program.Source is not null || validation.Expansion is null)
                return null;

            using var map = JsonDocument.Parse(mapBytes, new JsonDocumentOptions
            {
                AllowTrailingCommas = false,
                CommentHandling = JsonCommentHandling.Disallow,
                MaxDepth = 96,
            });
            var root = map.RootElement;
            if (root.ValueKind != JsonValueKind.Object ||
                root.GetProperty("schema").GetString() != "office-kit/ppj-map/v1" ||
                !Hash(root.GetProperty("programSha256").GetString()).Equals(validation.ProgramSha256, StringComparison.OrdinalIgnoreCase) ||
                root.GetProperty("expandedElementCount").GetInt32() != validation.Expansion.ExpandedElementCount)
                return null;

            var nodeMap = Encoding.UTF8.GetBytes(root.GetProperty("nodeMap").GetRawText());
            var nodeMapSha256 = Hash(root.GetProperty("nodeMapSha256").GetString());
            if (!Sha256(nodeMap).Equals(nodeMapSha256, StringComparison.OrdinalIgnoreCase) ||
                !nodeMapSha256.Equals(validation.Expansion.NodeMapSha256, StringComparison.OrdinalIgnoreCase) ||
                !nodeMap.AsSpan().SequenceEqual(validation.Expansion.NodeMapJson))
                return null;

            if (!ValidateNativeBindings(root.GetProperty("nativeBindings"), validation.Expansion, parts))
                return null;

            var assets = RecoverAssets(
                validation.Program,
                root.GetProperty("assets"),
                parts,
                contentTypes,
                relationshipBytes);
            if (assets is null) return null;

            var diagnostics = NativeDriftDiagnostics(root.GetProperty("nativePackage"), parts);
            var result = new PresentationProgramResult
            {
                ProgramJson = ByteString.CopyFrom(validation.CanonicalJson),
                OriginalProgramJson = ByteString.CopyFrom(programBytes),
                ProgramSha256 = validation.ProgramSha256,
                NodeMapJson = request.IncludeNodeMap ? ByteString.CopyFrom(nodeMap) : ByteString.Empty,
                RestoredEmbeddedProgram = true,
                SourceBound = false,
                ExpandedElementCount = checked((uint)validation.Expansion.ExpandedElementCount),
            };
            result.Assets.Add(assets);
            return new(result, diagnostics);
        }
        catch (Exception exception) when (exception is JsonException or XmlException or InvalidDataException or
                                          KeyNotFoundException or InvalidOperationException or FormatException or
                                          OverflowException or ArgumentException or CodecException)
        {
            // An unusable snapshot is not write authority. Ordinary PPTX
            // projection below will preserve it as opaque source-owned data.
            return null;
        }
    }

    internal static byte[] Embed(
        byte[] pptx,
        ReadOnlySpan<byte> originalProgramJson,
        PpjValidationResult validation,
        PresentationArtifact presentation,
        IReadOnlyList<Asset> compiledAssets,
        EffectiveCodecLimits limits)
    {
        if (!validation.IsValid || validation.Program is null || validation.Expansion is null)
            throw new InvalidOperationException("Only a validated authored PPJ can be embedded.");
        if (validation.Program.Source is not null)
            throw new InvalidOperationException("Source-bound PPJ must never be embedded into third-party output.");
        if (originalProgramJson.IsEmpty)
            throw new InvalidOperationException("Authored PPJ source bytes are required for exact recovery.");

        var parts = ReadParts(pptx);
        AddToSourceFreePackage(
            parts,
            originalProgramJson,
            validation,
            NativeBindings(presentation, validation.Expansion),
            compiledAssets);
        var output = WriteDeterministicArchive(parts);
        ValidateEmbeddedOutput(output, limits);
        return output;
    }

    internal static void AddToSourceFreePackage(
        Dictionary<string, byte[]> parts,
        ReadOnlySpan<byte> originalProgramJson,
        PpjValidationResult validation,
        IReadOnlyList<PptxNativeBinding> bindings,
        IReadOnlyList<Asset> compiledAssets)
    {
        if (!validation.IsValid || validation.Program is null || validation.Expansion is null)
            throw new InvalidOperationException("Only a validated authored PPJ can be embedded.");
        if (validation.Program.Source is not null)
            throw new InvalidOperationException("Source-bound PPJ must never be embedded into third-party output.");
        if (originalProgramJson.IsEmpty)
            throw new InvalidOperationException("Authored PPJ source bytes are required for exact recovery.");

        RejectReservedParts(parts);
        var assets = BindAssets(validation.Program, compiledAssets);
        var nativeParts = parts
            .Where(part => !part.Key.Equals(ContentTypesPath, StringComparison.OrdinalIgnoreCase) &&
                           !part.Key.Equals(RootRelationshipsPath, StringComparison.OrdinalIgnoreCase) &&
                           !part.Key.StartsWith("officeKit/", StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Key, StringComparer.Ordinal)
            .Select(part => (Path: part.Key, Sha256: Sha256(part.Value)))
            .ToArray();
        var map = WriteProgramMap(validation, assets, bindings, nativeParts);

        parts[ContentTypesPath] = AddContentTypes(parts[ContentTypesPath], assets);
        parts[RootRelationshipsPath] = AddRootProgramRelationship(parts[RootRelationshipsPath]);
        parts[ProgramPath] = originalProgramJson.ToArray();
        parts[ProgramMapPath] = map;
        parts[ProgramRelationshipsPath] = WriteProgramRelationships(assets);
        foreach (var asset in assets.GroupBy(asset => asset.PartPath, StringComparer.Ordinal).Select(group => group.First()))
            parts[asset.PartPath] = asset.Data;
    }

    internal static void ValidateEmbeddedOutput(byte[] output, EffectiveCodecLimits limits)
    {
        _ = PackageGuards.ValidateAndCollectOpaque(output, limits, OpcPackageProfile.Pptx, includeSourcePackage: false);
    }

    private static Dictionary<string, byte[]> ReadParts(byte[] pptx)
    {
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        using var source = new MemoryStream(pptx, writable: false);
        using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: false);
        foreach (var entry in archive.Entries)
        {
            if (entry.FullName.EndsWith("/", StringComparison.Ordinal)) continue;
            using var entryStream = entry.Open();
            using var copy = new MemoryStream();
            entryStream.CopyTo(copy);
            if (!parts.TryAdd(entry.FullName, copy.ToArray()))
                throw new CodecException("ppj.embedded.duplicatePart", $"Generated PPTX contains duplicate part {entry.FullName}.", entry.FullName);
        }
        if (!parts.ContainsKey(ContentTypesPath) || !parts.ContainsKey(RootRelationshipsPath))
            throw new CodecException("ppj.embedded.package", "Generated PPTX is missing required OPC package metadata.");
        return parts;
    }

    private static void RejectReservedParts(IReadOnlyDictionary<string, byte[]> parts)
    {
        var collision = parts.Keys.FirstOrDefault(path => path.StartsWith("officeKit/", StringComparison.OrdinalIgnoreCase));
        if (collision is not null)
            throw new CodecException("ppj.embedded.reservedPart", $"Generated PPTX already contains reserved OfficeKit part {collision}.", collision);
    }

    private static IReadOnlyList<EmbeddedAsset> BindAssets(PpjProgramModel program, IReadOnlyList<Asset> compiledAssets)
    {
        var byHash = compiledAssets
            .GroupBy(asset => asset.Sha256, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
        var output = new List<EmbeddedAsset>(program.Assets.Count);
        foreach (var declaration in program.Assets)
        {
            if (!byHash.TryGetValue(declaration.Sha256, out var compiled) ||
                !compiled.ContentType.Equals(declaration.MimeType, StringComparison.OrdinalIgnoreCase) ||
                !Sha256(compiled.Data.Span).Equals(declaration.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "ppj.embedded.assetBinding",
                    $"Compiled asset bytes for PPJ asset {declaration.Id} are unavailable.",
                    "$.assets");
            output.Add(new(
                declaration.Id,
                declaration.Uri,
                declaration.MimeType,
                declaration.Sha256.ToLowerInvariant(),
                $"officeKit/assets/{declaration.Sha256.ToLowerInvariant()}.bin",
                compiled.Data.ToByteArray()));
        }
        return output;
    }

    private static IReadOnlyList<PptxNativeBinding> NativeBindings(
        PresentationArtifact presentation,
        PpjExpansionResult expansion)
    {
        var semanticIds = expansion.Nodes.Select(node => node.Id).ToHashSet(StringComparer.Ordinal);
        var output = new List<PptxNativeBinding>();
        for (var pageIndex = 0; pageIndex < presentation.Slides.Count; pageIndex++)
        {
            var slide = presentation.Slides[pageIndex];
            var flattened = Flatten(slide.Elements).ToArray();
            for (var index = 0; index < flattened.Length; index++)
            {
                var element = flattened[index];
                // A bounded semantic primitive may lower to several native
                // DrawingML children (for example a PPJ heatmap becomes one
                // editable group of cells and labels). Those compiler-owned
                // children are not independent PPJ nodes. Keep their native
                // slots in the ID count, but bind only stable program IDs.
                if (!semanticIds.Contains(element.Id)) continue;
                output.Add(new PptxNativeBinding(
                    slide.Id,
                    element.Id,
                    element.ContentCase.ToString(),
                    $"ppt/slides/slide{pageIndex + 1}.xml",
                    checked((uint)(index + 2))));
            }
        }
        return output;
    }

    private static IEnumerable<PresentationElement> Flatten(IEnumerable<PresentationElement> elements)
    {
        foreach (var element in elements)
        {
            yield return element;
            if (element.ContentCase == PresentationElement.ContentOneofCase.Group)
                foreach (var child in Flatten(element.Group.Children)) yield return child;
        }
    }

    private static byte[] WriteProgramMap(
        PpjValidationResult validation,
        IReadOnlyList<EmbeddedAsset> assets,
        IReadOnlyList<PptxNativeBinding> bindings,
        IReadOnlyList<(string Path, string Sha256)> nativeParts)
    {
        using var nodeMap = JsonDocument.Parse(validation.Expansion!.NodeMapJson);
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = false }))
        {
            writer.WriteStartObject();
            writer.WriteString("schema", "office-kit/ppj-map/v1");
            writer.WriteString("programSha256", validation.ProgramSha256);
            writer.WriteString("nodeMapSha256", validation.Expansion.NodeMapSha256);
            writer.WriteNumber("expandedElementCount", validation.Expansion.ExpandedElementCount);
            writer.WritePropertyName("nodeMap");
            nodeMap.RootElement.WriteTo(writer);
            writer.WriteStartArray("nativeBindings");
            foreach (var binding in bindings.OrderBy(item => item.PartPath, StringComparer.Ordinal).ThenBy(item => item.NativeId))
            {
                writer.WriteStartObject();
                writer.WriteString("page", binding.PageId);
                writer.WriteString("id", binding.ElementId);
                writer.WriteString("type", binding.Type);
                writer.WriteString("part", binding.PartPath);
                writer.WriteNumber("nativeId", binding.NativeId);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteStartArray("assets");
            foreach (var asset in assets)
            {
                writer.WriteStartObject();
                writer.WriteString("id", asset.Id);
                writer.WriteString("uri", asset.Uri);
                writer.WriteString("mimeType", asset.MimeType);
                writer.WriteString("sha256", asset.Sha256);
                writer.WriteString("part", asset.PartPath);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteStartObject("nativePackage");
            writer.WriteNumber("partCount", nativeParts.Count);
            writer.WriteStartArray("parts");
            foreach (var part in nativeParts)
            {
                writer.WriteStartObject();
                writer.WriteString("path", part.Path);
                writer.WriteString("sha256", part.Sha256);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
            writer.WriteEndObject();
        }
        return stream.ToArray();
    }

    private static bool ValidateNativeBindings(
        JsonElement bindings,
        PpjExpansionResult expansion,
        IReadOnlyDictionary<string, byte[]> parts)
    {
        if (bindings.ValueKind != JsonValueKind.Array) return false;
        var expected = expansion.Nodes.Select(node => node.Id).ToHashSet(StringComparer.Ordinal);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var nativeSlots = new HashSet<string>(StringComparer.Ordinal);
        foreach (var binding in bindings.EnumerateArray())
        {
            if (binding.ValueKind != JsonValueKind.Object) return false;
            var id = binding.GetProperty("id").GetString();
            var page = binding.GetProperty("page").GetString();
            var type = binding.GetProperty("type").GetString();
            var part = binding.GetProperty("part").GetString();
            var nativeId = binding.GetProperty("nativeId").GetUInt32();
            if (id is null || page is null || type is null || part is null || nativeId < 2 ||
                !expected.Contains(id) || !seen.Add(id) || !parts.ContainsKey(part) ||
                !nativeSlots.Add($"{part}\0{nativeId}"))
                return false;
        }
        return seen.SetEquals(expected);
    }

    private static IReadOnlyList<Asset>? RecoverAssets(
        PpjProgramModel program,
        JsonElement mapAssets,
        IReadOnlyDictionary<string, byte[]> parts,
        IReadOnlyDictionary<string, string> contentTypes,
        byte[] relationshipBytes)
    {
        if (mapAssets.ValueKind != JsonValueKind.Array || mapAssets.GetArrayLength() != program.Assets.Count)
            return null;
        var records = mapAssets.EnumerateArray().ToDictionary(
            item => item.GetProperty("id").GetString() ?? string.Empty,
            item => item,
            StringComparer.Ordinal);
        if (records.Count != program.Assets.Count) return null;
        var relationshipTargets = RelationshipTargets(relationshipBytes, ProgramAssetRelationshipType);
        var output = new List<Asset>(program.Assets.Count);
        foreach (var declaration in program.Assets)
        {
            if (!records.TryGetValue(declaration.Id, out var record)) return null;
            var uri = record.GetProperty("uri").GetString();
            var mimeType = record.GetProperty("mimeType").GetString();
            var sha256 = Hash(record.GetProperty("sha256").GetString());
            var part = record.GetProperty("part").GetString();
            if (uri != declaration.Uri || mimeType is null ||
                !mimeType.Equals(declaration.MimeType, StringComparison.OrdinalIgnoreCase) ||
                !sha256.Equals(declaration.Sha256, StringComparison.OrdinalIgnoreCase) ||
                part is null || !part.Equals($"officeKit/assets/{sha256}.bin", StringComparison.Ordinal) ||
                !parts.TryGetValue(part, out var data) || !Sha256(data).Equals(sha256, StringComparison.OrdinalIgnoreCase) ||
                !HasContentType(contentTypes, part, mimeType) ||
                !relationshipTargets.Contains($"assets/{sha256}.bin"))
                return null;
            output.Add(new Asset
            {
                Id = declaration.Id,
                FileName = declaration.Uri,
                ContentType = declaration.MimeType,
                Data = ByteString.CopyFrom(data),
                Sha256 = sha256,
            });
        }
        var expectedRelationships = program.Assets.Select(asset => $"assets/{asset.Sha256.ToLowerInvariant()}.bin").ToHashSet(StringComparer.Ordinal);
        return relationshipTargets.SetEquals(expectedRelationships) ? output : null;
    }

    private static IReadOnlyList<Diagnostic> NativeDriftDiagnostics(
        JsonElement nativePackage,
        IReadOnlyDictionary<string, byte[]> parts)
    {
        if (nativePackage.ValueKind != JsonValueKind.Object ||
            nativePackage.GetProperty("parts").ValueKind != JsonValueKind.Array)
            throw new InvalidDataException("Embedded PPJ native package map is invalid.");
        var changed = 0;
        foreach (var record in nativePackage.GetProperty("parts").EnumerateArray())
        {
            var path = record.GetProperty("path").GetString();
            var expected = Hash(record.GetProperty("sha256").GetString());
            if (path is null || !parts.TryGetValue(path, out var bytes) ||
                !Sha256(bytes).Equals(expected, StringComparison.OrdinalIgnoreCase))
                changed++;
        }
        if (changed == 0) return [];
        return
        [
            new Diagnostic
            {
                Severity = DiagnosticSeverity.Warning,
                Code = "ppj.embedded.nativeDriftIgnored",
                Message = $"Restored the authoritative embedded PPJ and ignored native drift in {changed} mapped PPTX part(s).",
                SourcePath = ProgramMapPath,
            },
        ];
    }

    private static IReadOnlyDictionary<string, string> ContentTypeOverrides(byte[] bytes)
    {
        var document = LoadXml(bytes, ContentTypesPath);
        XNamespace ns = ContentTypesNamespace;
        var root = document.Root ?? throw new InvalidDataException("PPTX content types have no root element.");
        return root.Elements(ns + "Override").ToDictionary(
            item => (item.Attribute("PartName")?.Value ?? string.Empty).TrimStart('/'),
            item => item.Attribute("ContentType")?.Value ?? string.Empty,
            StringComparer.OrdinalIgnoreCase);
    }

    private static bool HasContentType(
        IReadOnlyDictionary<string, string> overrides,
        string part,
        string expected) =>
        overrides.TryGetValue(part, out var actual) && actual.Equals(expected, StringComparison.OrdinalIgnoreCase);

    private static bool HasExactRelationship(byte[] bytes, string type, string target, bool requireSingle)
    {
        var targets = RelationshipTargets(bytes, type);
        return targets.Contains(target) && (!requireSingle || targets.Count == 1);
    }

    private static HashSet<string> RelationshipTargets(byte[] bytes, string type)
    {
        var document = LoadXml(bytes, "relationships");
        XNamespace ns = RelationshipsNamespace;
        var root = document.Root ?? throw new InvalidDataException("Relationship part has no root element.");
        var targets = new HashSet<string>(StringComparer.Ordinal);
        foreach (var relationship in root.Elements(ns + "Relationship").Where(item =>
                     string.Equals(item.Attribute("Type")?.Value, type, StringComparison.Ordinal)))
        {
            if (string.Equals(relationship.Attribute("TargetMode")?.Value, "External", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Embedded PPJ relationships must be internal.");
            var target = relationship.Attribute("Target")?.Value;
            if (string.IsNullOrWhiteSpace(target) || !targets.Add(target))
                throw new InvalidDataException("Embedded PPJ relationship target is invalid.");
        }
        return targets;
    }

    private static string Hash(string? value)
    {
        if (value is null || value.Length != 64 || !value.All(char.IsAsciiHexDigit))
            throw new InvalidDataException("Embedded PPJ map contains an invalid SHA-256 value.");
        return value.ToLowerInvariant();
    }

    private static byte[] AddContentTypes(byte[] bytes, IReadOnlyList<EmbeddedAsset> assets)
    {
        var document = LoadXml(bytes, ContentTypesPath);
        XNamespace ns = ContentTypesNamespace;
        var root = document.Root ?? throw new CodecException("ppj.embedded.contentTypes", "PPTX content types have no root element.", ContentTypesPath);
        AddOverride(root, ns, ProgramPath, ProgramContentType);
        AddOverride(root, ns, ProgramMapPath, ProgramMapContentType);
        foreach (var asset in assets.GroupBy(asset => asset.PartPath, StringComparer.Ordinal).Select(group => group.First()))
            AddOverride(root, ns, asset.PartPath, asset.MimeType);
        return WriteXml(document);
    }

    private static void AddOverride(XElement root, XNamespace ns, string path, string contentType)
    {
        var partName = $"/{path}";
        if (root.Elements(ns + "Override").Any(item =>
                string.Equals(item.Attribute("PartName")?.Value, partName, StringComparison.OrdinalIgnoreCase)))
            throw new CodecException("ppj.embedded.contentTypeCollision", $"PPTX already declares reserved part {partName}.", ContentTypesPath);
        root.Add(new XElement(ns + "Override", new XAttribute("PartName", partName), new XAttribute("ContentType", contentType)));
    }

    private static byte[] AddRootProgramRelationship(byte[] bytes)
    {
        var document = LoadXml(bytes, RootRelationshipsPath);
        XNamespace ns = RelationshipsNamespace;
        var root = document.Root ?? throw new CodecException("ppj.embedded.relationships", "PPTX package relationships have no root element.", RootRelationshipsPath);
        if (root.Elements(ns + "Relationship").Any(item =>
                string.Equals(item.Attribute("Type")?.Value, ProgramRelationshipType, StringComparison.Ordinal)))
            throw new CodecException("ppj.embedded.relationshipCollision", "PPTX already declares an OfficeKit program relationship.", RootRelationshipsPath);
        var ids = root.Elements(ns + "Relationship").Select(item => item.Attribute("Id")?.Value ?? string.Empty).ToHashSet(StringComparer.Ordinal);
        var id = "rIdOfficeKitProgram";
        for (var suffix = 1; ids.Contains(id); suffix++) id = $"rIdOfficeKitProgram{suffix}";
        root.Add(new XElement(ns + "Relationship",
            new XAttribute("Id", id),
            new XAttribute("Type", ProgramRelationshipType),
            new XAttribute("Target", ProgramPath)));
        return WriteXml(document);
    }

    private static byte[] WriteProgramRelationships(IReadOnlyList<EmbeddedAsset> assets)
    {
        XNamespace ns = RelationshipsNamespace;
        var root = new XElement(ns + "Relationships",
            new XElement(ns + "Relationship",
                new XAttribute("Id", "rIdOfficeKitProgramMap"),
                new XAttribute("Type", ProgramMapRelationshipType),
                new XAttribute("Target", "program-map.json")));
        var unique = assets.GroupBy(asset => asset.PartPath, StringComparer.Ordinal).Select(group => group.First()).OrderBy(asset => asset.PartPath, StringComparer.Ordinal).ToArray();
        for (var index = 0; index < unique.Length; index++)
            root.Add(new XElement(ns + "Relationship",
                new XAttribute("Id", $"rIdOfficeKitAsset{index + 1:D4}"),
                new XAttribute("Type", ProgramAssetRelationshipType),
                new XAttribute("Target", $"assets/{unique[index].Sha256}.bin")));
        return WriteXml(new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root));
    }

    private static XDocument LoadXml(byte[] bytes, string path)
    {
        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = PpjProgramValidator.MaxSourceBytes,
            });
            return XDocument.Load(reader, LoadOptions.None);
        }
        catch (XmlException exception)
        {
            throw new CodecException("ppj.embedded.xml", $"PPTX package metadata is invalid: {exception.Message}", path);
        }
    }

    private static byte[] WriteXml(XDocument document)
    {
        using var stream = new MemoryStream();
        using (var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            NewLineHandling = NewLineHandling.None,
            OmitXmlDeclaration = false,
        }))
        {
            document.Save(writer);
        }
        return stream.ToArray();
    }

    private static byte[] WriteDeterministicArchive(IReadOnlyDictionary<string, byte[]> parts)
    {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            var timestamp = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
            foreach (var part in parts.OrderBy(part => part.Key, StringComparer.Ordinal))
            {
                var entry = archive.CreateEntry(part.Key, CompressionLevel.Optimal);
                entry.LastWriteTime = timestamp;
                using var target = entry.Open();
                target.Write(part.Value);
            }
        }
        return output.ToArray();
    }

    private static string Sha256(byte[] bytes) => Sha256(bytes.AsSpan());
    private static string Sha256(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
