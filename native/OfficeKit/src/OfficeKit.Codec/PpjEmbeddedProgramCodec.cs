using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
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

    private sealed record NativeBinding(
        string PageId,
        string ElementId,
        string Type,
        string PartPath,
        uint NativeId);

    internal static byte[] Embed(
        byte[] pptx,
        PpjValidationResult validation,
        PresentationArtifact presentation,
        IReadOnlyList<Asset> compiledAssets,
        EffectiveCodecLimits limits)
    {
        if (!validation.IsValid || validation.Program is null || validation.Expansion is null)
            throw new InvalidOperationException("Only a validated authored PPJ can be embedded.");
        if (validation.Program.Source is not null)
            throw new InvalidOperationException("Source-bound PPJ must never be embedded into third-party output.");

        var parts = ReadParts(pptx);
        RejectReservedParts(parts);
        var assets = BindAssets(validation.Program, compiledAssets);
        var nativeParts = parts
            .Where(part => !part.Key.Equals(ContentTypesPath, StringComparison.OrdinalIgnoreCase) &&
                           !part.Key.Equals(RootRelationshipsPath, StringComparison.OrdinalIgnoreCase) &&
                           !part.Key.StartsWith("officeKit/", StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Key, StringComparer.Ordinal)
            .Select(part => (Path: part.Key, Sha256: Sha256(part.Value)))
            .ToArray();
        var bindings = NativeBindings(presentation);
        var map = WriteProgramMap(validation, assets, bindings, nativeParts);

        parts[ContentTypesPath] = AddContentTypes(parts[ContentTypesPath], assets);
        parts[RootRelationshipsPath] = AddRootProgramRelationship(parts[RootRelationshipsPath]);
        parts[ProgramPath] = validation.CanonicalJson;
        parts[ProgramMapPath] = map;
        parts[ProgramRelationshipsPath] = WriteProgramRelationships(assets);
        foreach (var asset in assets.GroupBy(asset => asset.PartPath, StringComparer.Ordinal).Select(group => group.First()))
            parts[asset.PartPath] = asset.Data;

        var output = WriteDeterministicArchive(parts);
        _ = PackageGuards.ValidateAndCollectOpaque(output, limits, OpcPackageProfile.Pptx, includeSourcePackage: false);
        return output;
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

    private static IReadOnlyList<NativeBinding> NativeBindings(PresentationArtifact presentation)
    {
        var output = new List<NativeBinding>();
        for (var pageIndex = 0; pageIndex < presentation.Slides.Count; pageIndex++)
        {
            var slide = presentation.Slides[pageIndex];
            var flattened = Flatten(slide.Elements).ToArray();
            for (var index = 0; index < flattened.Length; index++)
            {
                var element = flattened[index];
                output.Add(new(
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
        IReadOnlyList<NativeBinding> bindings,
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
