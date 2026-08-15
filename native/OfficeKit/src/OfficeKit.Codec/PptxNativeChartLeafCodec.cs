using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeKit.Artifact.Wire.V1;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal sealed record PptxNativeChartTitleLeaf(uint Index, string Text, XElement Element);

internal sealed record PptxNativeChartResolution(
    PresentationNativeChart Binding,
    ChartPart Part,
    XDocument Document,
    IReadOnlyList<PptxNativeChartTitleLeaf> TitleLeaves);

// Describes and re-proves a deliberately tiny source-owned chart surface.
// A native chart remains opaque; only direct rich-title a:r/a:t leaves are
// projected. Series, formulas, caches, workbook links, styles, extensions,
// and every other ChartSpace token stay owned by the original package.
internal static class PptxNativeChartLeafCodec
{
    private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string ChartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    private const int MaxTitleLeaves = 256;
    private const int MaxLeafLength = 32_767;

    private static readonly XNamespace ChartNs = ChartNamespace;
    private static readonly XNamespace DrawingNs = DrawingNamespace;

    internal static bool TryDescribe(OpenXmlElement source, OpenXmlPart owner, out PresentationNativeChart binding)
    {
        if (!TryResolve(source, owner, out var resolved))
        {
            binding = null!;
            return false;
        }
        binding = resolved.Binding;
        return true;
    }

    internal static bool TryResolve(OpenXmlElement source, OpenXmlPart owner, out PptxNativeChartResolution resolved)
    {
        resolved = null!;
        if (source is not P.GraphicFrame frame || owner is not SlidePart || source.Parent is not P.ShapeTree ||
            PptxNativeObjectCatalog.Classify(source) != "graphicFrame")
            return false;
        var graphicData = frame.Graphic?.GraphicData;
        if (graphicData is null || graphicData.Uri?.Value != ChartNamespace) return false;
        var references = graphicData.Elements<C.ChartReference>().ToArray();
        if (references.Length != 1 || string.IsNullOrWhiteSpace(references[0].Id?.Value)) return false;
        var relationshipId = references[0].Id!.Value!;

        ChartPart part;
        try
        {
            part = owner.GetPartById(relationshipId) as ChartPart ?? null!;
        }
        catch (ArgumentOutOfRangeException)
        {
            return false;
        }
        if (part is null || !part.ContentType.Equals(ChartContentType, StringComparison.OrdinalIgnoreCase) ||
            !part.RelationshipType.EndsWith("/chart", StringComparison.Ordinal)) return false;

        byte[] bytes;
        XDocument document;
        try
        {
            bytes = ReadPart(part);
            using var memory = new MemoryStream(bytes, writable: false);
            using var reader = XmlReader.Create(memory, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = false,
                IgnoreProcessingInstructions = false,
                IgnoreWhitespace = false,
            });
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (Exception exception) when (exception is XmlException or IOException or UnauthorizedAccessException)
        {
            return false;
        }

        var root = document.Root;
        if (root?.Name != ChartNs + "chartSpace") return false;
        var charts = root.Elements(ChartNs + "chart").ToArray();
        if (charts.Length != 1) return false;
        var titles = charts[0].Elements(ChartNs + "title").ToArray();
        if (titles.Length != 1) return false;
        var textOwners = titles[0].Elements(ChartNs + "tx").ToArray();
        if (textOwners.Length != 1) return false;
        var richBodies = textOwners[0].Elements(ChartNs + "rich").ToArray();
        if (richBodies.Length != 1 || richBodies[0].Descendants(DrawingNs + "fld").Any()) return false;
        var textElements = richBodies[0].Descendants(DrawingNs + "t").ToArray();
        if (textElements.Length is 0 or > MaxTitleLeaves || textElements.Any(element =>
                element.Parent?.Name != DrawingNs + "r" || !ValidText(element.Value))) return false;

        var binding = new PresentationNativeChart
        {
            PartPath = PartPath(part),
            ContentType = part.ContentType,
            SourceSha256 = Hash(bytes),
            RelationshipId = relationshipId,
        };
        var leaves = textElements.Select((element, index) =>
            new PptxNativeChartTitleLeaf(checked((uint)index), element.Value, element)).ToArray();
        binding.TitleLeaves.Add(leaves.Select(leaf => new PresentationNativeChartTitleLeaf
        {
            TextLeafIndex = leaf.Index,
            Text = leaf.Text,
        }));
        resolved = new PptxNativeChartResolution(binding, part, document, leaves);
        return true;
    }

    internal static bool SameBinding(PresentationNativeChart expected, PresentationNativeChart actual) =>
        expected.PartPath.Equals(actual.PartPath, StringComparison.OrdinalIgnoreCase) &&
        expected.ContentType.Equals(actual.ContentType, StringComparison.OrdinalIgnoreCase) &&
        expected.SourceSha256.Equals(actual.SourceSha256, StringComparison.OrdinalIgnoreCase) &&
        expected.RelationshipId == actual.RelationshipId &&
        expected.TitleLeaves.SequenceEqual(actual.TitleLeaves);

    internal static bool HasUniqueInboundRelationship(PresentationPart presentationPart, ChartPart target)
    {
        var count = 0;
        var visited = new HashSet<OpenXmlPart>();
        var pending = new Queue<OpenXmlPart>();
        pending.Enqueue(presentationPart);
        while (pending.Count > 0)
        {
            var owner = pending.Dequeue();
            if (!visited.Add(owner)) continue;
            foreach (var relationship in owner.Parts)
            {
                if (ReferenceEquals(relationship.OpenXmlPart, target)) count++;
                pending.Enqueue(relationship.OpenXmlPart);
            }
        }
        return count == 1;
    }

    private static bool ValidText(string value) =>
        value.Length <= MaxLeafLength &&
        !value.Any(character => character is >= '\u0000' and <= '\u0008' or '\u000B' or '\u000C' or >= '\u000E' and <= '\u001F');

    private static byte[] ReadPart(OpenXmlPart part)
    {
        using var input = part.GetStream(FileMode.Open, FileAccess.Read);
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
