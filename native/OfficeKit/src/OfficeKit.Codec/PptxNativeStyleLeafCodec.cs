using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal sealed record PptxNativeStyleLeaf(uint Index, string Kind, string Value, P.Shape Shape);

// Describes only the smallest style surface that can be proven without
// rebuilding an opaque group.  The group remains source-bound; this codec
// never edits geometry, effects, children, or unsupported fill topologies.
internal static class PptxNativeStyleLeafCodec
{
    private const int MaxLeaves = 4_096;

    internal static bool TryDescribe(OpenXmlElement source, out IReadOnlyList<PptxNativeStyleLeaf> leaves)
    {
        leaves = Array.Empty<PptxNativeStyleLeaf>();
        if (source is not P.GroupShape group) return false;

        var fillLeaves = new List<(string Kind, string Value, P.Shape Shape)>();
        var lineLeaves = new List<(string Kind, string Value, P.Shape Shape)>();
        foreach (var shape in group.Descendants<P.Shape>())
        {
            var properties = shape.ShapeProperties;
            if (properties is null) continue;
            var fills = properties.ChildElements
                .Where(child => child is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill)
                .ToArray();
            if (fills.Length == 1 && fills[0] is A.SolidFill solid && TryReadColor(solid, "fill", out var fillKind, out var fillValue))
                fillLeaves.Add((fillKind, fillValue, shape));

            var outlines = properties.Elements<A.Outline>().ToArray();
            if (outlines.Length != 1) continue;
            var lineFills = outlines[0].ChildElements
                .Where(child => child is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill)
                .ToArray();
            if (lineFills.Length == 1 && lineFills[0] is A.SolidFill lineSolid && TryReadColor(lineSolid, "line", out var lineKind, out var lineValue))
                lineLeaves.Add((lineKind, lineValue, shape));
        }

        var described = fillLeaves.Concat(lineLeaves).ToArray();
        if (described.Length == 0 || described.Length > MaxLeaves) return false;
        leaves = described.Select((item, index) => new PptxNativeStyleLeaf(checked((uint)index), item.Kind, item.Value, item.Shape)).ToArray();
        return true;
    }

    internal static bool TryResolve(OpenXmlElement source, uint index, out PptxNativeStyleLeaf leaf)
    {
        leaf = null!;
        if (!TryDescribe(source, out var leaves) || index >= (uint)leaves.Count) return false;
        leaf = leaves[(int)index];
        return true;
    }

    private static bool TryReadColor(A.SolidFill solid, string prefix, out string kind, out string value)
    {
        kind = string.Empty;
        value = string.Empty;
        if (solid.ChildElements.Count != 1 || solid.FirstChild is null || solid.FirstChild.ChildElements.Count != 0) return false;
        if (solid.FirstChild is A.SchemeColor scheme &&
            scheme.Val?.Value is { } schemeValue &&
            scheme.GetAttributes().All(attribute => attribute.LocalName == "val") &&
            PptxColor.TrySchemeToken(schemeValue, out var schemeToken))
        {
            kind = $"{prefix}Scheme";
            value = schemeToken;
            return true;
        }
        if (solid.FirstChild is A.RgbColorModelHex rgb &&
            rgb.Val?.Value is { Length: 6 } rgbValue && rgbValue.All(Uri.IsHexDigit) &&
            rgb.GetAttributes().All(attribute => attribute.LocalName == "val"))
        {
            kind = $"{prefix}Rgb";
            value = rgbValue.ToUpperInvariant();
            return true;
        }
        return false;
    }
}
