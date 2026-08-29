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

        var described = new List<PptxNativeStyleLeaf>();
        foreach (var shape in group.Descendants<P.Shape>())
        {
            var properties = shape.ShapeProperties;
            if (properties is null) continue;
            var fills = properties.ChildElements
                .Where(child => child is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill)
                .ToArray();
            if (fills.Length != 1 || fills[0] is not A.SolidFill solid) continue;
            if (solid.ChildElements.Count != 1 || solid.FirstChild is null || solid.FirstChild.ChildElements.Count != 0) continue;

            if (solid.FirstChild is A.SchemeColor scheme &&
                scheme.Val?.Value is { } schemeValue &&
                scheme.GetAttributes().All(attribute => attribute.LocalName == "val") &&
                PptxColor.TrySchemeToken(schemeValue, out var schemeToken))
            {
                described.Add(new PptxNativeStyleLeaf(checked((uint)described.Count), "fillScheme", schemeToken, shape));
            }
            else if (solid.FirstChild is A.RgbColorModelHex rgb &&
                     rgb.Val?.Value is { Length: 6 } rgbValue && rgbValue.All(Uri.IsHexDigit) &&
                     rgb.GetAttributes().All(attribute => attribute.LocalName == "val"))
            {
                described.Add(new PptxNativeStyleLeaf(checked((uint)described.Count), "fillRgb", rgbValue.ToUpperInvariant(), shape));
            }

            if (described.Count > MaxLeaves) return false;
        }

        if (described.Count == 0) return false;
        leaves = described;
        return true;
    }

    internal static bool TryResolve(OpenXmlElement source, uint index, out PptxNativeStyleLeaf leaf)
    {
        leaf = null!;
        if (!TryDescribe(source, out var leaves) || index >= (uint)leaves.Count) return false;
        leaf = leaves[(int)index];
        return true;
    }
}
