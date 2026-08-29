using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeKit.Codec;

internal sealed record PptxNativeStyleLeaf(uint Index, string Kind, string Value, P.Shape Shape);

// Describes only the smallest style surface that can be proven without
// rebuilding a source-bound group.  The group remains source-bound; this codec
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
        var lineWidthLeaves = new List<(string Kind, string Value, P.Shape Shape)>();
        var lineStyleLeaves = new List<(string Kind, string Value, P.Shape Shape)>();
        var lineCapLeaves = new List<(string Kind, string Value, P.Shape Shape)>();
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
            if (TryReadWidth(outlines[0], out var widthValue))
                lineWidthLeaves.Add(("lineWidthEmu", widthValue, shape));
            if (TryReadStyle(outlines[0], out var styleValue))
                lineStyleLeaves.Add(("lineStyle", styleValue, shape));
            if (TryReadCap(outlines[0], out var capValue))
                lineCapLeaves.Add(("lineCap", capValue, shape));
            var lineFills = outlines[0].ChildElements
                .Where(child => child is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill)
                .ToArray();
            if (lineFills.Length == 1 && lineFills[0] is A.SolidFill lineSolid && TryReadColor(lineSolid, "line", out var lineKind, out var lineValue))
                lineLeaves.Add((lineKind, lineValue, shape));
        }

        // Keep existing fill/line color/width indexes stable; append the dash
        // family so adding this capability cannot retarget a prior
        // source-bound leaf.
        var described = fillLeaves.Concat(lineLeaves).Concat(lineWidthLeaves).Concat(lineStyleLeaves).Concat(lineCapLeaves).ToArray();
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

    private static bool TryReadWidth(A.Outline outline, out string value)
    {
        value = string.Empty;
        var attributes = outline.GetAttributes()
            .Where(attribute => attribute.LocalName == "w")
            .ToArray();
        if (attributes.Length != 1 || !ulong.TryParse(
                attributes[0].Value,
                System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture,
                out var width) || width > 20_116_800)
            return false;
        // Require the source token to be canonical so expectedHash and the
        // token splice remain deterministic across import/round-trip.
        var canonical = width.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (!string.Equals(attributes[0].Value, canonical, StringComparison.Ordinal) || width == 0) return false;
        value = canonical;
        return true;
    }

    private static bool TryReadStyle(A.Outline outline, out string value)
    {
        value = string.Empty;
        var dashes = outline.Elements<A.PresetDash>().ToArray();
        if (dashes.Length != 1 || !PptxLineStyleCodec.TryReadPresetDash(dashes[0], out var style)) return false;
        // A dash token on a no-fill outline has no useful visible editing
        // meaning; only expose it when the outline has one simple solid paint.
        var fills = outline.ChildElements.Where(child => child is A.NoFill or A.SolidFill).ToArray();
        if (fills.Length != 1 || fills[0] is not A.SolidFill solid || solid.ChildElements.Count != 1) return false;
        if (solid.FirstChild is A.RgbColorModelHex rgb)
        {
            if (rgb.Val?.Value is not { Length: 6 } token || !token.All(Uri.IsHexDigit) ||
                rgb.ChildElements.Any() || !rgb.GetAttributes().All(attribute => attribute.LocalName == "val")) return false;
        }
        else if (solid.FirstChild is A.SchemeColor scheme)
        {
            if (scheme.Val?.Value is not { } token || !PptxColor.TrySchemeToken(token, out _) ||
                scheme.ChildElements.Any() || !scheme.GetAttributes().All(attribute => attribute.LocalName == "val")) return false;
        }
        else return false;
        value = style;
        return true;
    }

    private static bool TryReadCap(A.Outline outline, out string value)
    {
        value = string.Empty;
        var attributes = outline.GetAttributes()
            .Where(attribute => attribute.LocalName == "cap")
            .ToArray();
        if (attributes.Length != 1 || !outline.GetAttributes().All(attribute => attribute.LocalName is "w" or "cap" or "cmpd" or "algn")) return false;
        if (!PptxLineStyleCodec.TryReadCapValue(attributes[0].Value, out var cap)) return false;

        // A cap token is useful only when the outline paint is a simple
        // explicit solid color. Preserve effect-bearing and inherited line
        // graphs as opaque rather than exposing a misleading partial style.
        var fills = outline.ChildElements.Where(child => child is A.NoFill or A.SolidFill).ToArray();
        if (fills.Length != 1 || fills[0] is not A.SolidFill solid || solid.ChildElements.Count != 1)
            return false;
        if (solid.FirstChild is A.RgbColorModelHex rgb)
        {
            if (rgb.Val?.Value is not { Length: 6 } token || !token.All(Uri.IsHexDigit) ||
                rgb.ChildElements.Any() || !rgb.GetAttributes().All(attribute => attribute.LocalName == "val"))
                return false;
        }
        else if (solid.FirstChild is A.SchemeColor scheme)
        {
            if (scheme.Val?.Value is not { } token || !PptxColor.TrySchemeToken(token, out _) ||
                scheme.ChildElements.Any() || !scheme.GetAttributes().All(attribute => attribute.LocalName == "val"))
                return false;
        }
        else return false;
        value = cap;
        return true;
    }
}
