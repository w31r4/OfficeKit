using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Canonical direct DrawingML outer-shadow profile shared by ordinary shapes
// and pictures. More complex effect graphs remain source-owned and fail closed.
internal static class PptxShadowCodec
{
    private static readonly IReadOnlySet<string> Alignments = new HashSet<string>(StringComparer.Ordinal)
    {
        "tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br",
    };

    internal static bool TryRead(OpenXmlCompositeElement? properties, out PresentationShadow? shadow)
    {
        shadow = null;
        var lists = properties?.Elements<A.EffectList>().ToArray() ?? [];
        if (lists.Length == 0) return true;
        if (lists.Length != 1 || lists[0].ChildElements.Count != 1 || lists[0].FirstChild is not A.OuterShadow outer ||
            !TryReadOuterShadow(outer, out shadow)) return false;
        return true;
    }

    // Effect-list owners such as text glow may need to prove an outer shadow
    // sibling without treating the whole list as a shadow-only graph. Keep
    // that proof separate so shape/image classification remains strict.
    internal static bool TryReadOuterShadow(A.OuterShadow outer, out PresentationShadow? shadow)
    {
        shadow = null;
        if (!HasOnlyAttributes(outer, "blurRad", "dist", "dir", "algn", "rotWithShape") || outer.BlurRadius?.Value is < 0 || outer.Distance?.Value is < 0 ||
            outer.Direction?.Value is < 0 or >= 21_600_000 || outer.Alignment?.Value is { } alignment && !Alignments.Contains(AlignmentName(alignment)) ||
            outer.ChildElements.Count != 1) return false;

        var color = outer.FirstChild;
        var colorRgb = string.Empty;
        var colorScheme = string.Empty;
        if (color is A.RgbColorModelHex rgb)
        {
            if (rgb.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit) || !HasOnlyAttributes(rgb, "val")) return false;
            colorRgb = PptxColor.Normalize(value);
        }
        else if (color is A.SchemeColor scheme)
        {
            if (!HasOnlyAttributes(scheme, "val") || scheme.Val?.Value is not { } value || !PptxColor.TrySchemeToken(value, out colorScheme)) return false;
        }
        else return false;

        var alphas = color.Elements<A.Alpha>().ToArray();
        if (color.ChildElements.Count != alphas.Length || alphas.Length > 1 ||
            alphas.Length == 1 && (alphas[0].Val?.Value is not (>= 0 and <= 100_000) || !HasOnlyAttributes(alphas[0], "val")))
            return false;

        var parsed = new PresentationShadow
        {
            ColorRgb = colorRgb,
        };
        if (!string.IsNullOrEmpty(colorScheme)) parsed.ColorScheme = colorScheme;
        if (outer.BlurRadius?.Value is { } blur) parsed.BlurRadiusEmu = blur;
        if (outer.Distance?.Value is { } distance) parsed.DistanceEmu = distance;
        if (outer.Direction?.Value is { } direction) parsed.DirectionAngle60000 = direction;
        if (outer.Alignment?.Value is { } parsedAlignment) parsed.Alignment = AlignmentName(parsedAlignment);
        if (outer.RotateWithShape?.Value is { } rotateWithShape) parsed.RotateWithShape = rotateWithShape;
        if (alphas.SingleOrDefault()?.Val?.Value is { } opacity) parsed.OpacityThousandthPercent = checked((uint)opacity);
        shadow = parsed;
        return true;
    }

    internal static void Apply(OpenXmlCompositeElement properties, PresentationShadow? shadow)
    {
        properties.GetFirstChild<A.EffectList>()?.Remove();
        if (shadow is null) return;
        OpenXmlElement color = !string.IsNullOrEmpty(shadow.ColorScheme)
            ? new A.SchemeColor { Val = PptxColor.SchemeValue(shadow.ColorScheme) }
            : new A.RgbColorModelHex { Val = PptxColor.Normalize(shadow.ColorRgb) };
        if (shadow.HasOpacityThousandthPercent) color.Append(new A.Alpha { Val = checked((int)shadow.OpacityThousandthPercent) });
        var outer = new A.OuterShadow(color);
        if (shadow.HasBlurRadiusEmu) outer.BlurRadius = shadow.BlurRadiusEmu;
        if (shadow.HasDistanceEmu) outer.Distance = shadow.DistanceEmu;
        if (shadow.HasDirectionAngle60000) outer.Direction = shadow.DirectionAngle60000;
        if (shadow.HasAlignment) outer.Alignment = AlignmentValue(shadow.Alignment);
        if (shadow.HasRotateWithShape) outer.RotateWithShape = shadow.RotateWithShape;
        properties.AddChild(new A.EffectList(outer), true);
    }

    internal static void Validate(PresentationShadow? shadow, string elementId, string subject = "shape")
    {
        if (shadow is null) return;
        var hasRgb = !string.IsNullOrWhiteSpace(shadow.ColorRgb);
        var hasScheme = shadow.HasColorScheme && !string.IsNullOrWhiteSpace(shadow.ColorScheme);
        if (hasRgb == hasScheme)
            throw new CodecException("invalid_presentation_shadow", $"Presentation {subject} {elementId} must use exactly one RGB or theme shadow color.");
        if (hasRgb) PptxColor.Normalize(shadow.ColorRgb);
        else PptxColor.NormalizeScheme(shadow.ColorScheme);
        if (shadow.HasBlurRadiusEmu && shadow.BlurRadiusEmu < 0 || shadow.HasDistanceEmu && shadow.DistanceEmu < 0 ||
            shadow.HasDirectionAngle60000 && shadow.DirectionAngle60000 is < 0 or >= 21_600_000 ||
            shadow.HasOpacityThousandthPercent && shadow.OpacityThousandthPercent > 100_000 ||
            shadow.HasAlignment && !Alignments.Contains(shadow.Alignment))
            throw new CodecException("invalid_presentation_shadow", $"Presentation {subject} {elementId} has invalid shadow geometry or opacity.");
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => attribute.NamespaceUri.Length == 0 && allowed.Contains(attribute.LocalName));
    }

    private static string AlignmentName(A.RectangleAlignmentValues value) =>
        value == A.RectangleAlignmentValues.TopLeft ? "tl" :
        value == A.RectangleAlignmentValues.Top ? "t" :
        value == A.RectangleAlignmentValues.TopRight ? "tr" :
        value == A.RectangleAlignmentValues.Left ? "l" :
        value == A.RectangleAlignmentValues.Center ? "ctr" :
        value == A.RectangleAlignmentValues.Right ? "r" :
        value == A.RectangleAlignmentValues.BottomLeft ? "bl" :
        value == A.RectangleAlignmentValues.Bottom ? "b" :
        value == A.RectangleAlignmentValues.BottomRight ? "br" : string.Empty;

    private static A.RectangleAlignmentValues AlignmentValue(string value) => value switch
    {
        "tl" => A.RectangleAlignmentValues.TopLeft,
        "t" => A.RectangleAlignmentValues.Top,
        "tr" => A.RectangleAlignmentValues.TopRight,
        "l" => A.RectangleAlignmentValues.Left,
        "ctr" => A.RectangleAlignmentValues.Center,
        "r" => A.RectangleAlignmentValues.Right,
        "bl" => A.RectangleAlignmentValues.BottomLeft,
        "b" => A.RectangleAlignmentValues.Bottom,
        "br" => A.RectangleAlignmentValues.BottomRight,
        _ => throw new CodecException("invalid_presentation_shadow", $"Unsupported Presentation shadow alignment {value}."),
    };
}
