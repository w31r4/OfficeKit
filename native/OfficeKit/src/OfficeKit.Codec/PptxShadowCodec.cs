using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Canonical direct DrawingML outer-shadow profile shared by ordinary shapes
// and pictures. More complex effect graphs remain source-owned and fail closed.
internal static class PptxShadowCodec
{
    internal static bool TryRead(OpenXmlCompositeElement? properties, out PresentationShadow? shadow)
    {
        shadow = null;
        var lists = properties?.Elements<A.EffectList>().ToArray() ?? [];
        if (lists.Length == 0) return true;
        if (lists.Length != 1 || lists[0].ChildElements.Count != 1 || lists[0].FirstChild is not A.OuterShadow outer ||
            !HasOnlyAttributes(outer, "blurRad", "dist", "dir") || outer.BlurRadius?.Value is < 0 || outer.Distance?.Value is < 0 ||
            outer.Direction?.Value is < 0 or >= 21_600_000 || outer.ChildElements.Count != 1 || outer.FirstChild is not A.RgbColorModelHex color ||
            color.Val?.Value is not { Length: 6 } rgb || !rgb.All(Uri.IsHexDigit) || !HasOnlyAttributes(color, "val")) return false;
        var alphas = color.Elements<A.Alpha>().ToArray();
        if (color.ChildElements.Count != alphas.Length || alphas.Length > 1 ||
            alphas.Length == 1 && (alphas[0].Val?.Value is not (>= 0 and <= 100_000) || !HasOnlyAttributes(alphas[0], "val")))
            return false;
        shadow = new PresentationShadow
        {
            ColorRgb = PptxColor.Normalize(rgb),
            BlurRadiusEmu = outer.BlurRadius?.Value ?? 0L,
            DistanceEmu = outer.Distance?.Value ?? 0L,
            DirectionAngle60000 = outer.Direction?.Value ?? 0,
            OpacityThousandthPercent = checked((uint)(alphas.SingleOrDefault()?.Val?.Value ?? 100_000)),
        };
        return true;
    }

    internal static void Apply(OpenXmlCompositeElement properties, PresentationShadow? shadow)
    {
        properties.GetFirstChild<A.EffectList>()?.Remove();
        if (shadow is null) return;
        var color = new A.RgbColorModelHex { Val = PptxColor.Normalize(shadow.ColorRgb) };
        color.Append(new A.Alpha { Val = checked((int)shadow.OpacityThousandthPercent) });
        var outer = new A.OuterShadow(color)
        {
            BlurRadius = shadow.BlurRadiusEmu,
            Distance = shadow.DistanceEmu,
            Direction = shadow.DirectionAngle60000,
        };
        properties.AddChild(new A.EffectList(outer), true);
    }

    internal static void Validate(PresentationShadow? shadow, string elementId, string subject = "shape")
    {
        if (shadow is null) return;
        PptxColor.Normalize(shadow.ColorRgb);
        if (shadow.BlurRadiusEmu < 0 || shadow.DistanceEmu < 0 || shadow.DirectionAngle60000 is < 0 or >= 21_600_000 || shadow.OpacityThousandthPercent > 100_000)
            throw new CodecException("invalid_presentation_shadow", $"Presentation {subject} {elementId} has invalid shadow geometry or opacity.");
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => attribute.NamespaceUri.Length == 0 && allowed.Contains(attribute.LocalName));
    }
}
