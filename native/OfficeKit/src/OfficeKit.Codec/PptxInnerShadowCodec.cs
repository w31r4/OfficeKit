using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Canonical authored DrawingML inner shadow. Imported effect graphs remain
// outside this slice; the codec only writes the bounded PPJ profile.
internal static class PptxInnerShadowCodec
{
    private const int MaxDirectionAngle60000 = 21_600_000;

    internal static bool TryRead(OpenXmlCompositeElement? properties, out PresentationInnerShadow? shadow)
    {
        shadow = null;
        var lists = properties?.Elements<A.EffectList>().ToArray() ?? [];
        if (lists.Length == 0) return true;
        if (lists.Length != 1) return false;

        var effects = lists[0].ChildElements;
        var innerShadows = effects.OfType<A.InnerShadow>().ToArray();
        var outerShadows = effects.OfType<A.OuterShadow>().ToArray();
        if (innerShadows.Length == 0)
            return effects.Count == 1 && outerShadows.Length == 1 &&
                   PptxShadowCodec.TryReadOuterShadow(outerShadows[0], out _);
        if (innerShadows.Length != 1 || outerShadows.Length > 1 ||
            effects.Count != innerShadows.Length + outerShadows.Length ||
            !ReferenceEquals(effects[0], innerShadows[0]) ||
            outerShadows.Length == 1 && !ReferenceEquals(effects[1], outerShadows[0]) ||
            outerShadows.Length == 1 && !PptxShadowCodec.TryReadOuterShadow(outerShadows[0], out _))
            return false;

        return TryReadInnerShadow(innerShadows[0], out shadow);
    }

    internal static void Apply(OpenXmlCompositeElement properties, PresentationInnerShadow? shadow)
    {
        var effectList = properties.GetFirstChild<A.EffectList>();
        if (shadow is null)
        {
            // A source-bound edit may clear only the recognized inner-shadow
            // owner. Preserve a valid following outer shadow and never touch
            // an unrecognized sibling graph.
            if (effectList is null || !TryRead(properties, out var existing) || existing is null) return;
            effectList.GetFirstChild<A.InnerShadow>()?.Remove();
            if (effectList.ChildElements.Count == 0) effectList.Remove();
            return;
        }
        Validate(shadow, "authored");

        var native = new A.InnerShadow(Color(shadow));
        if (shadow.HasBlurRadiusEmu) native.BlurRadius = checked((int)shadow.BlurRadiusEmu);
        if (shadow.HasDistanceEmu) native.Distance = checked((int)shadow.DistanceEmu);
        if (shadow.HasDirectionAngle60000) native.Direction = checked((int)shadow.DirectionAngle60000);
        if (effectList is null)
        {
            properties.AddChild(new A.EffectList(native), true);
            return;
        }

        if (effectList.ChildElements.Any(child => child is not A.Glow && child is not A.OuterShadow))
            throw new CodecException(
                "unsupported_presentation_effects",
                "Authored inner shadow can only be combined with the bounded glow and outer-shadow effects.");

        effectList.GetFirstChild<A.InnerShadow>()?.Remove();
        var outerShadow = effectList.GetFirstChild<A.OuterShadow>();
        if (outerShadow is not null)
            effectList.InsertBefore(native, outerShadow);
        else if (effectList.GetFirstChild<A.Glow>() is { } glow)
            effectList.InsertAfter(native, glow);
        else
            effectList.Append(native);
    }

    internal static void Validate(PresentationInnerShadow? shadow, string elementId, string subject = "shape")
    {
        if (shadow is null) return;
        var hasRgb = !string.IsNullOrWhiteSpace(shadow.ColorRgb);
        var hasScheme = shadow.HasColorScheme && !string.IsNullOrWhiteSpace(shadow.ColorScheme);
        if (hasRgb == hasScheme)
            throw new CodecException(
                "invalid_presentation_inner_shadow",
                $"Presentation {subject} {elementId} must use exactly one RGB or theme inner-shadow color.");
        if (hasRgb) PptxColor.Normalize(shadow.ColorRgb);
        else PptxColor.NormalizeScheme(shadow.ColorScheme);
        if (shadow.HasBlurRadiusEmu && shadow.BlurRadiusEmu < 0 ||
            shadow.HasDistanceEmu && shadow.DistanceEmu < 0 ||
            shadow.HasDirectionAngle60000 && shadow.DirectionAngle60000 is < 0 or >= 21_600_000 ||
            shadow.HasOpacityThousandthPercent && shadow.OpacityThousandthPercent > 100_000)
            throw new CodecException(
                "invalid_presentation_inner_shadow",
                $"Presentation {subject} {elementId} has invalid inner-shadow geometry or opacity.");
    }

    private static OpenXmlElement Color(PresentationInnerShadow shadow)
    {
        OpenXmlElement color = shadow.HasColorScheme
            ? new A.SchemeColor { Val = PptxColor.SchemeValue(shadow.ColorScheme) }
            : new A.RgbColorModelHex { Val = PptxColor.Normalize(shadow.ColorRgb) };
        if (shadow.HasOpacityThousandthPercent)
            color.Append(new A.Alpha { Val = checked((int)shadow.OpacityThousandthPercent) });
        return color;
    }

    private static bool TryReadInnerShadow(A.InnerShadow source, out PresentationInnerShadow? shadow)
    {
        shadow = null;
        if (!HasOnlyAttributes(source, "blurRad", "dist", "dir") ||
            source.BlurRadius?.Value is < 0 || source.Distance?.Value is < 0 ||
            source.Direction?.Value is < 0 or >= MaxDirectionAngle60000 ||
            source.ChildElements.Count != 1)
            return false;

        var color = source.FirstChild;
        var colorRgb = string.Empty;
        var colorScheme = string.Empty;
        if (color is A.RgbColorModelHex rgb)
        {
            if (!HasOnlyAttributes(rgb, "val") || rgb.Val?.Value is not { Length: 6 } value || !value.All(Uri.IsHexDigit))
                return false;
            colorRgb = PptxColor.Normalize(value);
        }
        else if (color is A.SchemeColor scheme)
        {
            if (!HasOnlyAttributes(scheme, "val") || scheme.Val?.Value is not { } value || !PptxColor.TrySchemeToken(value, out colorScheme))
                return false;
        }
        else return false;

        var alphas = color.Elements<A.Alpha>().ToArray();
        if (color.ChildElements.Count != alphas.Length || alphas.Length > 1 ||
            alphas.Length == 1 && (alphas[0].Val?.Value is not (>= 0 and <= 100_000) || !HasOnlyAttributes(alphas[0], "val")))
            return false;

        var parsed = new PresentationInnerShadow
        {
            ColorRgb = colorRgb,
        };
        if (!string.IsNullOrEmpty(colorScheme)) parsed.ColorScheme = colorScheme;
        if (source.BlurRadius?.Value is { } blur) parsed.BlurRadiusEmu = blur;
        if (source.Distance?.Value is { } distance) parsed.DistanceEmu = distance;
        if (source.Direction?.Value is { } direction) parsed.DirectionAngle60000 = direction;
        if (alphas.SingleOrDefault()?.Val?.Value is { } opacity) parsed.OpacityThousandthPercent = checked((uint)opacity);
        shadow = parsed;
        return true;
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => attribute.NamespaceUri.Length == 0 && allowed.Contains(attribute.LocalName));
    }
}
