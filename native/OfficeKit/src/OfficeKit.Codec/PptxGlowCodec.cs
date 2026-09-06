using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Canonical DrawingML glow. Imported graphs are still bounded: this codec can
// read one direct glow, optionally followed by one proven outer shadow, while
// all other effect-list topologies remain source-owned.
internal static class PptxGlowCodec
{
    private const long MaxRadiusEmu = 12_700_000L;

    internal static bool TryRead(OpenXmlCompositeElement? properties, out PresentationGlow? glow)
    {
        glow = null;
        var lists = properties?.Elements<A.EffectList>().ToArray() ?? [];
        if (lists.Length == 0) return true;
        if (lists.Length != 1) return false;

        var effects = lists[0].ChildElements;
        var glowElements = effects.OfType<A.Glow>().ToArray();
        var shadowElements = effects.OfType<A.OuterShadow>().ToArray();
        if (glowElements.Length == 0)
            return effects.Count == 1 && shadowElements.Length == 1 &&
                   PptxShadowCodec.TryReadOuterShadow(shadowElements[0], out _);
        if (glowElements.Length != 1 || shadowElements.Length > 1 ||
            effects.Count != glowElements.Length + shadowElements.Length ||
            !ReferenceEquals(effects[0], glowElements[0]) ||
            shadowElements.Length == 1 && !ReferenceEquals(effects[1], shadowElements[0]) ||
            shadowElements.Length == 1 && !PptxShadowCodec.TryReadOuterShadow(shadowElements[0], out _))
            return false;

        return TryReadGlow(glowElements[0], out glow);
    }

    internal static void Apply(OpenXmlCompositeElement properties, PresentationGlow? glow)
    {
        var effectList = properties.GetFirstChild<A.EffectList>();
        if (glow is null)
        {
            // A source-bound text edit may clear only the recognized glow
            // owner. Never remove an unrecognized sibling graph.
            if (effectList is null || !TryRead(properties, out var existing) || existing is null) return;
            effectList.GetFirstChild<A.Glow>()?.Remove();
            if (effectList.ChildElements.Count == 0) effectList.Remove();
            return;
        }
        Validate(glow, "authored");

        var native = new A.Glow(Color(glow))
        {
            Radius = checked((uint)glow.RadiusEmu),
        };
        if (effectList is null)
        {
            properties.AddChild(new A.EffectList(native), true);
            return;
        }

        if (!TryRead(properties, out _) || effectList.ChildElements.Any(child => child is not A.Glow and not A.OuterShadow))
            throw new CodecException(
                "unsupported_presentation_effects",
                "Glow can only be combined with a proven bounded outer-shadow effect.");

        effectList.GetFirstChild<A.Glow>()?.Remove();
        var outerShadow = effectList.GetFirstChild<A.OuterShadow>();
        if (outerShadow is not null)
            effectList.InsertBefore(native, outerShadow);
        else
            effectList.Append(native);
    }

    internal static void Validate(PresentationGlow? glow, string elementId, string subject = "shape")
    {
        if (glow is null) return;
        var hasRgb = !string.IsNullOrWhiteSpace(glow.ColorRgb);
        var hasScheme = glow.HasColorScheme && !string.IsNullOrWhiteSpace(glow.ColorScheme);
        if (hasRgb == hasScheme)
            throw new CodecException(
                "invalid_presentation_glow",
                $"Presentation {subject} {elementId} must use exactly one RGB or theme glow color.");
        if (hasRgb) PptxColor.Normalize(glow.ColorRgb);
        else PptxColor.NormalizeScheme(glow.ColorScheme);
        if (!glow.HasRadiusEmu || glow.RadiusEmu is < 0 or > MaxRadiusEmu ||
            glow.HasOpacityThousandthPercent && glow.OpacityThousandthPercent > 100_000)
            throw new CodecException(
                "invalid_presentation_glow",
                $"Presentation {subject} {elementId} has invalid glow radius or opacity.");
    }

    private static bool TryReadGlow(A.Glow source, out PresentationGlow? glow)
    {
        glow = null;
        if (!HasOnlyAttributes(source, "rad") || source.Radius?.Value is not { } radius || radius > MaxRadiusEmu ||
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

        var parsed = new PresentationGlow
        {
            ColorRgb = colorRgb,
            RadiusEmu = radius,
        };
        if (!string.IsNullOrEmpty(colorScheme)) parsed.ColorScheme = colorScheme;
        if (alphas.SingleOrDefault()?.Val?.Value is { } opacity) parsed.OpacityThousandthPercent = checked((uint)opacity);
        glow = parsed;
        return true;
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => attribute.NamespaceUri.Length == 0 && allowed.Contains(attribute.LocalName));
    }

    private static OpenXmlElement Color(PresentationGlow glow)
    {
        OpenXmlElement color = glow.HasColorScheme
            ? new A.SchemeColor { Val = PptxColor.SchemeValue(glow.ColorScheme) }
            : new A.RgbColorModelHex { Val = PptxColor.Normalize(glow.ColorRgb) };
        if (glow.HasOpacityThousandthPercent)
            color.Append(new A.Alpha { Val = checked((int)glow.OpacityThousandthPercent) });
        return color;
    }
}
