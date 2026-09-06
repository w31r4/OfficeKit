using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Canonical authored DrawingML reflection. Imported effect graphs stay
// outside this slice; the codec only writes the bounded PPJ profile.
internal static class PptxReflectionCodec
{
    private const long MaxBlurRadiusEmu = 12_700_000L;
    private const long MaxDistanceEmu = 1_270_000_000L;
    private const int FullReflectionEndPosition = 100_000;
    private const int MaxDirectionAngle60000 = 21_600_000;

    internal static bool TryRead(OpenXmlCompositeElement? properties, out PresentationReflection? reflection)
    {
        reflection = null;
        var lists = properties?.Elements<A.EffectList>().ToArray() ?? [];
        if (lists.Length == 0) return true;
        if (lists.Length != 1) return false;

        var effects = lists[0].ChildElements;
        var reflections = effects.OfType<A.Reflection>().ToArray();
        var outerShadows = effects.OfType<A.OuterShadow>().ToArray();
        if (reflections.Length == 0)
            return effects.Count == 1 && outerShadows.Length == 1 &&
                   PptxShadowCodec.TryReadOuterShadow(outerShadows[0], out _);
        if (reflections.Length != 1 || outerShadows.Length > 1 ||
            effects.Count != reflections.Length + outerShadows.Length ||
            !ReferenceEquals(effects[effects.Count - 1], reflections[0]) ||
            outerShadows.Length == 1 &&
                (!ReferenceEquals(effects[0], outerShadows[0]) ||
                 !PptxShadowCodec.TryReadOuterShadow(outerShadows[0], out _)))
            return false;

        return TryReadDirectReflection(reflections[0], out reflection);
    }

    internal static void Apply(OpenXmlCompositeElement properties, PresentationReflection? reflection)
    {
        var effectList = properties.GetFirstChild<A.EffectList>();
        if (reflection is null)
        {
            // A source-bound edit may clear only the recognized reflection
            // owner. Preserve a valid preceding outer shadow and never touch
            // an unrecognized sibling graph.
            if (effectList is null || !TryRead(properties, out var existing) || existing is null) return;
            effectList.GetFirstChild<A.Reflection>()?.Remove();
            if (effectList.ChildElements.Count == 0) effectList.Remove();
            return;
        }
        Validate(reflection, "authored");

        var native = new A.Reflection
        {
            BlurRadius = checked((long)reflection.BlurRadiusEmu),
            StartOpacity = checked((int)reflection.StartOpacityThousandthPercent),
            StartPosition = 0,
            EndAlpha = checked((int)reflection.EndOpacityThousandthPercent),
            EndPosition = FullReflectionEndPosition,
            Distance = checked((long)reflection.DistanceEmu),
            Direction = checked((int)reflection.DirectionAngle60000),
        };
        if (effectList is null)
        {
            properties.AddChild(new A.EffectList(native), true);
            return;
        }

        if (effectList.ChildElements.Any(child => child is not A.Glow && child is not A.InnerShadow &&
                                                  child is not A.OuterShadow && child is not A.SoftEdge))
            throw new CodecException(
                "unsupported_presentation_effects",
                "Authored reflection can only be combined with the bounded glow, inner-shadow, outer-shadow, and soft-edge effects.");

        effectList.GetFirstChild<A.Reflection>()?.Remove();
        var softEdge = effectList.GetFirstChild<A.SoftEdge>();
        if (softEdge is not null)
            effectList.InsertBefore(native, softEdge);
        else
            effectList.Append(native);
    }

    internal static void Validate(PresentationReflection? reflection, string elementId, string subject = "shape")
    {
        if (reflection is null) return;
        if (!reflection.HasBlurRadiusEmu || reflection.BlurRadiusEmu is < 0 or > MaxBlurRadiusEmu ||
            !reflection.HasStartOpacityThousandthPercent || reflection.StartOpacityThousandthPercent > 100_000 ||
            !reflection.HasEndOpacityThousandthPercent || reflection.EndOpacityThousandthPercent > 100_000 ||
            !reflection.HasDistanceEmu || reflection.DistanceEmu is < 0 or > MaxDistanceEmu ||
            !reflection.HasDirectionAngle60000 || reflection.DirectionAngle60000 is < 0 or >= 21_600_000)
            throw new CodecException(
                "invalid_presentation_reflection",
                $"Presentation {subject} {elementId} has invalid reflection geometry or opacity.");
    }

    internal static bool TryReadDirectReflection(A.Reflection source, out PresentationReflection? reflection)
    {
        reflection = null;
        if (!HasOnlyAttributes(source, "blurRad", "stA", "stPos", "endA", "endPos", "dist", "dir") ||
            source.ChildElements.Count != 0 ||
            source.StartPosition?.Value is not 0 ||
            source.EndPosition?.Value is not FullReflectionEndPosition ||
            source.BlurRadius?.Value is < 0 or > MaxBlurRadiusEmu ||
            source.StartOpacity?.Value is < 0 or > 100_000 ||
            source.EndAlpha?.Value is < 0 or > 100_000 ||
            source.Distance?.Value is < 0 or > MaxDistanceEmu ||
            source.Direction?.Value is < 0 or >= MaxDirectionAngle60000)
            return false;

        reflection = new PresentationReflection();
        if (source.BlurRadius?.Value is { } blur) reflection.BlurRadiusEmu = blur;
        if (source.StartOpacity?.Value is { } startOpacity) reflection.StartOpacityThousandthPercent = checked((uint)startOpacity);
        if (source.EndAlpha?.Value is { } endOpacity) reflection.EndOpacityThousandthPercent = checked((uint)endOpacity);
        if (source.Distance?.Value is { } distance) reflection.DistanceEmu = distance;
        if (source.Direction?.Value is { } direction) reflection.DirectionAngle60000 = direction;
        return true;
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => attribute.NamespaceUri.Length == 0 && allowed.Contains(attribute.LocalName));
    }
}
