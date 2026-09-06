using DocumentFormat.OpenXml;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Canonical authored DrawingML soft edge plus the strict direct source-bound
// owner. Broader imported effect graphs stay outside this slice; the codec
// only writes the bounded PPJ radius profile alongside the direct effects.
internal static class PptxSoftEdgeCodec
{
    private const long MaxRadiusEmu = 12_700_000L;

    internal static bool TryRead(OpenXmlCompositeElement? properties, out PresentationSoftEdge? softEdge)
    {
        softEdge = null;
        var lists = properties?.Elements<A.EffectList>().ToArray() ?? [];
        if (lists.Length == 0) return true;
        if (lists.Length != 1) return false;

        var effects = lists[0].ChildElements;
        var softEdges = effects.OfType<A.SoftEdge>().ToArray();
        if (softEdges.Length == 0)
        {
            if (effects.Count == 1 && effects[0] is A.OuterShadow outer &&
                PptxShadowCodec.TryReadOuterShadow(outer, out _))
                return true;
            return PptxReflectionCodec.TryRead(properties, out var reflection) && reflection is not null;
        }

        if (softEdges.Length != 1 || effects.Count > 2 ||
            !ReferenceEquals(effects[effects.Count - 1], softEdges[0]))
            return false;
        if (effects.Count == 2 && effects[0] is A.OuterShadow outerShadow)
        {
            if (!PptxShadowCodec.TryReadOuterShadow(outerShadow, out _)) return false;
        }
        else if (effects.Count == 2 && effects[0] is A.Reflection reflection)
        {
            if (!PptxReflectionCodec.TryReadDirectReflection(reflection, out _)) return false;
        }
        else if (effects.Count != 1)
            return false;

        return TryReadSoftEdge(softEdges[0], out softEdge);
    }

    internal static void Apply(OpenXmlCompositeElement properties, PresentationSoftEdge? softEdge)
    {
        var effectList = properties.GetFirstChild<A.EffectList>();
        if (softEdge is null)
        {
            // A source-bound edit may clear only the recognized soft-edge
            // owner. Preserve a valid preceding outer shadow/reflection.
            if (effectList is null || !TryRead(properties, out var existing) || existing is null) return;
            effectList.GetFirstChild<A.SoftEdge>()?.Remove();
            if (effectList.ChildElements.Count == 0) effectList.Remove();
            return;
        }
        Validate(softEdge, "authored");

        var native = new A.SoftEdge
        {
            Radius = checked((uint)softEdge.RadiusEmu),
        };
        if (effectList is null)
        {
            properties.AddChild(new A.EffectList(native), true);
            return;
        }

        if (effectList.ChildElements.Any(child => child is not A.Glow && child is not A.InnerShadow &&
                                                  child is not A.OuterShadow && child is not A.Reflection))
            throw new CodecException(
                "unsupported_presentation_effects",
                "Authored soft edge can only be combined with the bounded glow, inner-shadow, outer-shadow, and reflection effects.");

        effectList.GetFirstChild<A.SoftEdge>()?.Remove();
        var reflection = effectList.GetFirstChild<A.Reflection>();
        if (reflection is not null)
            effectList.InsertAfter(native, reflection);
        else if (effectList.GetFirstChild<A.OuterShadow>() is { } outerShadow)
            effectList.InsertAfter(native, outerShadow);
        else
            effectList.Append(native);
    }

    internal static void Validate(PresentationSoftEdge? softEdge, string elementId, string subject = "shape")
    {
        if (softEdge is null) return;
        if (!softEdge.HasRadiusEmu || softEdge.RadiusEmu is < 0 or > MaxRadiusEmu)
            throw new CodecException(
                "invalid_presentation_soft_edge",
                $"Presentation {subject} {elementId} has invalid soft-edge radius.");
    }

    private static bool TryReadSoftEdge(A.SoftEdge source, out PresentationSoftEdge? softEdge)
    {
        softEdge = null;
        if (!HasOnlyAttributes(source, "rad") || source.ChildElements.Count != 0 ||
            source.Radius?.Value is not { } radius || radius > MaxRadiusEmu)
            return false;
        softEdge = new PresentationSoftEdge { RadiusEmu = radius };
        return true;
    }

    private static bool HasOnlyAttributes(OpenXmlElement element, params string[] names)
    {
        var allowed = names.ToHashSet(StringComparer.Ordinal);
        return element.GetAttributes().All(attribute => attribute.NamespaceUri.Length == 0 && allowed.Contains(attribute.LocalName));
    }
}
