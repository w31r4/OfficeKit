using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Own the whole bounded chart character effect list; each modeled effect has
// independent presence, while unknown descendants retain source ownership.
internal static class XlsxChartTextEffectsCodec
{
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static bool TryRead(XElement effects, out PresentationGlow? glow, out PresentationInnerShadow? innerShadow, out PresentationShadow? shadow, out PresentationSoftEdge? softEdge, out PresentationReflection? reflection)
    {
        reflection = null;
        glow = null;
        innerShadow = null;
        shadow = null;
        softEdge = null;
        if (effects.Name != DrawingNs + "effectLst" ||
            effects.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
            effects.DescendantNodes().Any(node => node is not XElement && (node is not XText text || !string.IsNullOrWhiteSpace(text.Value))) ||
            effects.Descendants(DrawingNs + "alpha").Any(alpha => alpha.HasElements)) return false;
        try
        {
            var previousRank = -1;
            foreach (var child in effects.Elements())
            {
                if (child.Name.Namespace != DrawingNs) return false;
                var rank = child.Name.LocalName switch
                {
                    "glow" => 0, "innerShdw" => 1, "outerShdw" => 2, "reflection" => 3, "softEdge" => 4, _ => -1,
                };
                if (rank <= previousRank) return false;
                previousRank = rank;
                var isolated = new XElement(DrawingNs + "effectLst", new XElement(child));
                var properties = new A.RunProperties(new A.EffectList(isolated.ToString(SaveOptions.DisableFormatting)));
                switch (rank)
                {
                    case 0:
                        if (!PptxGlowCodec.TryRead(properties, out glow) || glow is null) return false;
                        break;
                    case 1:
                        if (!PptxInnerShadowCodec.TryRead(properties, out innerShadow) || innerShadow is null ||
                            innerShadow.HasBlurRadiusEmu && innerShadow.BlurRadiusEmu > 12_700_000 ||
                            innerShadow.HasDistanceEmu && innerShadow.DistanceEmu > 1_270_000_000) return false;
                        break;
                    case 2:
                        if (!PptxShadowCodec.TryReadTextEffects(isolated, out shadow) || shadow is null) return false;
                        break;
                    case 3:
                        if (child.HasElements || properties.GetFirstChild<A.EffectList>()?.GetFirstChild<A.Reflection>() is not { } nativeReflection ||
                            !PptxReflectionCodec.TryReadDirectReflection(nativeReflection, out reflection, allowVariablePositions: true)) return false;
                        break;
                    case 4:
                        if (child.HasElements || properties.GetFirstChild<A.EffectList>()?.GetFirstChild<A.SoftEdge>() is not { } native ||
                            !PptxSoftEdgeValueCodec.TryRead(native, out softEdge)) return false;
                        break;
                }
            }
            return previousRank >= 0;
        }
        catch (Exception error) when (error is FormatException or OverflowException or ArgumentException or InvalidOperationException)
        {
            return false;
        }
    }

    internal static XElement Element(SpreadsheetChartTextStyleArtifact style)
    {
        var properties = new A.RunProperties();
        PptxShadowCodec.Validate(style.Shadow, "chart-text", "chart text");
        PptxShadowCodec.Apply(properties, style.Shadow);
        PptxGlowCodec.Apply(properties, style.Glow);
        PptxInnerShadowCodec.Apply(properties, style.InnerShadow);
        PptxReflectionCodec.Apply(properties, style.Reflection);
        if (style.SoftEdge is not null)
        {
            PptxSoftEdgeValueCodec.Validate(style.SoftEdge, "chart-text", "chart text");
            var native = new A.SoftEdge { Radius = checked((uint)style.SoftEdge.RadiusEmu) };
            if (properties.GetFirstChild<A.EffectList>() is { } effects) effects.Append(native);
            else properties.AddChild(new A.EffectList(native), true);
        }
        return XElement.Parse(properties.GetFirstChild<A.EffectList>()!.OuterXml);
    }

    internal static string Semantics(SpreadsheetChartTextStyleArtifact style) =>
        style.Reflection is null && style.InnerShadow is null && style.SoftEdge is null && style.Glow is null && style.Shadow is null ? "default-effects" : Element(style).ToString(SaveOptions.DisableFormatting);
}
