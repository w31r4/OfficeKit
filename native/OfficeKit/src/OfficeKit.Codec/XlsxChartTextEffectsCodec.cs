using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeKit.Codec;

// Own the whole bounded chart character effect list; each modeled effect has
// independent presence, while unknown descendants retain source ownership.
internal static class XlsxChartTextEffectsCodec
{
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static bool TryRead(XElement effects, out PresentationGlow? glow, out PresentationShadow? shadow)
    {
        glow = null;
        shadow = null;
        if (effects.Name != DrawingNs + "effectLst" ||
            effects.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
            effects.DescendantNodes().Any(node => node is not XElement && (node is not XText text || !string.IsNullOrWhiteSpace(text.Value))) ||
            effects.Descendants(DrawingNs + "alpha").Any(alpha => alpha.HasElements)) return false;
        try
        {
            var properties = new A.RunProperties(new A.EffectList(effects.ToString(SaveOptions.DisableFormatting)));
            // The existing typed glow reader proves exact order/count and all
            // known sibling/color attributes, including the shadow if present.
            if (!PptxGlowCodec.TryRead(properties, out glow)) return false;
            if (effects.Element(DrawingNs + "outerShdw") is { } outer &&
                !PptxShadowCodec.TryReadTextEffects(new XElement(DrawingNs + "effectLst", new XElement(outer)), out shadow)) return false;
            return glow is not null || shadow is not null;
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
        return XElement.Parse(properties.GetFirstChild<A.EffectList>()!.OuterXml);
    }

    internal static string Semantics(SpreadsheetChartTextStyleArtifact style) =>
        style.Glow is null && style.Shadow is null ? "default-effects" : Element(style).ToString(SaveOptions.DisableFormatting);
}
