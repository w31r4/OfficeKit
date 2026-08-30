using System.Globalization;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns the bounded chart-area and plot-area paint profile shared by PPTX and
// XLSX ChartSpace: direct no-fill or direct sRGB solid fill with optional alpha.
internal static class XlsxChartSurfaceFillCodec
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static void Validate(SpreadsheetChartSurfaceFill? fill, string subject)
    {
        if (fill is null) return;
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.None)
            throw new CodecException("invalid_chart_fill", $"{subject} must select no_fill or solid_rgb.");
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.NoFill && !fill.NoFill)
            throw new CodecException("invalid_chart_fill", $"{subject} no_fill must be true.");
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.SolidRgb)
            _ = PptxColor.Normalize(fill.SolidRgb);
        if (fill.HasOpacityThousandthPercent &&
            (fill.FillCase != SpreadsheetChartSurfaceFill.FillOneofCase.SolidRgb || fill.OpacityThousandthPercent > 100_000))
            throw new CodecException("invalid_chart_fill", $"{subject} opacity requires a solid fill from 0 to 100000.");
    }

    internal static bool TryRead(XElement? properties, out SpreadsheetChartSurfaceFill? fill)
    {
        fill = null;
        if (properties is null) return true;
        if (properties.HasAttributes || properties.Elements().Take(2).Count() != 1) return false;
        var paint = properties.Elements().Single();
        if (paint.Name == DrawingNs + "noFill" && !paint.HasAttributes && !paint.HasElements)
        {
            fill = new SpreadsheetChartSurfaceFill { NoFill = true };
            return true;
        }
        if (paint.Name != DrawingNs + "solidFill" || paint.HasAttributes || paint.Elements().Take(2).Count() != 1)
            return false;
        var color = paint.Elements().Single();
        var rgb = (string?)color.Attribute("val");
        if (color.Name != DrawingNs + "srgbClr" || rgb is not { Length: 6 } || !rgb.All(Uri.IsHexDigit) ||
            color.Attributes().Count() != 1)
            return false;
        var alpha = color.Elements(DrawingNs + "alpha").Take(2).ToArray();
        if (color.Elements().Any(child => child.Name != DrawingNs + "alpha") || alpha.Length > 1)
            return false;
        fill = new SpreadsheetChartSurfaceFill { SolidRgb = rgb.ToUpperInvariant() };
        if (alpha.Length == 1)
        {
            var value = (string?)alpha[0].Attribute("val");
            if (alpha[0].HasElements || alpha[0].Attributes().Count() != 1 ||
                !uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var opacity) || opacity > 100_000)
                return false;
            fill.OpacityThousandthPercent = opacity;
        }
        return true;
    }

    internal static XElement? Element(SpreadsheetChartSurfaceFill? fill, string subject)
    {
        if (fill is null) return null;
        Validate(fill, subject);
        XElement paint;
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.NoFill)
            paint = new XElement(DrawingNs + "noFill");
        else
        {
            var color = new XElement(DrawingNs + "srgbClr", new XAttribute("val", PptxColor.Normalize(fill.SolidRgb)));
            if (fill.HasOpacityThousandthPercent)
                color.Add(new XElement(DrawingNs + "alpha", new XAttribute("val", fill.OpacityThousandthPercent)));
            paint = new XElement(DrawingNs + "solidFill", color);
        }
        return new XElement(ChartNs + "spPr", paint);
    }

    internal static void Patch(XElement owner, SpreadsheetChartSurfaceFill? fill, string subject)
    {
        var existing = owner.Element(ChartNs + "spPr");
        if (existing is not null && !TryRead(existing, out _))
            throw new CodecException("unsupported_chart_edit", $"{subject} uses an unmodeled fill graph.");
        var replacement = Element(fill, subject);
        if (replacement is null)
        {
            existing?.Remove();
            return;
        }
        if (existing is not null) existing.ReplaceWith(replacement);
        else
        {
            var extension = owner.Element(ChartNs + "extLst");
            if (extension is null) owner.Add(replacement);
            else extension.AddBeforeSelf(replacement);
        }
    }

    internal static string Semantics(SpreadsheetChartSurfaceFill? fill) => fill?.FillCase switch
    {
        SpreadsheetChartSurfaceFill.FillOneofCase.NoFill => "none",
        SpreadsheetChartSurfaceFill.FillOneofCase.SolidRgb => string.Join(':',
            PptxColor.Normalize(fill.SolidRgb),
            fill.HasOpacityThousandthPercent ? fill.OpacityThousandthPercent.ToString(CultureInfo.InvariantCulture) : "opaque"),
        _ => "default",
    };
}
