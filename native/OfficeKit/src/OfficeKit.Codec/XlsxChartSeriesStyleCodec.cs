using System.Globalization;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns the bounded series-level DrawingML style projection. The public model
// retains the legacy six-digit RGB scalar and adds the shared no/solid/gradient
// paint profile for PPJ. Other series shape properties remain source-owned;
// unrecognized fill kinds make the containing chart read-only.
internal static class XlsxChartSeriesStyleCodec
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly HashSet<XName> FillNames =
    [
        DrawingNs + "noFill",
        DrawingNs + "solidFill",
        DrawingNs + "gradFill",
        DrawingNs + "blipFill",
        DrawingNs + "pattFill",
        DrawingNs + "grpFill",
    ];

    internal static void Validate(SpreadsheetChartSeriesArtifact series, string worksheetId, string chartId)
    {
        ValidateFill(series.Fill, worksheetId, chartId, series.Name, "fill");
        if (series.Fill is not null && series.SeriesFill is not null)
            throw new CodecException("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chartId} series {series.Name} cannot combine fill and series_fill.");
        XlsxChartSurfaceFillCodec.Validate(series.SeriesFill, $"Worksheet {worksheetId} chart {chartId} series {series.Name} fill");
    }

    internal static void ValidateFill(SpreadsheetColor? fill, string worksheetId, string chartId, string seriesName, string subject)
    {
        if (fill is null) return;
        if (fill.SourceCase != SpreadsheetColor.SourceOneofCase.Rgb ||
            fill.HasTint ||
            fill.Rgb.Length != 6 ||
            !fill.Rgb.All(Uri.IsHexDigit))
        {
            throw new CodecException(
                "invalid_spreadsheet_chart",
                $"Worksheet {worksheetId} chart {chartId} series {seriesName} {subject} must be an untinted six-digit RGB solid color.");
        }
    }

    internal static bool TryRead(XElement nativeSeries, SpreadsheetChartSeriesArtifact series)
    {
        var shapeProperties = nativeSeries.Element(ChartNs + "spPr");
        if (shapeProperties is null) return true;
        var paints = shapeProperties.Elements().Where(item => FillNames.Contains(item.Name)).ToArray();
        if (paints.Length == 0) return true;
        if (paints.Length != 1 || !XlsxChartSurfaceFillCodec.TryReadPaint(paints[0], out var fill) || fill is null) return false;
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.SolidRgb && !fill.HasOpacityThousandthPercent)
            series.Fill = new SpreadsheetColor { Rgb = fill.SolidRgb };
        else
            series.SeriesFill = fill;
        return true;
    }

    internal static bool TryReadSolidFill(
        XElement? shapeProperties,
        out SpreadsheetColor? fill,
        out uint? opacityThousandthPercent)
    {
        fill = null;
        opacityThousandthPercent = null;
        if (shapeProperties is null) return true;
        var fills = shapeProperties.Elements().Where(item => FillNames.Contains(item.Name)).ToArray();
        if (fills.Length == 0) return true;
        if (fills.Length != 1 || fills[0].Name != DrawingNs + "solidFill") return false;

        var solidFill = fills[0];
        if (solidFill.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)) return false;
        var colors = solidFill.Elements().ToArray();
        if (colors.Length != 1 || colors[0].Name != DrawingNs + "srgbClr") return false;
        var color = colors[0];
        if (color.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "val")) return false;
        var value = (string?)color.Attribute("val");
        if (value is null || value.Length != 6 || !value.All(Uri.IsHexDigit)) return false;
        var transforms = color.Elements().ToArray();
        if (transforms.Length > 1 || transforms.Any(transform => transform.Name != DrawingNs + "alpha")) return false;
        if (transforms.SingleOrDefault() is { } alpha)
        {
            if (!IsScalarAlpha(alpha, out var opacity)) return false;
            opacityThousandthPercent = opacity;
        }
        fill = new SpreadsheetColor { Rgb = value.ToUpperInvariant() };
        return true;
    }

    internal static XElement? PropertiesElement(SpreadsheetChartSeriesArtifact series, bool markerOnly = false)
    {
        var semanticFill = EffectiveFill(series);
        var fill = semanticFill is null ? null : XlsxChartSurfaceFillCodec.PaintElement(semanticFill, $"Series {series.Name} fill");
        var line = XlsxChartSeriesLineStyleCodec.Element(series.Line, markerOnly);
        return fill is null && line is null ? null : new XElement(ChartNs + "spPr", fill, line);
    }

    internal static void Patch(XElement nativeSeries, SpreadsheetChartSeriesArtifact target)
    {
        var shapeProperties = nativeSeries.Element(ChartNs + "spPr");
        var existingPaints = shapeProperties?.Elements().Where(item => FillNames.Contains(item.Name)).ToArray() ?? [];
        if (existingPaints.Length > 1 || existingPaints.Length == 1 &&
            !XlsxChartSurfaceFillCodec.TryReadPaint(existingPaints[0], out _))
            throw new CodecException("unsupported_chart_edit", $"Series {target.Name} uses an unmodeled fill graph.");
        var existingPaint = existingPaints.SingleOrDefault();
        var targetFill = EffectiveFill(target);
        var semanticallyEqual = existingPaint is null
            ? targetFill is null
            : XlsxChartSurfaceFillCodec.TryReadPaint(existingPaint, out var existingFill) &&
              XlsxChartSurfaceFillCodec.Semantics(existingFill) == XlsxChartSurfaceFillCodec.Semantics(targetFill);
        if (!semanticallyEqual && targetFill is null)
        {
            existingPaint?.Remove();
            if (shapeProperties is not null && !shapeProperties.Elements().Any() && !shapeProperties.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)) shapeProperties.Remove();
            return;
        }

        if (!semanticallyEqual && shapeProperties is null)
        {
            shapeProperties = new XElement(ChartNs + "spPr");
            var before = nativeSeries.Elements().FirstOrDefault(item => item.Name != ChartNs + "idx" && item.Name != ChartNs + "order" && item.Name != ChartNs + "tx");
            if (before is null) nativeSeries.Add(shapeProperties);
            else before.AddBeforeSelf(shapeProperties);
        }
        if (!semanticallyEqual)
        {
            var replacement = XlsxChartSurfaceFillCodec.PaintElement(targetFill!, $"Series {target.Name} fill");
            if (existingPaint is not null) existingPaint.ReplaceWith(replacement);
            else
            {
                var before = shapeProperties!.Elements().FirstOrDefault(item => IsShapePropertyTail(item.Name));
                if (before is null) shapeProperties.Add(replacement);
                else before.AddBeforeSelf(replacement);
            }
        }
    }

    internal static string Semantics(SpreadsheetChartSeriesArtifact series) =>
        XlsxChartSurfaceFillCodec.Semantics(EffectiveFill(series));

    internal static XElement SolidFillElement(string rgb, uint? opacityThousandthPercent = null)
    {
        var color = new XElement(DrawingNs + "srgbClr", new XAttribute("val", rgb.ToUpperInvariant()));
        if (opacityThousandthPercent is { } opacity)
            color.Add(new XElement(DrawingNs + "alpha", new XAttribute("val", opacity)));
        return new XElement(DrawingNs + "solidFill", color);
    }

    private static bool IsScalarAlpha(XElement element, out uint opacity)
    {
        opacity = 0;
        return element.Attributes().All(attribute => attribute.IsNamespaceDeclaration || attribute.Name == "val") &&
            element.Nodes().All(node => node is XText text && string.IsNullOrWhiteSpace(text.Value)) &&
            uint.TryParse((string?)element.Attribute("val"), NumberStyles.None, CultureInfo.InvariantCulture, out opacity) &&
            opacity <= 100_000;
    }

    private static SpreadsheetChartSurfaceFill? EffectiveFill(SpreadsheetChartSeriesArtifact series)
    {
        if (series.SeriesFill is not null) return series.SeriesFill;
        return series.Fill is null ? null : new SpreadsheetChartSurfaceFill { SolidRgb = series.Fill.Rgb };
    }

    private static bool IsShapePropertyTail(XName name) =>
        name == DrawingNs + "ln" || name == DrawingNs + "effectLst" || name == DrawingNs + "effectDag" ||
        name == DrawingNs + "scene3d" || name == DrawingNs + "sp3d" || name == DrawingNs + "extLst";
}
