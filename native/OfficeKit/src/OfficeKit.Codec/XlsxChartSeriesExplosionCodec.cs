using System.Globalization;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns one bounded series-level c:explosion scalar. It applies only to
// ordinary pie/doughnut series; point-level c:dPt explosion remains owned by
// XlsxChartPointStyleCodec. Unknown series children and non-circular uses fail
// closed instead of being flattened into the PPJ series model.
internal static class XlsxChartSeriesExplosionCodec
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    internal static void Validate(
        SpreadsheetChartSeriesArtifact series,
        SpreadsheetChartType chartType,
        string worksheetId,
        string chartId)
    {
        if (!series.HasExplosion) return;
        if (chartType is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut) || series.Explosion > 400)
            throw new CodecException(
                "invalid_spreadsheet_chart",
                $"Worksheet {worksheetId} chart {chartId} series {series.Name} explosion requires pie/doughnut and must be 0 through 400.");
    }

    internal static bool TryRead(
        XElement nativeSeries,
        SpreadsheetChartSeriesArtifact series,
        SpreadsheetChartType chartType)
    {
        var elements = nativeSeries.Elements(ChartNs + "explosion").Take(2).ToArray();
        if (elements.Length == 0) return true;
        if (elements.Length != 1 || chartType is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut) ||
            !TryUnsignedScalar(elements[0], out var value) || value > 400)
            return false;
        series.Explosion = value;
        return true;
    }

    internal static XElement? Element(SpreadsheetChartSeriesArtifact series) =>
        series.HasExplosion
            ? new XElement(ChartNs + "explosion", new XAttribute("val", series.Explosion.ToString(CultureInfo.InvariantCulture)))
            : null;

    internal static void Patch(
        XElement nativeSeries,
        SpreadsheetChartSeriesArtifact target,
        SpreadsheetChartType chartType,
        string errorCode,
        string subject)
    {
        var elements = nativeSeries.Elements(ChartNs + "explosion").Take(2).ToArray();
        if (elements.Length > 1 || elements.Length == 1 &&
            (chartType is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut) ||
             !TryUnsignedScalar(elements[0], out var current) || current > 400))
            throw new CodecException(errorCode, $"{subject} has an invalid series-level c:explosion.");

        if (!target.HasExplosion)
        {
            elements.SingleOrDefault()?.Remove();
            return;
        }
        if (chartType is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut) || target.Explosion > 400)
            throw new CodecException(errorCode, $"{subject} has an invalid pie/doughnut series explosion.");

        var replacement = Element(target)!;
        if (elements.Length == 1)
        {
            elements[0].ReplaceWith(replacement);
            return;
        }

        var before = nativeSeries.Elements().FirstOrDefault(element =>
            element.Name == ChartNs + "invertIfNegative" ||
            element.Name == ChartNs + "marker" ||
            element.Name == ChartNs + "trendline" ||
            element.Name == ChartNs + "errBars" ||
            element.Name == ChartNs + "dLbls" ||
            element.Name == ChartNs + "cat" ||
            element.Name == ChartNs + "val" ||
            element.Name == ChartNs + "xVal" ||
            element.Name == ChartNs + "yVal" ||
            element.Name == ChartNs + "bubbleSize" ||
            element.Name == ChartNs + "smooth" ||
            element.Name == ChartNs + "dPt" ||
            element.Name == ChartNs + "extLst");
        if (before is null) nativeSeries.Add(replacement);
        else before.AddBeforeSelf(replacement);
    }

    internal static string Semantics(SpreadsheetChartSeriesArtifact series) =>
        series.HasExplosion ? series.Explosion.ToString(CultureInfo.InvariantCulture) : "absent";

    private static bool TryUnsignedScalar(XElement element, out uint value)
    {
        value = 0;
        if (element.HasElements || element.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value))) return false;
        var attributes = element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        return attributes.Length == 1 && attributes[0].Name == "val" &&
            uint.TryParse(attributes[0].Value, NumberStyles.None, CultureInfo.InvariantCulture, out value);
    }
}
