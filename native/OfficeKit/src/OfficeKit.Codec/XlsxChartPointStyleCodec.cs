using System.Globalization;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns one bounded sparse c:dPt profile. It deliberately models direct point
// paint/outline plus circular-slice explosion only; marker, picture, 3D,
// effects, extensions and unknown point graphs make the chart read-only.
internal static class XlsxChartPointStyleCodec
{
    private const int MaxPointStyles = 256;
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

    internal static void Validate(
        SpreadsheetChartSeriesArtifact series,
        SpreadsheetChartType chartType,
        string worksheetId,
        string chartId)
    {
        if (series.PointStyles.Count == 0) return;
        if (chartType is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut))
            throw Invalid(worksheetId, chartId, series.Name, "point styles require a bar, column, pie, or doughnut chart");
        if (series.PointStyles.Count > MaxPointStyles)
            throw Invalid(worksheetId, chartId, series.Name, $"point styles exceed the {MaxPointStyles}-entry budget");

        uint? previous = null;
        var missing = series.MissingValueIndexes.ToHashSet();
        foreach (var point in series.PointStyles)
        {
            if (point.Index >= (uint)series.Values.Count || previous is not null && point.Index <= previous.Value || missing.Contains(point.Index))
                throw Invalid(worksheetId, chartId, series.Name, "point-style indexes must be strictly increasing, unique, in range, and non-missing");
            if (point.Fill is null && point.Line is null && !point.HasExplosion)
                throw Invalid(worksheetId, chartId, series.Name, $"point {point.Index} has no visual override");
            XlsxChartSurfaceFillCodec.Validate(point.Fill, $"Worksheet {worksheetId} chart {chartId} series {series.Name} point {point.Index} fill");
            XlsxChartSeriesLineStyleCodec.ValidateLine(point.Line, worksheetId, chartId, series.Name, $"point {point.Index} line");
            if (point.HasExplosion && (chartType is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut) || point.Explosion > 400))
                throw Invalid(worksheetId, chartId, series.Name, $"point {point.Index} explosion requires pie/doughnut and must be 0 through 400");
            previous = point.Index;
        }
    }

    internal static bool TryRead(
        XElement nativeSeries,
        SpreadsheetChartSeriesArtifact series,
        SpreadsheetChartType chartType)
    {
        var nativePoints = nativeSeries.Elements(ChartNs + "dPt").ToArray();
        if (nativePoints.Length == 0) return true;
        if (nativePoints.Length > MaxPointStyles ||
            chartType is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut)) return false;

        var projected = new List<SpreadsheetChartPointStyleArtifact>(nativePoints.Length);
        uint? previous = null;
        var missing = series.MissingValueIndexes.ToHashSet();
        foreach (var native in nativePoints)
        {
            if (native.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)) return false;
            var children = native.Elements().ToArray();
            if (children.Any(child => child.Name != ChartNs + "idx" && child.Name != ChartNs + "explosion" && child.Name != ChartNs + "spPr") ||
                children.Count(child => child.Name == ChartNs + "idx") != 1 ||
                children.Count(child => child.Name == ChartNs + "explosion") > 1 ||
                children.Count(child => child.Name == ChartNs + "spPr") > 1) return false;

            var indexElement = children.Single(child => child.Name == ChartNs + "idx");
            if (!TryUnsignedScalar(indexElement, out var index) || index >= (uint)series.Values.Count ||
                previous is not null && index <= previous.Value || missing.Contains(index)) return false;

            var output = new SpreadsheetChartPointStyleArtifact { Index = index };
            var explosion = children.SingleOrDefault(child => child.Name == ChartNs + "explosion");
            if (explosion is not null)
            {
                if (chartType is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut) ||
                    !TryUnsignedScalar(explosion, out var value) || value > 400) return false;
                output.Explosion = value;
            }

            var shapeProperties = children.SingleOrDefault(child => child.Name == ChartNs + "spPr");
            if (shapeProperties is not null)
            {
                if (shapeProperties.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
                    shapeProperties.Elements().Any(child => !FillNames.Contains(child.Name) && child.Name != DrawingNs + "ln") ||
                    shapeProperties.Elements().Count(child => FillNames.Contains(child.Name)) > 1 ||
                    shapeProperties.Elements(DrawingNs + "ln").Take(2).Count() > 1 ||
                    !XlsxChartSurfaceFillCodec.TryRead(shapeProperties, out var fill) ||
                    !XlsxChartSeriesLineStyleCodec.TryReadLine(shapeProperties, out var line)) return false;
                if (fill is not null) output.Fill = fill;
                if (line is not null) output.Line = line;
            }
            if (output.Fill is null && output.Line is null && !output.HasExplosion) return false;
            projected.Add(output);
            previous = index;
        }
        series.PointStyles.Add(projected);
        return true;
    }

    internal static IEnumerable<XElement> Elements(SpreadsheetChartSeriesArtifact series)
    {
        foreach (var point in series.PointStyles)
        {
            var fill = point.Fill is null ? null : XlsxChartSurfaceFillCodec.PaintElement(point.Fill, $"Series {series.Name} point {point.Index} fill");
            var line = XlsxChartSeriesLineStyleCodec.Element(point.Line);
            yield return new XElement(
                ChartNs + "dPt",
                new XElement(ChartNs + "idx", new XAttribute("val", point.Index)),
                point.HasExplosion ? new XElement(ChartNs + "explosion", new XAttribute("val", point.Explosion)) : null,
                fill is null && line is null ? null : new XElement(ChartNs + "spPr", fill, line));
        }
    }

    internal static void Patch(
        XElement nativeSeries,
        SpreadsheetChartSeriesArtifact target,
        SpreadsheetChartType chartType,
        string errorCode,
        string subject)
    {
        var current = new SpreadsheetChartSeriesArtifact();
        current.Values.Add(target.Values);
        current.MissingValueIndexes.Add(target.MissingValueIndexes);
        if (!TryRead(nativeSeries, current, chartType))
            throw new CodecException(errorCode, $"{subject} uses an unsupported native point-style graph.");
        if (Semantics(current.PointStyles) == Semantics(target.PointStyles)) return;

        foreach (var existing in nativeSeries.Elements(ChartNs + "dPt").ToArray()) existing.Remove();
        var replacements = Elements(target).ToArray();
        if (replacements.Length == 0) return;
        var before = nativeSeries.Elements().FirstOrDefault(element =>
            element.Name == ChartNs + "dLbls" || element.Name == ChartNs + "trendline" ||
            element.Name == ChartNs + "errBars" || element.Name == ChartNs + "cat" ||
            element.Name == ChartNs + "val" || element.Name == ChartNs + "xVal" ||
            element.Name == ChartNs + "yVal" || element.Name == ChartNs + "bubbleSize" ||
            element.Name == ChartNs + "smooth" || element.Name == ChartNs + "extLst");
        if (before is null) nativeSeries.Add(replacements);
        else foreach (var replacement in replacements) before.AddBeforeSelf(replacement);
    }

    internal static string Semantics(IEnumerable<SpreadsheetChartPointStyleArtifact> points) =>
        string.Join('\u001e', points.Select(point => string.Join(':',
            point.Index,
            XlsxChartSurfaceFillCodec.Semantics(point.Fill),
            XlsxChartSeriesLineStyleCodec.Semantics(point.Line),
            point.HasExplosion ? point.Explosion.ToString(CultureInfo.InvariantCulture) : "absent")));

    private static bool TryUnsignedScalar(XElement element, out uint value)
    {
        value = 0;
        if (element.HasElements || element.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value))) return false;
        var attributes = element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        return attributes.Length == 1 && attributes[0].Name == "val" &&
            uint.TryParse(attributes[0].Value, NumberStyles.None, CultureInfo.InvariantCulture, out value);
    }

    private static CodecException Invalid(string worksheetId, string chartId, string seriesName, string message) =>
        new("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chartId} series {seriesName} {message}.");
}
