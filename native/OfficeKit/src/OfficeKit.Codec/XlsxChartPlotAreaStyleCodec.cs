using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns the bounded c:plotArea/c:spPr surface-and-outline pair. Plot-area
// fills and outlines share one native shape-properties container, so they must
// be read and patched together to preserve the sibling style field.
internal static class XlsxChartPlotAreaStyleCodec
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static bool TryRead(
        XElement? properties,
        out SpreadsheetChartSurfaceFill? fill,
        out SpreadsheetChartLineStyleArtifact? line)
    {
        fill = null;
        line = null;
        if (properties is null) return true;
        if (properties.HasAttributes || properties.Nodes().OfType<XText>().Any(node => !string.IsNullOrWhiteSpace(node.Value))) return false;
        var children = properties.Elements().ToArray();
        if (children.Any(child => !IsPaint(child.Name) && child.Name != DrawingNs + "ln") ||
            children.Count(child => IsPaint(child.Name)) > 1 ||
            children.Count(child => child.Name == DrawingNs + "ln") > 1)
            return false;
        var paint = children.SingleOrDefault(child => IsPaint(child.Name));
        if (!XlsxChartSurfaceFillCodec.TryReadPaint(paint, out fill)) return false;
        return XlsxChartSeriesLineStyleCodec.TryReadLine(properties, out line);
    }

    internal static XElement? Element(
        SpreadsheetChartSurfaceFill? fill,
        SpreadsheetChartLineStyleArtifact? line,
        string subject)
    {
        if (fill is null && line is null) return null;
        var output = new XElement(ChartNs + "spPr");
        if (fill is not null) output.Add(XlsxChartSurfaceFillCodec.PaintElement(fill, subject + " fill"));
        if (line is not null) output.Add(XlsxChartSeriesLineStyleCodec.Element(line)!);
        return output;
    }

    internal static void Patch(
        XElement plotArea,
        SpreadsheetChartSurfaceFill? fill,
        SpreadsheetChartLineStyleArtifact? line,
        string subject)
    {
        var existing = plotArea.Element(ChartNs + "spPr");
        SpreadsheetChartSurfaceFill? currentFill = null;
        SpreadsheetChartLineStyleArtifact? currentLine = null;
        if (existing is not null && !TryRead(existing, out currentFill, out currentLine))
            throw new CodecException("unsupported_chart_edit", $"{subject} uses an unmodeled surface or outline graph.");
        if (XlsxChartSurfaceFillCodec.Semantics(currentFill) == XlsxChartSurfaceFillCodec.Semantics(fill) &&
            XlsxChartSeriesLineStyleCodec.Semantics(currentLine) == XlsxChartSeriesLineStyleCodec.Semantics(line))
            return;

        var replacement = Element(fill, line, subject);
        if (replacement is null)
        {
            existing?.Remove();
            return;
        }
        if (existing is not null) existing.ReplaceWith(replacement);
        else
        {
            var extension = plotArea.Element(ChartNs + "extLst");
            if (extension is null) plotArea.Add(replacement);
            else extension.AddBeforeSelf(replacement);
        }
    }

    private static bool IsPaint(XName name) => name == DrawingNs + "noFill" ||
        name == DrawingNs + "solidFill" || name == DrawingNs + "gradFill";
}
