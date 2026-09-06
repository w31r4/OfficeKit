using System.Globalization;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns the bounded c:plotArea/c:dTable profile shared by worksheet and
// presentation ChartParts. Text properties, extensions, and other native
// graphs remain source-owned so a data-table edit never flattens them.
internal static class XlsxChartDataTableCodec
{
    private static readonly string[] BooleanValues = ["0", "1", "false", "true"];
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    internal static void Validate(
        SpreadsheetChartDataTableArtifact? dataTable,
        string worksheetId,
        string chartId)
    {
        if (dataTable is null) return;
        XlsxChartSurfaceFillCodec.Validate(dataTable.Fill, $"Worksheet {worksheetId} chart {chartId} data-table fill");
        XlsxChartSeriesLineStyleCodec.ValidateLine(
            dataTable.Line,
            worksheetId,
            chartId,
            chartId,
            "data-table line");
    }

    internal static bool TryRead(XElement plotArea, SpreadsheetChartArtifact chart)
    {
        var matches = plotArea.Elements(ChartNs + "dTable").Take(2).ToArray();
        if (matches.Length > 1) return false;
        if (!TryReadElement(matches.SingleOrDefault(), out var dataTable)) return false;
        chart.DataTable = dataTable;
        return true;
    }

    internal static XElement? Element(
        SpreadsheetChartDataTableArtifact? dataTable,
        string subject)
    {
        if (dataTable is null) return null;
        var output = new XElement(ChartNs + "dTable");
        if (dataTable.HasShowHorizontalBorder)
            output.Add(BooleanElement("showHorzBorder", dataTable.ShowHorizontalBorder));
        if (dataTable.HasShowVerticalBorder)
            output.Add(BooleanElement("showVertBorder", dataTable.ShowVerticalBorder));
        if (dataTable.HasShowOutlineBorder)
            output.Add(BooleanElement("showOutline", dataTable.ShowOutlineBorder));
        if (dataTable.HasShowLegendKey)
            output.Add(BooleanElement("showKeys", dataTable.ShowLegendKey));
        if (XlsxChartPlotAreaStyleCodec.Element(dataTable.Fill, dataTable.Line, subject) is { } style)
            output.Add(style);
        return output;
    }

    internal static void Patch(
        XElement plotArea,
        SpreadsheetChartDataTableArtifact? dataTable,
        string errorCode,
        string subject)
    {
        var existing = plotArea.Elements(ChartNs + "dTable").Take(2).ToArray();
        if (existing.Length > 1 || existing.Length == 1 && !TryReadElement(existing[0], out _))
            throw new CodecException(errorCode, $"{subject} has an invalid c:dTable.");
        if (existing.Length == 0 && dataTable is null) return;

        var replacement = Element(dataTable, $"{subject} data table");
        if (existing.Length == 1)
        {
            if (replacement is null) existing[0].Remove();
            else existing[0].ReplaceWith(replacement);
            return;
        }
        if (replacement is null) return;
        var following = plotArea.Elements().FirstOrDefault(element =>
            element.Name == ChartNs + "spPr" || element.Name == ChartNs + "extLst");
        if (following is null) plotArea.Add(replacement);
        else following.AddBeforeSelf(replacement);
    }

    internal static string Semantics(SpreadsheetChartDataTableArtifact? dataTable)
    {
        if (dataTable is null) return "absent";
        return string.Join(
            '\u001e',
            OptionalBooleanSemantics(dataTable.HasShowHorizontalBorder, dataTable.ShowHorizontalBorder),
            OptionalBooleanSemantics(dataTable.HasShowVerticalBorder, dataTable.ShowVerticalBorder),
            OptionalBooleanSemantics(dataTable.HasShowOutlineBorder, dataTable.ShowOutlineBorder),
            OptionalBooleanSemantics(dataTable.HasShowLegendKey, dataTable.ShowLegendKey),
            XlsxChartSurfaceFillCodec.Semantics(dataTable.Fill),
            XlsxChartSeriesLineStyleCodec.Semantics(dataTable.Line));
    }

    private static bool TryReadElement(
        XElement? source,
        out SpreadsheetChartDataTableArtifact? dataTable)
    {
        dataTable = null;
        if (source is null) return true;
        if (source.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
            source.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
            return false;
        var children = source.Elements().ToArray();
        if (children.Any(child => child.Name != ChartNs + "showHorzBorder" &&
                                  child.Name != ChartNs + "showVertBorder" &&
                                  child.Name != ChartNs + "showOutline" &&
                                  child.Name != ChartNs + "showKeys" &&
                                  child.Name != ChartNs + "spPr"))
            return false;
        if (children.Count(child => child.Name == ChartNs + "spPr") > 1)
            return false;
        if (!TryOptionalBoolean(source, "showHorzBorder", out var hasHorizontal, out var horizontal) ||
            !TryOptionalBoolean(source, "showVertBorder", out var hasVertical, out var vertical) ||
            !TryOptionalBoolean(source, "showOutline", out var hasOutline, out var outline) ||
            !TryOptionalBoolean(source, "showKeys", out var hasKeys, out var keys) ||
            !XlsxChartPlotAreaStyleCodec.TryRead(source.Element(ChartNs + "spPr"), out var fill, out var line))
            return false;

        dataTable = new SpreadsheetChartDataTableArtifact();
        if (hasHorizontal) dataTable.ShowHorizontalBorder = horizontal;
        if (hasVertical) dataTable.ShowVerticalBorder = vertical;
        if (hasOutline) dataTable.ShowOutlineBorder = outline;
        if (hasKeys) dataTable.ShowLegendKey = keys;
        if (fill is not null) dataTable.Fill = fill;
        if (line is not null) dataTable.Line = line;
        return true;
    }

    private static bool TryOptionalBoolean(
        XElement owner,
        string name,
        out bool present,
        out bool value)
    {
        present = false;
        value = false;
        var matches = owner.Elements(ChartNs + name).Take(2).ToArray();
        if (matches.Length == 0) return true;
        if (matches.Length != 1) return false;
        var element = matches[0];
        if (element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "val") ||
            element.Elements().Any() ||
            element.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
            return false;
        var text = (string?)element.Attribute("val");
        if (text is null || !BooleanValues.Contains(text, StringComparer.Ordinal)) return false;
        present = true;
        value = text is "1" or "true";
        return true;
    }

    private static XElement BooleanElement(string name, bool value) =>
        new(ChartNs + name, new XAttribute("val", value ? "1" : "0"));

    private static string OptionalBooleanSemantics(bool present, bool value) =>
        present ? value ? "1" : "0" : "absent";
}
