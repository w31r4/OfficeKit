using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns one plot-level c:dLbls container with optional c:dLblPos plus direct
// c:showVal/c:showCatName booleans plus presence-aware c:showSerName.
// Percentage visibility is presence-aware for circular-chart labels; bubble-
// size visibility is presence-aware for bubble charts. Leader-line visibility
// is presence-aware without claiming ownership of leader-line geometry. The remaining
// unsupported standard show flags are accepted only when false and retained
// during another bounded edit.
internal static class XlsxChartDataLabelsCodec
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly string[] OrderedFlags = ["showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "showLeaderLines"];
    private static readonly HashSet<string> AllowedChildren = new(["numFmt", "spPr", "txPr", "dLblPos", .. OrderedFlags], StringComparer.Ordinal);
    private static readonly HashSet<string> BooleanValues = new(StringComparer.Ordinal) { "0", "1", "false", "true" };
    private static readonly HashSet<string> PositionValues = new(StringComparer.Ordinal) { "bestFit", "b", "ctr", "inBase", "inEnd", "l", "outEnd", "r", "t" };

    internal static void Validate(SpreadsheetChartArtifact chart, string worksheetId)
    {
        if (chart.DataLabels is { NumberFormatCode.Length: > 0 } labels &&
            (labels.NumberFormatCode.Length > 255 || labels.NumberFormatCode.Any(char.IsControl)))
            throw new CodecException("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chart.Id} data-label number format is invalid.");
        if (chart.DataLabels?.HasShowPercent == true && chart.DataLabels.ShowPercent && chart.Type is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut))
            throw new CodecException("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chart.Id} percentage data labels require a pie or doughnut chart.");
        if (chart.DataLabels?.HasShowBubbleSize == true && chart.DataLabels.ShowBubbleSize && chart.Type != SpreadsheetChartType.Bubble)
            throw new CodecException("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chart.Id} bubble-size data labels require a bubble chart.");
        if (chart.DataLabels?.HasPosition == true && chart.DataLabels.Position is not (
            SpreadsheetChartDataLabelPosition.BestFit or SpreadsheetChartDataLabelPosition.Bottom or
            SpreadsheetChartDataLabelPosition.Center or SpreadsheetChartDataLabelPosition.InsideBase or
            SpreadsheetChartDataLabelPosition.InsideEnd or SpreadsheetChartDataLabelPosition.Left or
            SpreadsheetChartDataLabelPosition.OutsideEnd or SpreadsheetChartDataLabelPosition.Right or
            SpreadsheetChartDataLabelPosition.Top))
            throw new CodecException("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chart.Id} data-label position is unsupported.");
        XlsxChartSurfaceFillCodec.Validate(chart.DataLabels?.Fill, $"Worksheet {worksheetId} chart {chart.Id} data-label fill");
        XlsxChartSeriesLineStyleCodec.ValidateLine(chart.DataLabels?.Line, worksheetId, chart.Id, "chart", "data-label line");
    }

    internal static bool TryRead(XElement plot, SpreadsheetChartArtifact chart)
    {
        var containers = plot.Elements(ChartNs + "dLbls").ToArray();
        if (containers.Length == 0) return true;
        if (containers.Length != 1) return false;
        var labels = containers[0];
        if (labels.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || HasUnexpectedText(labels)) return false;
        var children = labels.Elements().ToArray();
        if (children.Any(child => child.Name.Namespace != ChartNs || !AllowedChildren.Contains(child.Name.LocalName)) ||
            AllowedChildren.Any(name => children.Count(child => child.Name == ChartNs + name) > 1)) return false;
        var showValue = children.SingleOrDefault(child => child.Name == ChartNs + "showVal");
        var showCategoryName = children.SingleOrDefault(child => child.Name == ChartNs + "showCatName");
        if (!TryBoolean(showValue, out var value) || !TryBoolean(showCategoryName, out var categoryName)) return false;
        foreach (var name in OrderedFlags.Where(name => name is not "showVal" and not "showCatName" and not "showSerName" and not "showPercent" and not "showBubbleSize" and not "showLeaderLines"))
        {
            var element = children.SingleOrDefault(child => child.Name == ChartNs + name);
            if (element is not null && (!TryBoolean(element, out var enabled) || enabled)) return false;
        }
        var dataLabels = new SpreadsheetChartDataLabelsArtifact { ShowValue = value, ShowCategoryName = categoryName };
        var numberFormat = children.SingleOrDefault(child => child.Name == ChartNs + "numFmt");
        if (numberFormat is not null)
        {
            var attributes = numberFormat.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
            var code = (string?)numberFormat.Attribute("formatCode");
            var sourceLinked = (string?)numberFormat.Attribute("sourceLinked");
            if (numberFormat.HasElements || HasUnexpectedText(numberFormat) || attributes.Length != 2 ||
                attributes.Any(attribute => attribute.Name != "formatCode" && attribute.Name != "sourceLinked") ||
                code is null || code.Length > 255 || code.Any(char.IsControl) || sourceLinked is not ("0" or "false"))
                return false;
            dataLabels.NumberFormatCode = code;
        }
        var showSeriesName = children.SingleOrDefault(child => child.Name == ChartNs + "showSerName");
        if (showSeriesName is not null)
        {
            if (!TryBoolean(showSeriesName, out var seriesName)) return false;
            dataLabels.ShowSeriesName = seriesName;
        }
        var showPercent = children.SingleOrDefault(child => child.Name == ChartNs + "showPercent");
        if (showPercent is not null)
        {
            if (!TryBoolean(showPercent, out var percent)) return false;
            if (percent && chart.Type is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut)) return false;
            dataLabels.ShowPercent = percent;
        }
        var showBubbleSize = children.SingleOrDefault(child => child.Name == ChartNs + "showBubbleSize");
        if (showBubbleSize is not null)
        {
            if (!TryBoolean(showBubbleSize, out var bubbleSize)) return false;
            if (bubbleSize && chart.Type != SpreadsheetChartType.Bubble) return false;
            dataLabels.ShowBubbleSize = bubbleSize;
        }
        var showLeaderLines = children.SingleOrDefault(child => child.Name == ChartNs + "showLeaderLines");
        if (showLeaderLines is not null)
        {
            if (!TryBoolean(showLeaderLines, out var leaderLines)) return false;
            dataLabels.ShowLeaderLines = leaderLines;
        }
        var nativePosition = children.SingleOrDefault(child => child.Name == ChartNs + "dLblPos");
        if (nativePosition is not null)
        {
            if (!TryScalar(nativePosition, PositionValues, out var positionValue) || !TryPosition(positionValue!, out var position)) return false;
            dataLabels.Position = position;
        }
        if (!XlsxChartTextStyleCodec.TryReadTextProperties(labels, out var textStyle) ||
            !TryReadProperties(labels.Element(ChartNs + "spPr"), out var fill, out var line)) return false;
        if (textStyle is not null) dataLabels.TextStyle = textStyle;
        if (fill is not null) dataLabels.Fill = fill;
        if (line is not null) dataLabels.Line = line;
        chart.DataLabels = dataLabels;
        return true;
    }

    internal static XElement? Element(SpreadsheetChartDataLabelsArtifact? labels) => labels is null ? null :
        new XElement(ChartNs + "dLbls",
            labels.NumberFormatCode.Length == 0 ? null : NumberFormatElement(labels.NumberFormatCode),
            DataLabelPropertiesElement(labels.Fill, labels.Line),
            labels.TextStyle is null ? null : XlsxChartTextStyleCodec.TextPropertiesElement(labels.TextStyle),
            PositionElement(labels),
            BooleanElement("showVal", labels.ShowValue),
            BooleanElement("showCatName", labels.ShowCategoryName),
            labels.HasShowSeriesName ? BooleanElement("showSerName", labels.ShowSeriesName) : null,
            labels.HasShowPercent ? BooleanElement("showPercent", labels.ShowPercent) : null,
            labels.HasShowBubbleSize ? BooleanElement("showBubbleSize", labels.ShowBubbleSize) : null,
            labels.HasShowLeaderLines ? BooleanElement("showLeaderLines", labels.ShowLeaderLines) : null);

    internal static void Patch(XElement plot, SpreadsheetChartDataLabelsArtifact? labels)
    {
        var existing = plot.Element(ChartNs + "dLbls");
        if (labels is null) { existing?.Remove(); return; }
        if (existing is null)
        {
            var created = Element(labels)!;
            var lastSeries = plot.Elements(ChartNs + "ser").LastOrDefault();
            if (lastSeries is not null) lastSeries.AddAfterSelf(created);
            else
            {
                var firstAfterSeries = plot.Elements().FirstOrDefault(element => element.Name.LocalName is "dLbls" or "dropLines" or "hiLowLines" or "upDownBars" or "marker" or "smooth" or "gapWidth" or "overlap" or "firstSliceAng" or "holeSize" or "axId" or "extLst");
                if (firstAfterSeries is null) plot.Add(created);
                else firstAfterSeries.AddBeforeSelf(created);
            }
            return;
        }
        if (!TryReadProperties(existing.Element(ChartNs + "spPr"), out var currentFill, out var currentLine))
            throw new CodecException("unsupported_chart_edit", "Chart data labels use an unmodeled shape-properties graph.");
        XlsxChartTextStyleCodec.PatchTextProperties(existing, labels.TextStyle, AllowedChildren);
        PatchNumberFormat(existing, labels.NumberFormatCode);
        var existingPosition = existing.Element(ChartNs + "dLblPos");
        var replacementPosition = PositionElement(labels);
        if (replacementPosition is null) existingPosition?.Remove();
        else if (existingPosition is not null) existingPosition.ReplaceWith(replacementPosition);
        else
        {
            var firstFlag = existing.Elements().FirstOrDefault(element => OrderedFlags.Contains(element.Name.LocalName, StringComparer.Ordinal));
            if (firstFlag is null) existing.AddFirst(replacementPosition);
            else firstFlag.AddBeforeSelf(replacementPosition);
        }
        existing.Element(ChartNs + "showVal")!.SetAttributeValue("val", labels.ShowValue ? "1" : "0");
        existing.Element(ChartNs + "showCatName")!.SetAttributeValue("val", labels.ShowCategoryName ? "1" : "0");
        PatchOptionalBoolean(existing, "showSerName", labels.HasShowSeriesName ? labels.ShowSeriesName : null);
        PatchOptionalBoolean(existing, "showPercent", labels.HasShowPercent ? labels.ShowPercent : null);
        PatchOptionalBoolean(existing, "showBubbleSize", labels.HasShowBubbleSize ? labels.ShowBubbleSize : null);
        PatchOptionalBoolean(existing, "showLeaderLines", labels.HasShowLeaderLines ? labels.ShowLeaderLines : null);
        PatchProperties(existing, currentFill, currentLine, labels.Fill, labels.Line);
    }

    internal static string Semantics(SpreadsheetChartDataLabelsArtifact? labels) => labels is null
        ? "-"
        : $"value:{(labels.ShowValue ? 1 : 0)};category:{(labels.ShowCategoryName ? 1 : 0)};series:{(labels.HasShowSeriesName ? labels.ShowSeriesName ? "1" : "0" : "-")};percent:{(labels.HasShowPercent ? labels.ShowPercent ? "1" : "0" : "-")};bubbleSize:{(labels.HasShowBubbleSize ? labels.ShowBubbleSize ? "1" : "0" : "-")};leaderLines:{(labels.HasShowLeaderLines ? labels.ShowLeaderLines ? "1" : "0" : "-")};position:{(labels.HasPosition ? PositionValue(labels.Position) : "-")};format:{labels.NumberFormatCode};text:{XlsxChartTextStyleCodec.Semantics(labels.TextStyle)};fill:{XlsxChartSurfaceFillCodec.Semantics(labels.Fill)};line:{XlsxChartSeriesLineStyleCodec.Semantics(labels.Line)}";

    private static XElement? DataLabelPropertiesElement(
        SpreadsheetChartSurfaceFill? fill,
        SpreadsheetChartLineStyleArtifact? line)
    {
        if (fill is null && line is null) return null;
        return new XElement(
            ChartNs + "spPr",
            fill is null ? null : XlsxChartSurfaceFillCodec.PaintElement(fill, "Chart data labels"),
            line is null ? null : XlsxChartSeriesLineStyleCodec.Element(line));
    }

    private static bool TryReadProperties(
        XElement? properties,
        out SpreadsheetChartSurfaceFill? fill,
        out SpreadsheetChartLineStyleArtifact? line)
    {
        fill = null;
        line = null;
        if (properties is null) return true;
        if (properties.HasAttributes) return false;
        var children = properties.Elements().ToArray();
        if (children.Length == 0 ||
            children.Any(child => child.Name != DrawingNs + "noFill" &&
                                  child.Name != DrawingNs + "solidFill" &&
                                  child.Name != DrawingNs + "gradFill" &&
                                  child.Name != DrawingNs + "ln") ||
            children.Count(child => child.Name == DrawingNs + "noFill" ||
                                   child.Name == DrawingNs + "solidFill" ||
                                   child.Name == DrawingNs + "gradFill") > 1 ||
            children.Count(child => child.Name == DrawingNs + "ln") > 1)
            return false;
        var paint = children.FirstOrDefault(child => child.Name == DrawingNs + "noFill" ||
                                                     child.Name == DrawingNs + "solidFill" ||
                                                     child.Name == DrawingNs + "gradFill");
        if (!XlsxChartSurfaceFillCodec.TryReadPaint(paint, out fill) ||
            !XlsxChartSeriesLineStyleCodec.TryReadLine(properties, out line)) return false;
        return true;
    }

    private static void PatchProperties(
        XElement labels,
        SpreadsheetChartSurfaceFill? currentFill,
        SpreadsheetChartLineStyleArtifact? currentLine,
        SpreadsheetChartSurfaceFill? targetFill,
        SpreadsheetChartLineStyleArtifact? targetLine)
    {
        if (DataLabelPropertiesSemantics(currentFill, currentLine) == DataLabelPropertiesSemantics(targetFill, targetLine)) return;
        var existing = labels.Element(ChartNs + "spPr");
        var replacement = DataLabelPropertiesElement(targetFill, targetLine);
        if (replacement is null)
        {
            existing?.Remove();
            return;
        }
        if (existing is not null)
        {
            existing.ReplaceWith(replacement);
            return;
        }
        var next = labels.Elements().FirstOrDefault(element => element.Name == ChartNs + "txPr" ||
            element.Name == ChartNs + "dLblPos" || OrderedFlags.Contains(element.Name.LocalName, StringComparer.Ordinal));
        if (next is null) labels.Add(replacement);
        else next.AddBeforeSelf(replacement);
    }

    private static string DataLabelPropertiesSemantics(
        SpreadsheetChartSurfaceFill? fill,
        SpreadsheetChartLineStyleArtifact? line) =>
        string.Join('\0', XlsxChartSurfaceFillCodec.Semantics(fill), XlsxChartSeriesLineStyleCodec.Semantics(line));

    private static void PatchNumberFormat(XElement labels, string code)
    {
        var existing = labels.Element(ChartNs + "numFmt");
        if (code.Length == 0) { existing?.Remove(); return; }
        if (existing is not null && (string?)existing.Attribute("formatCode") == code &&
            (string?)existing.Attribute("sourceLinked") is "0" or "false") return;
        var replacement = NumberFormatElement(code);
        if (existing is not null) { existing.ReplaceWith(replacement); return; }
        var next = labels.Elements().FirstOrDefault(element => element.Name == ChartNs + "txPr" || element.Name == ChartNs + "dLblPos" || OrderedFlags.Contains(element.Name.LocalName, StringComparer.Ordinal));
        if (next is null) labels.AddFirst(replacement);
        else next.AddBeforeSelf(replacement);
    }

    private static XElement NumberFormatElement(string code) => new(ChartNs + "numFmt",
        new XAttribute("formatCode", code),
        new XAttribute("sourceLinked", "0"));

    private static XElement BooleanElement(string name, bool value) =>
        new(ChartNs + name, new XAttribute("val", value ? "1" : "0"));

    private static void PatchOptionalBoolean(XElement labels, string name, bool? value)
    {
        var existing = labels.Element(ChartNs + name);
        if (value is null) { existing?.Remove(); return; }
        if (existing is not null) { existing.SetAttributeValue("val", value.Value ? "1" : "0"); return; }
        var targetIndex = Array.IndexOf(OrderedFlags, name);
        var next = labels.Elements().FirstOrDefault(element => Array.IndexOf(OrderedFlags, element.Name.LocalName) > targetIndex);
        var created = BooleanElement(name, value.Value);
        if (next is null) labels.Add(created);
        else next.AddBeforeSelf(created);
    }

    private static XElement? PositionElement(SpreadsheetChartDataLabelsArtifact labels) => labels.HasPosition
        ? new XElement(ChartNs + "dLblPos", new XAttribute("val", PositionValue(labels.Position)))
        : null;

    private static bool TryBoolean(XElement? element, out bool value)
    {
        value = false;
        if (element is null || !TryScalar(element, BooleanValues, out var scalar)) return false;
        value = scalar is "1" or "true";
        return true;
    }

    private static bool TryScalar(XElement element, IReadOnlySet<string> allowed, out string? value)
    {
        value = (string?)element.Attribute("val");
        var attributes = element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        return !element.Elements().Any() && !HasUnexpectedText(element) && attributes.Length == 1 && attributes[0].Name == "val" && value is not null && allowed.Contains(value);
    }

    private static bool TryPosition(string value, out SpreadsheetChartDataLabelPosition position)
    {
        position = value switch
        {
            "bestFit" => SpreadsheetChartDataLabelPosition.BestFit,
            "b" => SpreadsheetChartDataLabelPosition.Bottom,
            "ctr" => SpreadsheetChartDataLabelPosition.Center,
            "inBase" => SpreadsheetChartDataLabelPosition.InsideBase,
            "inEnd" => SpreadsheetChartDataLabelPosition.InsideEnd,
            "l" => SpreadsheetChartDataLabelPosition.Left,
            "outEnd" => SpreadsheetChartDataLabelPosition.OutsideEnd,
            "r" => SpreadsheetChartDataLabelPosition.Right,
            "t" => SpreadsheetChartDataLabelPosition.Top,
            _ => SpreadsheetChartDataLabelPosition.Unspecified,
        };
        return position != SpreadsheetChartDataLabelPosition.Unspecified;
    }

    private static string PositionValue(SpreadsheetChartDataLabelPosition position) => position switch
    {
        SpreadsheetChartDataLabelPosition.BestFit => "bestFit",
        SpreadsheetChartDataLabelPosition.Bottom => "b",
        SpreadsheetChartDataLabelPosition.Center => "ctr",
        SpreadsheetChartDataLabelPosition.InsideBase => "inBase",
        SpreadsheetChartDataLabelPosition.InsideEnd => "inEnd",
        SpreadsheetChartDataLabelPosition.Left => "l",
        SpreadsheetChartDataLabelPosition.OutsideEnd => "outEnd",
        SpreadsheetChartDataLabelPosition.Right => "r",
        SpreadsheetChartDataLabelPosition.Top => "t",
        _ => throw new InvalidOperationException("Validated data-label position is unsupported."),
    };

    private static bool HasUnexpectedText(XElement element) => element.Nodes().Any(node => node switch
    {
        XElement => false,
        XText text => !string.IsNullOrWhiteSpace(text.Value),
        _ => true,
    });
}
