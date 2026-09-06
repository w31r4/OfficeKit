using System.Globalization;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns one bounded c:ser/c:dLbls graph: presence-aware series defaults and
// sparse c:dLbl overrides keyed by zero-based c:idx. Bubble-size visibility is
// presence-aware for bubble-series labels. Unsupported label text,
// layout, shape/effect, leader-line, extension, deletion, and source-linked
// number-format graphs leave the containing chart source-owned.
internal static class XlsxChartSeriesDataLabelsCodec
{
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly string[] OrderedFlags = ["showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize"];
    private static readonly HashSet<string> DefaultChildren = new(["dLbl", "numFmt", "txPr", "dLblPos", .. OrderedFlags], StringComparer.Ordinal);
    private static readonly HashSet<string> PointChildren = new(["idx", "numFmt", "txPr", "dLblPos", .. OrderedFlags], StringComparer.Ordinal);
    private static readonly HashSet<string> BooleanValues = new(StringComparer.Ordinal) { "0", "1", "false", "true" };
    private static readonly HashSet<string> PositionValues = new(StringComparer.Ordinal) { "bestFit", "b", "ctr", "inBase", "inEnd", "l", "outEnd", "r", "t" };

    internal static void Validate(SpreadsheetChartSeriesArtifact series, SpreadsheetChartType chartType, string worksheetId, string chartId)
    {
        var labels = series.DataLabels;
        if (labels is null) return;
        if (labels.Defaults is null && labels.Points.Count == 0)
            throw Invalid(worksheetId, chartId, series.Name, "must declare defaults or at least one point override");
        ValidateOverride(labels.Defaults, chartType, worksheetId, chartId, series.Name, "defaults");
        uint? previous = null;
        foreach (var point in labels.Points)
        {
            if (previous is not null && point.Index <= previous.Value)
                throw Invalid(worksheetId, chartId, series.Name, "point indexes must be unique and strictly increasing");
            if (point.Index >= series.Values.Count || series.MissingValueIndexes.Contains(point.Index))
                throw Invalid(worksheetId, chartId, series.Name, $"point index {point.Index} must address an existing non-missing value");
            if (point.Override is null)
                throw Invalid(worksheetId, chartId, series.Name, $"point index {point.Index} must declare at least one override");
            ValidateOverride(point.Override, chartType, worksheetId, chartId, series.Name, $"point {point.Index}");
            previous = point.Index;
        }
    }

    internal static bool TryRead(XElement seriesElement, SpreadsheetChartSeriesArtifact series, SpreadsheetChartType chartType)
    {
        var containers = seriesElement.Elements(ChartNs + "dLbls").ToArray();
        if (containers.Length == 0) return true;
        if (containers.Length != 1) return false;
        var container = containers[0];
        if (container.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || HasUnexpectedText(container)) return false;
        var children = container.Elements().ToArray();
        if (children.Any(child => child.Name.Namespace != ChartNs || !DefaultChildren.Contains(child.Name.LocalName)) ||
            DefaultChildren.Where(name => name != "dLbl").Any(name => children.Count(child => child.Name == ChartNs + name) > 1)) return false;

        if (!TryReadOverride(container, DefaultChildren, chartType, out var defaults, out var hasDefaults)) return false;
        var output = new SpreadsheetChartSeriesDataLabelsArtifact();
        if (hasDefaults) output.Defaults = defaults;
        var indexes = new HashSet<uint>();
        foreach (var pointElement in children.Where(child => child.Name == ChartNs + "dLbl"))
        {
            if (pointElement.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || HasUnexpectedText(pointElement)) return false;
            var pointChildren = pointElement.Elements().ToArray();
            if (pointChildren.Any(child => child.Name.Namespace != ChartNs || !PointChildren.Contains(child.Name.LocalName)) ||
                PointChildren.Any(name => pointChildren.Count(child => child.Name == ChartNs + name) > 1)) return false;
            var nativeIndex = pointElement.Element(ChartNs + "idx");
            if (nativeIndex is null || !TryUInt(nativeIndex, out var index) || index >= series.Values.Count ||
                series.MissingValueIndexes.Contains(index) || !indexes.Add(index)) return false;
            if (!TryReadOverride(pointElement, PointChildren, chartType, out var pointOverride, out var hasPointOverride) || !hasPointOverride) return false;
            output.Points.Add(new SpreadsheetChartPointDataLabelArtifact { Index = index, Override = pointOverride });
        }
        if (output.Defaults is null && output.Points.Count == 0) return false;
        var ordered = output.Points.OrderBy(point => point.Index).Select(point => point.Clone()).ToArray();
        output.Points.Clear();
        output.Points.Add(ordered);
        series.DataLabels = output;
        return true;
    }

    internal static XElement? Element(SpreadsheetChartSeriesDataLabelsArtifact? labels) => labels is null ? null :
        new XElement(ChartNs + "dLbls", labels.Points.OrderBy(point => point.Index).Select(PointElement), Fields(labels.Defaults));

    internal static void Patch(XElement seriesElement, SpreadsheetChartSeriesArtifact target, SpreadsheetChartType chartType, string errorCode, string subject)
    {
        var current = new SpreadsheetChartSeriesArtifact();
        current.Values.Add(target.Values);
        current.MissingValueIndexes.Add(target.MissingValueIndexes);
        if (!TryRead(seriesElement, current, chartType))
            throw new CodecException(errorCode, $"{subject} has series data labels outside the bounded profile.");
        if (Semantics(current.DataLabels) == Semantics(target.DataLabels)) return;

        var existing = seriesElement.Element(ChartNs + "dLbls");
        var replacement = Element(target.DataLabels);
        if (replacement is null) { existing?.Remove(); return; }
        if (existing is not null) { existing.ReplaceWith(replacement); return; }
        var next = seriesElement.Elements().FirstOrDefault(element => element.Name.LocalName is
            "trendline" or "errBars" or "cat" or "xVal" or "val" or "yVal" or "bubbleSize" or "smooth" or "extLst");
        if (next is null) seriesElement.Add(replacement);
        else next.AddBeforeSelf(replacement);
    }

    internal static string Semantics(SpreadsheetChartSeriesDataLabelsArtifact? labels)
    {
        if (labels is null) return "-";
        return $"default:{OverrideSemantics(labels.Defaults)};points:{string.Join('|', labels.Points.OrderBy(point => point.Index).Select(point => $"{point.Index}:{OverrideSemantics(point.Override)}"))}";
    }

    private static bool TryReadOverride(XElement owner, IReadOnlySet<string> allowedChildren, SpreadsheetChartType chartType, out SpreadsheetChartDataLabelOverrideArtifact output, out bool present)
    {
        var parsed = new SpreadsheetChartDataLabelOverrideArtifact();
        output = parsed;
        present = false;
        if (!ReadOptionalBoolean(owner, "showVal", value => parsed.ShowValue = value, ref present) ||
            !ReadOptionalBoolean(owner, "showCatName", value => parsed.ShowCategoryName = value, ref present) ||
            !ReadOptionalBoolean(owner, "showSerName", value => parsed.ShowSeriesName = value, ref present) ||
            !ReadOptionalBoolean(owner, "showPercent", value => parsed.ShowPercent = value, ref present) ||
            !ReadOptionalBoolean(owner, "showBubbleSize", value => parsed.ShowBubbleSize = value, ref present)) return false;
        foreach (var ignored in new[] { "showLegendKey" })
            if (owner.Element(ChartNs + ignored) is { } flag && (!TryBoolean(flag, out var value) || value)) return false;
        if (parsed.HasShowPercent && parsed.ShowPercent && chartType is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut)) return false;
        if (parsed.HasShowBubbleSize && parsed.ShowBubbleSize && chartType != SpreadsheetChartType.Bubble) return false;
        if (owner.Element(ChartNs + "dLblPos") is { } nativePosition)
        {
            if (!TryScalar(nativePosition, PositionValues, out var positionValue) || !TryPosition(positionValue!, out var position)) return false;
            parsed.Position = position;
            present = true;
        }
        if (owner.Element(ChartNs + "numFmt") is { } numberFormat)
        {
            var attributes = numberFormat.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
            var code = (string?)numberFormat.Attribute("formatCode");
            var sourceLinked = (string?)numberFormat.Attribute("sourceLinked");
            if (numberFormat.HasElements || HasUnexpectedText(numberFormat) || attributes.Length != 2 ||
                attributes.Any(attribute => attribute.Name != "formatCode" && attribute.Name != "sourceLinked") ||
                code is null || code.Length > 255 || code.Any(char.IsControl) || sourceLinked is not ("0" or "false")) return false;
            parsed.NumberFormatCode = code;
            present = true;
        }
        if (!XlsxChartTextStyleCodec.TryReadTextProperties(owner, out var textStyle)) return false;
        if (textStyle is not null) { parsed.TextStyle = textStyle; present = true; }
        return owner.Elements().All(child => allowedChildren.Contains(child.Name.LocalName));
    }

    private static IEnumerable<XElement?> Fields(SpreadsheetChartDataLabelOverrideArtifact? value)
    {
        if (value is null) yield break;
        if (value.NumberFormatCode.Length > 0) yield return NumberFormat(value.NumberFormatCode);
        if (value.TextStyle is not null) yield return XlsxChartTextStyleCodec.TextPropertiesElement(value.TextStyle);
        if (value.HasPosition) yield return new XElement(ChartNs + "dLblPos", new XAttribute("val", PositionValue(value.Position)));
        if (value.HasShowValue) yield return BooleanElement("showVal", value.ShowValue);
        if (value.HasShowCategoryName) yield return BooleanElement("showCatName", value.ShowCategoryName);
        if (value.HasShowSeriesName) yield return BooleanElement("showSerName", value.ShowSeriesName);
        if (value.HasShowPercent) yield return BooleanElement("showPercent", value.ShowPercent);
        if (value.HasShowBubbleSize) yield return BooleanElement("showBubbleSize", value.ShowBubbleSize);
    }

    private static XElement PointElement(SpreadsheetChartPointDataLabelArtifact point) => new(ChartNs + "dLbl", new XElement(ChartNs + "idx", new XAttribute("val", point.Index)), Fields(point.Override));
    private static string OverrideSemantics(SpreadsheetChartDataLabelOverrideArtifact? value) => value is null ? "-" : $"value:{Optional(value.HasShowValue, value.ShowValue)};category:{Optional(value.HasShowCategoryName, value.ShowCategoryName)};series:{Optional(value.HasShowSeriesName, value.ShowSeriesName)};percent:{Optional(value.HasShowPercent, value.ShowPercent)};bubbleSize:{Optional(value.HasShowBubbleSize, value.ShowBubbleSize)};position:{(value.HasPosition ? PositionValue(value.Position) : "-")};format:{value.NumberFormatCode};text:{XlsxChartTextStyleCodec.Semantics(value.TextStyle)}";
    private static string Optional(bool present, bool value) => present ? value ? "1" : "0" : "-";

    private static void ValidateOverride(SpreadsheetChartDataLabelOverrideArtifact? value, SpreadsheetChartType chartType, string worksheetId, string chartId, string series, string field)
    {
        if (value is null) return;
        if (!value.HasShowValue && !value.HasShowCategoryName && !value.HasShowSeriesName && !value.HasShowPercent && !value.HasShowBubbleSize && !value.HasPosition && value.TextStyle is null && value.NumberFormatCode.Length == 0)
            throw Invalid(worksheetId, chartId, series, $"{field} must declare at least one bounded property");
        if (value.NumberFormatCode.Length > 255 || value.NumberFormatCode.Any(char.IsControl))
            throw Invalid(worksheetId, chartId, series, $"{field} number format is invalid");
        if (value.HasShowPercent && value.ShowPercent && chartType is not (SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut))
            throw Invalid(worksheetId, chartId, series, $"{field} percentage labels require a pie or doughnut chart");
        if (value.HasShowBubbleSize && value.ShowBubbleSize && chartType != SpreadsheetChartType.Bubble)
            throw Invalid(worksheetId, chartId, series, $"{field} bubble-size labels require a bubble chart");
        if (value.HasPosition && value.Position is not (SpreadsheetChartDataLabelPosition.BestFit or SpreadsheetChartDataLabelPosition.Bottom or SpreadsheetChartDataLabelPosition.Center or SpreadsheetChartDataLabelPosition.InsideBase or SpreadsheetChartDataLabelPosition.InsideEnd or SpreadsheetChartDataLabelPosition.Left or SpreadsheetChartDataLabelPosition.OutsideEnd or SpreadsheetChartDataLabelPosition.Right or SpreadsheetChartDataLabelPosition.Top))
            throw Invalid(worksheetId, chartId, series, $"{field} position is unsupported");
        XlsxChartTextStyleCodec.ValidateStyle(value.TextStyle, worksheetId, chartId, $"series {series} {field}.text_style");
    }

    private static bool ReadOptionalBoolean(XElement owner, string name, Action<bool> assign, ref bool present)
    {
        if (owner.Element(ChartNs + name) is not { } element) return true;
        if (!TryBoolean(element, out var value)) return false;
        assign(value); present = true; return true;
    }

    private static XElement NumberFormat(string code) => new(ChartNs + "numFmt", new XAttribute("formatCode", code), new XAttribute("sourceLinked", "0"));
    private static XElement BooleanElement(string name, bool value) => new(ChartNs + name, new XAttribute("val", value ? "1" : "0"));
    private static bool TryBoolean(XElement element, out bool value) { value = false; if (!TryScalar(element, BooleanValues, out var scalar)) return false; value = scalar is "1" or "true"; return true; }
    private static bool TryUInt(XElement element, out uint value) { value = 0; var scalar = (string?)element.Attribute("val"); return TryScalar(element, scalar is null ? new HashSet<string>() : new HashSet<string> { scalar }, out _) && uint.TryParse(scalar, NumberStyles.None, CultureInfo.InvariantCulture, out value); }
    private static bool TryScalar(XElement element, IReadOnlySet<string> allowed, out string? value) { value = (string?)element.Attribute("val"); var attributes = element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray(); return !element.Elements().Any() && !HasUnexpectedText(element) && attributes.Length == 1 && attributes[0].Name == "val" && value is not null && allowed.Contains(value); }
    private static bool TryPosition(string value, out SpreadsheetChartDataLabelPosition position) { position = value switch { "bestFit" => SpreadsheetChartDataLabelPosition.BestFit, "b" => SpreadsheetChartDataLabelPosition.Bottom, "ctr" => SpreadsheetChartDataLabelPosition.Center, "inBase" => SpreadsheetChartDataLabelPosition.InsideBase, "inEnd" => SpreadsheetChartDataLabelPosition.InsideEnd, "l" => SpreadsheetChartDataLabelPosition.Left, "outEnd" => SpreadsheetChartDataLabelPosition.OutsideEnd, "r" => SpreadsheetChartDataLabelPosition.Right, "t" => SpreadsheetChartDataLabelPosition.Top, _ => SpreadsheetChartDataLabelPosition.Unspecified }; return position != SpreadsheetChartDataLabelPosition.Unspecified; }
    private static string PositionValue(SpreadsheetChartDataLabelPosition position) => position switch { SpreadsheetChartDataLabelPosition.BestFit => "bestFit", SpreadsheetChartDataLabelPosition.Bottom => "b", SpreadsheetChartDataLabelPosition.Center => "ctr", SpreadsheetChartDataLabelPosition.InsideBase => "inBase", SpreadsheetChartDataLabelPosition.InsideEnd => "inEnd", SpreadsheetChartDataLabelPosition.Left => "l", SpreadsheetChartDataLabelPosition.OutsideEnd => "outEnd", SpreadsheetChartDataLabelPosition.Right => "r", SpreadsheetChartDataLabelPosition.Top => "t", _ => throw new InvalidOperationException("Validated series data-label position is unsupported.") };
    private static bool HasUnexpectedText(XElement element) => element.Nodes().Any(node => node switch { XElement => false, XText text => !string.IsNullOrWhiteSpace(text.Value), _ => true });
    private static CodecException Invalid(string worksheetId, string chartId, string series, string message) => new("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chartId} series {series} data labels {message}.");
}
