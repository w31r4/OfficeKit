using System.Globalization;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns the bounded primary category/value or numeric-X value/value-axis projection for worksheet charts.
// Axis identity and all unmodeled formatting remain in the ChartPart; this
// module reads and patches titles, number formats, category label interval,
// linear value-axis bounds/unit, orientation, and bounded line/text styles.
internal static class XlsxChartAxisCodec
{
    private const uint MaxTickLabelInterval = 1_048_576;
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static void Validate(SpreadsheetChartArtifact chart, string worksheetId)
    {
        if (chart.Type is SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut)
        {
            if (chart.XAxis is not null || chart.YAxis is not null)
                throw Invalid(worksheetId, chart.Id, "pie and doughnut charts cannot carry category/value axes in the bounded profile.");
            return;
        }
        if ((chart.XAxis is null) != (chart.YAxis is null))
            throw Invalid(worksheetId, chart.Id, "must carry both x_axis and y_axis or neither for backward-compatible default authoring.");
        if (chart.XAxis is null) return;
        ValidateAxis(chart.XAxis, !UsesNumericXAxis(chart.Type), "x", worksheetId, chart.Id);
        ValidateAxis(chart.YAxis!, false, "y", worksheetId, chart.Id);
    }

    internal static bool TryRead(XElement plotArea, XElement plot, SpreadsheetChartArtifact chart, out bool editable)
        => TryReadAtPositions(plotArea, plot, chart, "b", "l", out editable);

    // Presentation combo charts own a second category/value pair for the one
    // supported secondary line plot. Keeping this in the axis codec makes the
    // native IDs/positions an implementation detail of one deep helper rather
    // than a second hand-rolled axis reader in the combo codec.
    internal static bool TryReadPresentationSecondary(XElement plotArea, XElement plot, SpreadsheetChartArtifact chart, out bool editable)
        => TryReadAtPositions(plotArea, plot, chart, "t", "r", out editable);

    private static bool TryReadAtPositions(XElement plotArea, XElement plot, SpreadsheetChartArtifact chart, string horizontalPosition, string verticalPosition, out bool editable)
    {
        editable = false;
        if (chart.Type is SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut)
        {
            editable = !plotArea.Elements().Any(IsAxis);
            return editable;
        }
        var numericX = UsesNumericXAxis(chart.Type);
        if (!TryLocate(plotArea, plot, numericX, horizontalPosition, verticalPosition, out var horizontalAxis, out var verticalAxis)) return false;
        if (!TryReadAxis(horizontalAxis, !numericX, horizontalPosition, out var xAxis, out var xEditable) ||
            !TryReadAxis(verticalAxis, false, verticalPosition, out var yAxis, out var yEditable)) return false;
        chart.XAxis = xAxis;
        chart.YAxis = yAxis;
        editable = xEditable && yEditable;
        return true;
    }

    internal static void AppendAuthored(XElement plotArea, SpreadsheetChartArtifact chart)
    {
        if (chart.Type is SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut) return;
        var xAxis = chart.XAxis ?? new SpreadsheetChartAxisArtifact();
        var yAxis = chart.YAxis ?? new SpreadsheetChartAxisArtifact();
        if (UsesNumericXAxis(chart.Type))
            plotArea.Add(BuildValueAxis(xAxis, "1", "2", "b"), BuildValueAxis(yAxis, "2", "1", "l"));
        else
            plotArea.Add(BuildCategoryAxis(xAxis), BuildValueAxis(yAxis, "2", "1", "l"));
    }

    internal static void AppendAuthoredPresentationSecondary(XElement plotArea, SpreadsheetChartArtifact chart)
    {
        var xAxis = chart.XAxis ?? new SpreadsheetChartAxisArtifact();
        var yAxis = chart.YAxis ?? new SpreadsheetChartAxisArtifact();
        // A right-hand value axis crosses its paired category axis at the
        // category maximum. Besides being the interoperable DrawingML form,
        // this makes the intended side explicit to hosts instead of relying on
        // their default crossing inference.
        plotArea.Add(BuildCategoryAxis(xAxis, "3", "4", "t"), BuildValueAxis(yAxis, "4", "3", "r", crosses: "max"));
    }

    internal static void Patch(XElement plotArea, XElement plot, SpreadsheetChartArtifact target)
    {
        if (target.Type is SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut) return;
        var numericX = UsesNumericXAxis(target.Type);
        if (target.XAxis is null || target.YAxis is null || !TryLocate(plotArea, plot, numericX, "b", "l", out var horizontalAxis, out var verticalAxis))
            throw new CodecException("unsupported_spreadsheet_chart_edit", $"Worksheet chart {target.Id} cannot change its primary-axis topology.");
        PatchAxis(horizontalAxis, target.XAxis, !numericX);
        PatchAxis(verticalAxis, target.YAxis, false);
    }

    internal static void PatchPresentationSecondary(XElement plotArea, XElement plot, SpreadsheetChartArtifact target)
    {
        if (target.XAxis is null || target.YAxis is null || !TryLocate(plotArea, plot, false, "t", "r", out var horizontalAxis, out var verticalAxis))
            throw new CodecException("unsupported_presentation_edit", "Presentation combo chart cannot change its secondary-axis topology.");
        PatchAxis(horizontalAxis, target.XAxis, true);
        PatchAxis(verticalAxis, target.YAxis, false);
    }

    internal static string Semantics(SpreadsheetChartArtifact chart) =>
        string.Join('\u001d', AxisSemantics(chart.XAxis), AxisSemantics(chart.YAxis));

    private static void ValidateAxis(SpreadsheetChartAxisArtifact axis, bool category, string axisName, string worksheetId, string chartId)
    {
        if (axis.Title.Length > 32_767 || HasControls(axis.Title)) throw Invalid(worksheetId, chartId, $"{axisName}-axis title is invalid.");
        if (axis.NumberFormatCode.Length > 255 || HasControls(axis.NumberFormatCode)) throw Invalid(worksheetId, chartId, $"{axisName}-axis number format is invalid.");
        if (axis.AxisLine is not null && axis.HasAxisLineVisible && !axis.AxisLineVisible)
            throw Invalid(worksheetId, chartId, $"{axisName}-axis cannot combine a hidden axis line with a line style.");
        if (axis.MajorGridlineStyle is not null && (!axis.HasShowMajorGridlines || !axis.ShowMajorGridlines))
            throw Invalid(worksheetId, chartId, $"{axisName}-axis grid-line style requires visible major gridlines.");
        if (axis.MajorGridlineStyle is not null && axis.HasMajorGridlineVisible && !axis.MajorGridlineVisible)
            throw Invalid(worksheetId, chartId, $"{axisName}-axis cannot combine a hidden major gridline with a line style.");
        XlsxChartSeriesLineStyleCodec.ValidateLine(axis.AxisLine, worksheetId, chartId, axisName, "axis line");
        XlsxChartSeriesLineStyleCodec.ValidateLine(axis.MajorGridlineStyle, worksheetId, chartId, axisName, "major gridline");
        if (category)
        {
            if (axis.HasMinimum || axis.HasMaximum || axis.HasMajorUnit) throw Invalid(worksheetId, chartId, $"{axisName}-axis cannot carry numeric minimum, maximum, or major unit.");
            if (axis.HasTickLabelInterval && axis.TickLabelInterval is < 1 or > MaxTickLabelInterval) throw Invalid(worksheetId, chartId, $"{axisName}-axis tick label interval must be 1 through {MaxTickLabelInterval}.");
            return;
        }
        if (axis.HasTickLabelInterval) throw Invalid(worksheetId, chartId, $"{axisName}-axis cannot carry a category tick label interval.");
        if (axis.HasMinimum && !double.IsFinite(axis.Minimum) || axis.HasMaximum && !double.IsFinite(axis.Maximum)) throw Invalid(worksheetId, chartId, $"{axisName}-axis minimum and maximum must be finite.");
        if (axis.HasMinimum && axis.HasMaximum && axis.Minimum >= axis.Maximum) throw Invalid(worksheetId, chartId, $"{axisName}-axis minimum must be less than maximum.");
        if (axis.HasMajorUnit && (!double.IsFinite(axis.MajorUnit) || axis.MajorUnit <= 0)) throw Invalid(worksheetId, chartId, $"{axisName}-axis major unit must be finite and positive.");
    }

    private static bool TryLocate(XElement plotArea, XElement plot, bool numericX, string horizontalPosition, string verticalPosition, out XElement horizontalAxis, out XElement verticalAxis)
    {
        horizontalAxis = null!;
        verticalAxis = null!;
        var plotIds = plot.Elements(ChartNs + "axId").Select(AxisValue).ToArray();
        if (plotIds.Length != 2 || plotIds.Any(string.IsNullOrEmpty) || plotIds.Distinct(StringComparer.Ordinal).Count() != 2) return false;
        var axes = plotArea.Elements().Where(IsAxis)
            .Where(item => plotIds.Contains(AxisValue(item.Element(ChartNs + "axId")), StringComparer.Ordinal))
            .ToArray();
        var categories = axes.Where(item => item.Name == ChartNs + "catAx").ToArray();
        var values = axes.Where(item => item.Name == ChartNs + "valAx").ToArray();
        if (axes.Length != 2) return false;
        if (numericX)
        {
            if (categories.Length != 0 || values.Length != 2) return false;
            var horizontal = values.Where(item => AxisValue(item.Element(ChartNs + "axPos")) == horizontalPosition).ToArray();
            var vertical = values.Where(item => AxisValue(item.Element(ChartNs + "axPos")) == verticalPosition).ToArray();
            if (horizontal.Length != 1 || vertical.Length != 1 || ReferenceEquals(horizontal[0], vertical[0])) return false;
            horizontalAxis = horizontal[0];
            verticalAxis = vertical[0];
        }
        else
        {
            if (categories.Length != 1 || values.Length != 1) return false;
            horizontalAxis = categories[0];
            verticalAxis = values[0];
        }
        var horizontalId = AxisValue(horizontalAxis.Element(ChartNs + "axId"));
        var verticalId = AxisValue(verticalAxis.Element(ChartNs + "axId"));
        if (plotIds.Length != 2 || plotIds.Any(string.IsNullOrEmpty) || plotIds.Distinct(StringComparer.Ordinal).Count() != 2 ||
            string.IsNullOrEmpty(horizontalId) || string.IsNullOrEmpty(verticalId) || horizontalId == verticalId ||
            !plotIds.Contains(horizontalId, StringComparer.Ordinal) || !plotIds.Contains(verticalId, StringComparer.Ordinal)) return false;
        return AxisValue(horizontalAxis.Element(ChartNs + "crossAx")) == verticalId && AxisValue(verticalAxis.Element(ChartNs + "crossAx")) == horizontalId;
    }

    private static bool UsesNumericXAxis(SpreadsheetChartType type) =>
        type is SpreadsheetChartType.Scatter or SpreadsheetChartType.Bubble;

    private static bool TryReadAxis(XElement source, bool category, string expectedPosition, out SpreadsheetChartAxisArtifact axis, out bool editable)
    {
        axis = new SpreadsheetChartAxisArtifact();
        editable = true;
        if (!Singleton(source, "scaling", out var scaling) || scaling is null ||
            !Singleton(scaling, "orientation", out var orientation) || orientation is null ||
            AxisValue(orientation) is not ("minMax" or "maxMin")) return false;
        axis.Reverse = AxisValue(orientation) == "maxMin";
        if (scaling.Element(ChartNs + "logBase") is not null) editable = false;
        if (!Singleton(source, "delete", out var deleted)) return false;
        if (deleted is not null)
        {
            var deletedValue = AxisValue(deleted);
            if (!IsFalse(deletedValue) && !IsTrue(deletedValue)) return false;
            axis.Visible = IsFalse(deletedValue);
        }
        if (!Singleton(source, "axPos", out var position) || position is null) return false;
        if (AxisValue(position) != expectedPosition) editable = false;
        if (!TryTitle(source, out var title, out var titleEditable) || !TryNumberFormat(source, out var numberFormat, out var numberFormatEditable)) return false;
        axis.Title = title;
        axis.NumberFormatCode = numberFormat;
        editable &= titleEditable && numberFormatEditable && XlsxChartTextStyleCodec.TryReadAxis(source, axis);
        if (!TryReadTickLabelVisibility(source, out var tickLabelsVisible, out var tickLabelsEditable)) return false;
        if (tickLabelsVisible is { } labelsVisible) axis.TickLabelsVisible = labelsVisible;
        editable &= tickLabelsEditable;
        if (!TryReadLineContainer(source.Element(ChartNs + "spPr"), out var axisLineVisible, out var axisLine, out var axisLineEditable)) return false;
        if (axisLineVisible is { } visible) axis.AxisLineVisible = visible;
        if (axisLine is not null) axis.AxisLine = axisLine;
        editable &= axisLineEditable;
        if (!Singleton(source, "majorGridlines", out var majorGridlines)) return false;
        if (majorGridlines is not null)
        {
            axis.ShowMajorGridlines = true;
            if (majorGridlines.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)) editable = false;
            if (!TryReadLineContainer(majorGridlines.Element(ChartNs + "spPr"), out var gridVisible, out var gridLine, out var gridEditable)) return false;
            if (majorGridlines.Elements().Any(element => element.Name != ChartNs + "spPr")) editable = false;
            if (gridVisible is { } gridLineVisible) axis.MajorGridlineVisible = gridLineVisible;
            if (gridLine is not null) axis.MajorGridlineStyle = gridLine;
            editable &= gridEditable;
        }
        if (category)
        {
            if (scaling.Element(ChartNs + "min") is not null || scaling.Element(ChartNs + "max") is not null) editable = false;
            if (!TryOptionalUInt(source, "tickLblSkip", out var hasInterval, out var interval) || hasInterval && interval is < 1 or > MaxTickLabelInterval) return false;
            if (hasInterval) axis.TickLabelInterval = interval;
            return true;
        }
        if (!TryOptionalDouble(scaling, "min", out var hasMinimum, out var minimum) ||
            !TryOptionalDouble(scaling, "max", out var hasMaximum, out var maximum) ||
            !TryOptionalDouble(source, "majorUnit", out var hasMajorUnit, out var majorUnit)) return false;
        if (hasMinimum) axis.Minimum = minimum;
        if (hasMaximum) axis.Maximum = maximum;
        if (hasMajorUnit) axis.MajorUnit = majorUnit;
        if (hasMinimum && hasMaximum && minimum >= maximum || hasMajorUnit && majorUnit <= 0) return false;
        return true;
    }

    private static bool TryTitle(XElement source, out string title, out bool editable)
    {
        title = string.Empty;
        editable = true;
        if (!Singleton(source, "title", out var element)) return false;
        if (element is null) return true;
        var richText = element.Descendants(DrawingNs + "t").ToArray();
        if (richText.Length > 0) title = string.Concat(richText.Select(item => item.Value));
        else
        {
            title = element.Descendants(ChartNs + "v").FirstOrDefault()?.Value ?? string.Empty;
            editable = false;
        }
        return title.Length <= 32_767 && !HasControls(title);
    }

    private static bool TryNumberFormat(XElement source, out string code, out bool editable)
    {
        code = string.Empty;
        editable = true;
        if (!Singleton(source, "numFmt", out var element)) return false;
        if (element is null) return true;
        code = (string?)element.Attribute("formatCode") ?? string.Empty;
        if (code.Length > 255 || HasControls(code) || element.Attribute("formatCode") is null) return false;
        var sourceLinked = (string?)element.Attribute("sourceLinked");
        if (sourceLinked is null || !IsFalse(sourceLinked)) editable = false;
        return true;
    }

    private static XElement BuildCategoryAxis(SpreadsheetChartAxisArtifact axis, string axisId = "1", string crossAxisId = "2", string position = "b")
    {
        var output = new XElement(ChartNs + "catAx",
            new XElement(ChartNs + "axId", new XAttribute("val", axisId)),
            new XElement(ChartNs + "scaling", new XElement(ChartNs + "orientation", new XAttribute("val", axis.HasReverse && axis.Reverse ? "maxMin" : "minMax"))),
            new XElement(ChartNs + "axPos", new XAttribute("val", position)));
        if (axis.HasVisible) output.Element(ChartNs + "axPos")!.AddBeforeSelf(new XElement(ChartNs + "delete", new XAttribute("val", axis.Visible ? "0" : "1")));
        if (axis.HasShowMajorGridlines && axis.ShowMajorGridlines)
            output.Add(new XElement(ChartNs + "majorGridlines",
                GridLineProperties(axis)));
        AppendTitleAndNumberFormat(output, axis);
        if (axis.HasTickLabelsVisible && !axis.TickLabelsVisible) output.Add(ValueElement("tickLblPos", "none"));
        if (AxisLineProperties(axis) is { } shapeProperties) output.Add(shapeProperties);
        XlsxChartTextStyleCodec.AppendAuthoredAxis(output, axis.TextStyle);
        output.Add(new XElement(ChartNs + "crossAx", new XAttribute("val", crossAxisId)));
        if (axis.HasTickLabelInterval) output.Add(ValueElement("tickLblSkip", axis.TickLabelInterval));
        return output;
    }

    private static XElement BuildValueAxis(SpreadsheetChartAxisArtifact axis, string axisId, string crossAxisId, string position, string? crosses = null)
    {
        var scaling = new XElement(ChartNs + "scaling", new XElement(ChartNs + "orientation", new XAttribute("val", axis.HasReverse && axis.Reverse ? "maxMin" : "minMax")));
        if (axis.HasMaximum) scaling.Add(ValueElement("max", axis.Maximum));
        if (axis.HasMinimum) scaling.Add(ValueElement("min", axis.Minimum));
        var output = new XElement(ChartNs + "valAx",
            new XElement(ChartNs + "axId", new XAttribute("val", axisId)), scaling,
            new XElement(ChartNs + "axPos", new XAttribute("val", position)));
        if (axis.HasVisible) output.Element(ChartNs + "axPos")!.AddBeforeSelf(new XElement(ChartNs + "delete", new XAttribute("val", axis.Visible ? "0" : "1")));
        if (axis.HasShowMajorGridlines && axis.ShowMajorGridlines)
            output.Add(new XElement(ChartNs + "majorGridlines",
                GridLineProperties(axis)));
        AppendTitleAndNumberFormat(output, axis);
        if (axis.HasTickLabelsVisible && !axis.TickLabelsVisible) output.Add(ValueElement("tickLblPos", "none"));
        if (AxisLineProperties(axis) is { } shapeProperties) output.Add(shapeProperties);
        XlsxChartTextStyleCodec.AppendAuthoredAxis(output, axis.TextStyle);
        output.Add(new XElement(ChartNs + "crossAx", new XAttribute("val", crossAxisId)));
        if (crosses is not null) output.Add(new XElement(ChartNs + "crosses", new XAttribute("val", crosses)));
        if (axis.HasMajorUnit) output.Add(ValueElement("majorUnit", axis.MajorUnit));
        return output;
    }

    private static void AppendTitleAndNumberFormat(XElement axis, SpreadsheetChartAxisArtifact semantic)
    {
        if (semantic.Title.Length > 0) axis.Add(XlsxChartTextStyleCodec.TitleElement(semantic.Title, semantic.TitleTextStyle));
        if (semantic.NumberFormatCode.Length > 0) axis.Add(NumberFormatElement(semantic.NumberFormatCode));
    }

    private static void PatchAxis(XElement native, SpreadsheetChartAxisArtifact target, bool category)
    {
        var scaling = native.Element(ChartNs + "scaling")!;
        SetRequiredOrientation(scaling, target.HasReverse && target.Reverse ? "maxMin" : "minMax");
        PatchValue(native, "delete", target.HasVisible, target.Visible ? 0 : 1, ["axPos", "majorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "extLst"]);
        PatchMajorGridlines(native, target);
        PatchTitle(native, target.Title, target.TitleTextStyle);
        PatchNumberFormat(native, target.NumberFormatCode);
        PatchTickLabelVisibility(native, target);
        PatchAxisLine(native, target);
        XlsxChartTextStyleCodec.PatchAxis(native, target.TextStyle);
        if (category)
        {
            PatchValue(native, "tickLblSkip", target.HasTickLabelInterval, target.TickLabelInterval, ["tickMarkSkip", "noMultiLvlLbl", "extLst"]);
            return;
        }
        PatchValue(scaling, "max", target.HasMaximum, target.Maximum, ["min", "extLst"]);
        PatchValue(scaling, "min", target.HasMinimum, target.Minimum, ["extLst"]);
        PatchValue(native, "majorUnit", target.HasMajorUnit, target.MajorUnit, ["minorUnit", "dispUnits", "extLst"]);
    }

    private static void PatchMajorGridlines(XElement axis, SpreadsheetChartAxisArtifact target)
    {
        var existing = axis.Element(ChartNs + "majorGridlines");
        if (!target.HasShowMajorGridlines || !target.ShowMajorGridlines)
        {
            existing?.Remove();
            return;
        }
        if (existing is not null)
        {
            if (existing.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
                existing.Elements().Any(element => element.Name != ChartNs + "spPr"))
                throw new CodecException("unsupported_spreadsheet_chart_edit", "Chart major gridlines use an unmodeled style graph.");
            PatchLineContainer(existing, target.MajorGridlineStyle,
                !target.HasMajorGridlineVisible || target.MajorGridlineVisible);
            return;
        }
        InsertBefore(axis, new XElement(ChartNs + "majorGridlines",
            GridLineProperties(target)), ["title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "extLst"]);
    }

    private static bool TryReadTickLabelVisibility(
        XElement source,
        out bool? visible,
        out bool editable)
    {
        visible = null;
        editable = true;
        if (!Singleton(source, "tickLblPos", out var element)) return false;
        if (element is null) return true;
        var value = AxisValue(element);
        if (value == "none")
        {
            visible = false;
            return true;
        }
        if (value == "nextTo") return true;
        if (value is "high" or "low")
        {
            editable = false;
            return true;
        }
        return false;
    }

    private static void PatchTickLabelVisibility(XElement axis, SpreadsheetChartAxisArtifact target)
    {
        if (!target.HasTickLabelsVisible) return;
        var existing = axis.Element(ChartNs + "tickLblPos");
        if (target.TickLabelsVisible)
        {
            existing?.Remove();
            return;
        }
        if (existing is null)
            InsertBefore(axis, ValueElement("tickLblPos", "none"), ["spPr", "txPr", "crossAx", "crosses", "crossesAt", "extLst"]);
        else
            existing.SetAttributeValue("val", "none");
    }

    private static bool TryReadLineContainer(
        XElement? shapeProperties,
        out bool? visible,
        out SpreadsheetChartLineStyleArtifact? line,
        out bool editable)
    {
        visible = null;
        line = null;
        editable = true;
        if (shapeProperties is null) return true;
        if (shapeProperties.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
            shapeProperties.Elements().Any(element => element.Name != DrawingNs + "ln"))
        {
            editable = false;
            return true;
        }
        var lines = shapeProperties.Elements(DrawingNs + "ln").Take(2).ToArray();
        if (lines.Length != 1)
        {
            editable = false;
            return true;
        }
        var nativeLine = lines[0];
        if (nativeLine.Attributes().All(attribute => attribute.IsNamespaceDeclaration) &&
            nativeLine.Elements().Count() == 1 && nativeLine.Element(DrawingNs + "noFill") is { } noFill &&
            !noFill.HasElements && noFill.Attributes().All(attribute => attribute.IsNamespaceDeclaration))
        {
            visible = false;
            return true;
        }
        if (!XlsxChartSeriesLineStyleCodec.TryReadLine(shapeProperties, out var parsed) ||
            parsed?.Color is null || !parsed.HasWidthPoints)
        {
            editable = false;
            return true;
        }
        visible = true;
        line = parsed;
        return true;
    }

    private static XElement? AxisLineProperties(SpreadsheetChartAxisArtifact axis)
    {
        if (axis.AxisLine is not null) return LineProperties(axis.AxisLine);
        if (axis.HasAxisLineVisible && !axis.AxisLineVisible) return HiddenLineProperties();
        return null;
    }

    private static XElement? GridLineProperties(SpreadsheetChartAxisArtifact axis)
    {
        if (axis.MajorGridlineStyle is not null) return LineProperties(axis.MajorGridlineStyle);
        if (axis.HasMajorGridlineVisible && !axis.MajorGridlineVisible) return HiddenLineProperties();
        return null;
    }

    private static XElement LineProperties(SpreadsheetChartLineStyleArtifact line) =>
        new(ChartNs + "spPr", XlsxChartSeriesLineStyleCodec.Element(line));

    private static XElement HiddenLineProperties() =>
        new(ChartNs + "spPr", new XElement(DrawingNs + "ln", new XElement(DrawingNs + "noFill")));

    private static void PatchAxisLine(XElement axis, SpreadsheetChartAxisArtifact target)
    {
        var visible = target.AxisLine is not null || !target.HasAxisLineVisible || target.AxisLineVisible;
        PatchLineContainer(axis, target.AxisLine, visible);
    }

    private static void PatchLineContainer(XElement owner, SpreadsheetChartLineStyleArtifact? line, bool visible)
    {
        var existing = owner.Element(ChartNs + "spPr");
        if (LineContainerMatches(existing, line, visible)) return;
        XElement? replacement = line is not null ? LineProperties(line) : visible ? null : HiddenLineProperties();
        if (replacement is null) { existing?.Remove(); return; }
        if (existing is not null) { existing.ReplaceWith(replacement); return; }
        InsertBefore(owner, replacement, ["txPr", "crossAx", "crosses", "crossesAt", "extLst"]);
    }

    private static bool LineContainerMatches(XElement? existing, SpreadsheetChartLineStyleArtifact? line, bool visible)
    {
        if (!TryReadLineContainer(existing, out var existingVisible, out var existingLine, out var editable) || !editable)
            return false;
        if (line is not null)
            return existingVisible == true &&
                   XlsxChartSeriesLineStyleCodec.Semantics(existingLine) == XlsxChartSeriesLineStyleCodec.Semantics(line);
        return visible ? existing is null : existingVisible == false;
    }

    private static void SetRequiredOrientation(XElement scaling, string value)
    {
        var orientation = scaling.Element(ChartNs + "orientation")
            ?? throw new CodecException("unsupported_spreadsheet_chart_edit", "Chart axis has no canonical orientation node.");
        orientation.SetAttributeValue("val", value);
    }

    private static void PatchTitle(XElement owner, string title, SpreadsheetChartTextStyleArtifact? style)
    {
        var existing = owner.Element(ChartNs + "title");
        if (title.Length == 0) { existing?.Remove(); return; }
        if (existing is null)
        {
            InsertBefore(owner, XlsxChartTextStyleCodec.TitleElement(title, style), ["numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "extLst"]);
            return;
        }
        var runs = existing.Descendants(DrawingNs + "t").ToArray();
        if (runs.Length == 0) throw new CodecException("unsupported_spreadsheet_chart_edit", "Referenced worksheet-chart axis titles are read-only.");
        runs[0].Value = title;
        foreach (var run in runs.Skip(1)) run.Value = string.Empty;
        XlsxChartTextStyleCodec.PatchTitle(existing, style);
    }

    private static void PatchNumberFormat(XElement owner, string code)
    {
        var existing = owner.Element(ChartNs + "numFmt");
        if (code.Length == 0) { existing?.Remove(); return; }
        if (existing is null)
        {
            InsertBefore(owner, NumberFormatElement(code), ["majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "extLst"]);
            return;
        }
        existing.SetAttributeValue("formatCode", code);
        existing.SetAttributeValue("sourceLinked", "0");
    }

    private static void PatchValue<T>(XElement owner, string name, bool present, T value, string[] laterNames)
    {
        var existing = owner.Element(ChartNs + name);
        if (!present) { existing?.Remove(); return; }
        if (existing is null) InsertBefore(owner, ValueElement(name, value), laterNames);
        else existing.SetAttributeValue("val", Format(value));
    }

    private static void InsertBefore(XElement owner, XElement value, IEnumerable<string> laterNames)
    {
        var later = new HashSet<XName>(laterNames.Select(name => ChartNs + name));
        var next = owner.Elements().FirstOrDefault(item => later.Contains(item.Name));
        if (next is null) owner.Add(value);
        else next.AddBeforeSelf(value);
    }

    private static XElement NumberFormatElement(string code) => new(ChartNs + "numFmt", new XAttribute("formatCode", code), new XAttribute("sourceLinked", "0"));
    private static XElement ValueElement<T>(string name, T value) => new(ChartNs + name, new XAttribute("val", Format(value)));
    private static string Format<T>(T value) => value switch { double number => number.ToString("R", CultureInfo.InvariantCulture), IFormattable item => item.ToString(null, CultureInfo.InvariantCulture), _ => value?.ToString() ?? string.Empty };

    private static bool TryOptionalUInt(XElement source, string name, out bool present, out uint value)
    {
        present = false; value = 0;
        if (!Singleton(source, name, out var element)) return false;
        if (element is null) return true;
        present = true;
        return uint.TryParse(AxisValue(element), NumberStyles.None, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryOptionalDouble(XElement source, string name, out bool present, out double value)
    {
        present = false; value = 0;
        if (!Singleton(source, name, out var element)) return false;
        if (element is null) return true;
        present = true;
        return double.TryParse(AxisValue(element), NumberStyles.Float, CultureInfo.InvariantCulture, out value) && double.IsFinite(value);
    }

    private static bool Singleton(XElement source, string name, out XElement? element)
    {
        var matches = source.Elements(ChartNs + name).Take(2).ToArray();
        element = matches.FirstOrDefault();
        return matches.Length <= 1;
    }

    private static string AxisValue(XElement? source) => (string?)source?.Attribute("val") ?? string.Empty;
    private static bool IsAxis(XElement source) => source.Name == ChartNs + "catAx" || source.Name == ChartNs + "dateAx" || source.Name == ChartNs + "valAx" || source.Name == ChartNs + "serAx";
    private static bool IsFalse(string value) => value is "0" or "false" or "off";
    private static bool IsTrue(string value) => value is "1" or "true" or "on";
    private static bool HasControls(string value) => value.Any(char.IsControl);
    private static string AxisSemantics(SpreadsheetChartAxisArtifact? axis) => axis is null ? "-" : string.Join('\u001f', axis.Title, axis.NumberFormatCode,
        axis.HasTickLabelInterval ? axis.TickLabelInterval.ToString(CultureInfo.InvariantCulture) : "-",
        axis.HasMinimum ? axis.Minimum.ToString("R", CultureInfo.InvariantCulture) : "-",
        axis.HasMaximum ? axis.Maximum.ToString("R", CultureInfo.InvariantCulture) : "-",
        axis.HasMajorUnit ? axis.MajorUnit.ToString("R", CultureInfo.InvariantCulture) : "-",
        axis.HasVisible ? (axis.Visible ? "visible" : "hidden") : "default-visible",
        axis.HasReverse ? (axis.Reverse ? "reverse" : "forward") : "default-direction",
        axis.HasAxisLineVisible ? (axis.AxisLineVisible ? "axis-line" : "no-axis-line") : "default-axis-line",
        XlsxChartSeriesLineStyleCodec.Semantics(axis.AxisLine),
        axis.HasShowMajorGridlines ? (axis.ShowMajorGridlines ? "gridlines" : "no-gridlines") : "default-gridlines",
        axis.HasMajorGridlineVisible ? (axis.MajorGridlineVisible ? "visible-gridline" : "hidden-gridline") : "default-gridline-visibility",
        XlsxChartSeriesLineStyleCodec.Semantics(axis.MajorGridlineStyle),
        axis.HasTickLabelsVisible ? (axis.TickLabelsVisible ? "tick-labels" : "no-tick-labels") : "default-tick-labels",
        XlsxChartTextStyleCodec.Semantics(axis.TextStyle),
        XlsxChartTextStyleCodec.Semantics(axis.TitleTextStyle));
    private static CodecException Invalid(string worksheetId, string chartId, string message) => new("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chartId} {message}");
}
