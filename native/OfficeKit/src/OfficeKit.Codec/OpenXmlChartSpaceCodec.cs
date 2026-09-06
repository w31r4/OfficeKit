using System.Globalization;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns the package-agnostic DrawingML ChartSpace profile shared by XLSX and
// PPTX. Callers retain package relationships, source bindings, anchors, and
// Presentation-only combo topology; this module owns one ordinary plot,
// literal/reference caches, styling, labels, and the paired primary axes.
internal static class OpenXmlChartSpaceCodec
{
    private const int MaxSeries = 256;
    private const int MaxPoints = 1_048_576;
    private static readonly string[] DisplayBlanksAsValues = ["zero", "gap", "span"];
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static bool TryRead(
        string xml,
        out SpreadsheetChartArtifact chart,
        out XDocument document,
        out bool editable,
        bool allowRichTitle = false,
        bool allowChartFrameDecorations = false)
    {
        chart = new SpreadsheetChartArtifact();
        editable = true;
        try { document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace); }
        catch (System.Xml.XmlException) { document = new XDocument(); return false; }
        var root = document.Root;
        var nativeChart = root?.Element(ChartNs + "chart");
        var plotArea = nativeChart?.Element(ChartNs + "plotArea");
        if (root?.Name != ChartNs + "chartSpace" || nativeChart is null || plotArea is null || root.Element(ChartNs + "externalData") is not null) return false;
        var plots = plotArea.Elements().Where(item => item.Name == ChartNs + "barChart" || item.Name == ChartNs + "lineChart" || item.Name == ChartNs + "pieChart" || item.Name == ChartNs + "areaChart" || item.Name == ChartNs + "doughnutChart" || item.Name == ChartNs + "scatterChart" || item.Name == ChartNs + "bubbleChart" || item.Name == ChartNs + "radarChart").ToArray();
        if (plots.Length != 1 || plotArea.Elements().Any(item => item.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal) && !plots.Contains(item))) return false;
        var plot = plots[0];
        chart.Type = plot.Name.LocalName switch
        {
            "barChart" => SpreadsheetChartType.Bar,
            "lineChart" => SpreadsheetChartType.Line,
            "pieChart" => SpreadsheetChartType.Pie,
            "areaChart" => SpreadsheetChartType.Area,
            "doughnutChart" => SpreadsheetChartType.Doughnut,
            "scatterChart" => SpreadsheetChartType.Scatter,
            "bubbleChart" => SpreadsheetChartType.Bubble,
            "radarChart" => SpreadsheetChartType.Radar,
            _ => SpreadsheetChartType.Unspecified,
        };
        if (chart.Type == SpreadsheetChartType.Unspecified) return false;
        editable &= PlotProfileEditable(plot, chart.Type);
        editable &= TryReadPlotOptions(plot, chart);
        if (!TryReadDisplayBlanksAs(nativeChart, out var displayBlanksAs)) editable = false;
        else if (displayBlanksAs is not null) chart.DisplayBlanksAs = displayBlanksAs;
        var title = nativeChart.Element(ChartNs + "title");
        if (title is not null)
        {
            var richText = title.Descendants(DrawingNs + "t").ToArray();
            var directValue = title.Descendants(ChartNs + "v").FirstOrDefault();
            chart.Title = richText.Length > 0 ? string.Concat(richText.Select(item => item.Value)) : directValue?.Value ?? string.Empty;
            if (richText.Length == 0) editable = false;
            if (!XlsxChartTextStyleCodec.TryReadTitle(title, chart) && !allowRichTitle) editable = false;
        }
        var legend = nativeChart.Element(ChartNs + "legend");
        chart.HasLegend = legend is not null;
        if (legend is not null)
        {
            if (!TryReadLegend(legend, chart)) editable = false;
        }
        var nativeSeries = plot.Elements(ChartNs + "ser").ToArray();
        if (nativeSeries.Length is < 1 or > MaxSeries) return false;
        string[]? commonCategories = null;
        foreach (var native in nativeSeries)
        {
            if (!TrySeries(native, chart.Type, out var series, out var categories, out var seriesEditable)) return false;
            editable &= seriesEditable;
            if (!UsesNumericXAxis(chart.Type))
            {
                if (commonCategories is null) commonCategories = categories;
                else if (!commonCategories.SequenceEqual(categories, StringComparer.Ordinal)) return false;
            }
            chart.Series.Add(series);
        }
        if (!UsesNumericXAxis(chart.Type)) chart.Categories.Add(commonCategories ?? []);
        editable &= XlsxChartLineOptionsCodec.TryRead(plot, chart);
        editable &= XlsxChartDataLabelsCodec.TryRead(plot, chart);
        if (!XlsxChartAxisCodec.TryRead(plotArea, plot, chart, out var axesEditable)) editable = false;
        else editable &= axesEditable;
        var chartSpaceProperties = root.Element(ChartNs + "spPr");
        if (!XlsxChartSurfaceFillCodec.TryRead(chartSpaceProperties, out var chartAreaFill, allowChartFrameDecorations))
        {
            // Presentation chart frame image fills are owned by the
            // Presentation-only frame codec because their relationship must
            // resolve against ChartPart rather than a worksheet drawing.
            // Leave that branch for PptxChartFrameCodec to inspect.
            if (!(allowChartFrameDecorations && chartSpaceProperties?.Elements(DrawingNs + "blipFill").Any() == true))
                editable = false;
        }
        else if (chartAreaFill is not null) chart.ChartAreaFill = chartAreaFill;
        if (!XlsxChartSurfaceFillCodec.TryRead(plotArea.Element(ChartNs + "spPr"), out var plotAreaFill)) editable = false;
        else if (plotAreaFill is not null) chart.PlotAreaFill = plotAreaFill;
        return chart.Title.Length <= 32_767 && !HasControls(chart.Title) && chart.Categories.Count <= MaxPoints;
    }

    internal static XDocument Build(SpreadsheetChartArtifact chart)
    {
        var series = chart.Series.Select((item, index) => SeriesElement(item, chart.Categories, index, chart.Type)).ToArray();
        XElement plot = chart.Type switch
        {
            SpreadsheetChartType.Bar => new XElement(ChartNs + "barChart", new XElement(ChartNs + "barDir", new XAttribute("val", BarDirectionToken(chart.BarDirection))), new XElement(ChartNs + "grouping", new XAttribute("val", GroupingToken(chart.Grouping, clustered: true))), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), chart.HasGapWidth ? new XElement(ChartNs + "gapWidth", new XAttribute("val", chart.GapWidth)) : null, new XElement(ChartNs + "axId", new XAttribute("val", "1")), new XElement(ChartNs + "axId", new XAttribute("val", "2"))),
            SpreadsheetChartType.Line => new XElement(ChartNs + "lineChart", XlsxChartLineOptionsCodec.GroupingElement(LineOptions(chart)), XlsxChartLineOptionsCodec.VaryColorsElement(chart.LineOptions), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), XlsxChartLineOptionsCodec.SmoothElement(chart.LineOptions), new XElement(ChartNs + "axId", new XAttribute("val", "1")), new XElement(ChartNs + "axId", new XAttribute("val", "2"))),
            SpreadsheetChartType.Area => new XElement(ChartNs + "areaChart", new XElement(ChartNs + "grouping", new XAttribute("val", GroupingToken(chart.Grouping, clustered: false))), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), new XElement(ChartNs + "axId", new XAttribute("val", "1")), new XElement(ChartNs + "axId", new XAttribute("val", "2"))),
            SpreadsheetChartType.Doughnut => new XElement(ChartNs + "doughnutChart", new XElement(ChartNs + "varyColors", new XAttribute("val", "1")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), new XElement(ChartNs + "firstSliceAng", new XAttribute("val", chart.HasFirstSliceAngle ? chart.FirstSliceAngle : 0U)), new XElement(ChartNs + "holeSize", new XAttribute("val", chart.HasDoughnutHoleSize ? chart.DoughnutHoleSize : 50U))),
            SpreadsheetChartType.Scatter => new XElement(ChartNs + "scatterChart", new XElement(ChartNs + "scatterStyle", new XAttribute("val", "marker")), new XElement(ChartNs + "varyColors", new XAttribute("val", "0")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), new XElement(ChartNs + "axId", new XAttribute("val", "1")), new XElement(ChartNs + "axId", new XAttribute("val", "2"))),
            SpreadsheetChartType.Bubble => new XElement(ChartNs + "bubbleChart", new XElement(ChartNs + "varyColors", new XAttribute("val", "0")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), new XElement(ChartNs + "bubble3D", new XAttribute("val", "0")), new XElement(ChartNs + "bubbleScale", new XAttribute("val", chart.HasBubbleScale ? chart.BubbleScale : 100U)), new XElement(ChartNs + "showNegBubbles", new XAttribute("val", "0")), new XElement(ChartNs + "sizeRepresents", new XAttribute("val", BubbleSizeModeToken(chart.BubbleSizeMode))), new XElement(ChartNs + "axId", new XAttribute("val", "1")), new XElement(ChartNs + "axId", new XAttribute("val", "2"))),
            SpreadsheetChartType.Radar => new XElement(ChartNs + "radarChart", new XElement(ChartNs + "radarStyle", new XAttribute("val", "standard")), new XElement(ChartNs + "varyColors", new XAttribute("val", "0")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), new XElement(ChartNs + "axId", new XAttribute("val", "1")), new XElement(ChartNs + "axId", new XAttribute("val", "2"))),
            SpreadsheetChartType.Pie => new XElement(ChartNs + "pieChart", new XElement(ChartNs + "varyColors", new XAttribute("val", "1")), series, XlsxChartDataLabelsCodec.Element(chart.DataLabels), chart.HasFirstSliceAngle ? new XElement(ChartNs + "firstSliceAng", new XAttribute("val", chart.FirstSliceAngle)) : null),
            _ => throw new InvalidOperationException("Validated chart type is unsupported."),
        };
        var plotArea = new XElement(ChartNs + "plotArea", new XElement(ChartNs + "layout"), plot);
        XlsxChartAxisCodec.AppendAuthored(plotArea, chart);
        if (XlsxChartSurfaceFillCodec.Element(chart.PlotAreaFill, "Chart plot area") is { } plotFill) plotArea.Add(plotFill);
        var nativeChart = new XElement(ChartNs + "chart");
        if (chart.Title.Length > 0) nativeChart.Add(XlsxChartTextStyleCodec.TitleElement(chart.Title, chart.TitleTextStyle));
        nativeChart.Add(plotArea);
        if (chart.HasLegend) nativeChart.Add(LegendElement(chart.LegendPosition, chart.LegendTextStyle, chart.HasLegendOverlay, chart.LegendOverlay));
        nativeChart.Add(new XElement(ChartNs + "plotVisOnly", new XAttribute("val", "1")));
        if (chart.HasDisplayBlanksAs)
            nativeChart.Add(new XElement(ChartNs + "dispBlanksAs", new XAttribute("val", DisplayBlanksAsToken(chart.DisplayBlanksAs))));
        var chartSpace = new XElement(ChartNs + "chartSpace", new XAttribute(XNamespace.Xmlns + "c", ChartNs), new XAttribute(XNamespace.Xmlns + "a", DrawingNs), nativeChart);
        if (XlsxChartSurfaceFillCodec.Element(chart.ChartAreaFill, "Chart area") is { } chartFill) chartSpace.Add(chartFill);
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), chartSpace);
    }

    internal static void Patch(
        XDocument document,
        SpreadsheetChartArtifact target,
        string errorCode,
        string subject,
        bool patchTitle = true,
        bool allowChartFrameDecorations = false)
    {
        var nativeChart = document.Root?.Element(ChartNs + "chart") ?? throw Topology(errorCode, subject, "is missing c:chart");
        if (patchTitle) PatchTitle(nativeChart, target.Title, target.TitleTextStyle, errorCode, subject);
        PatchLegend(nativeChart, target.HasLegend, target.LegendPosition, target.LegendTextStyle, target.HasLegendOverlay, target.LegendOverlay);
        PatchDisplayBlanksAs(nativeChart, target.HasDisplayBlanksAs, target.DisplayBlanksAs, errorCode, subject);
        var plotArea = nativeChart.Element(ChartNs + "plotArea") ?? throw Topology(errorCode, subject, "is missing c:plotArea");
        var plotName = target.Type switch
        {
            SpreadsheetChartType.Bar => "barChart",
            SpreadsheetChartType.Line => "lineChart",
            SpreadsheetChartType.Pie => "pieChart",
            SpreadsheetChartType.Area => "areaChart",
            SpreadsheetChartType.Doughnut => "doughnutChart",
            SpreadsheetChartType.Scatter => "scatterChart",
            SpreadsheetChartType.Bubble => "bubbleChart",
            SpreadsheetChartType.Radar => "radarChart",
            _ => throw new InvalidOperationException("Validated chart type is unsupported."),
        };
        var plot = plotArea.Element(ChartNs + plotName) ?? throw Topology(errorCode, subject, $"is missing c:{plotName}");
        var nativeSeries = plot.Elements(ChartNs + "ser").ToArray();
        if (nativeSeries.Length != target.Series.Count) throw Topology(errorCode, subject, "series topology changed unexpectedly");
        for (var index = 0; index < nativeSeries.Length; index++) PatchSeries(nativeSeries[index], target.Series[index], target.Categories, target.Type, errorCode, subject);
        if (target.Type == SpreadsheetChartType.Line) XlsxChartLineOptionsCodec.Patch(plot, LineOptions(target));
        PatchPlotOptions(plot, target);
        XlsxChartDataLabelsCodec.Patch(plot, target.DataLabels);
        XlsxChartAxisCodec.Patch(plotArea, plot, target);
        XlsxChartSurfaceFillCodec.Patch(document.Root!, target.ChartAreaFill, $"{subject} chart area", allowChartFrameDecorations);
        XlsxChartSurfaceFillCodec.Patch(plotArea, target.PlotAreaFill, $"{subject} plot area");
    }

    internal static bool TrySeries(XElement source, SpreadsheetChartType chartType, out SpreadsheetChartSeriesArtifact series, out string[] categories, out bool editable)
    {
        series = new SpreadsheetChartSeriesArtifact(); categories = []; editable = true;
        var tx = source.Element(ChartNs + "tx");
        var directName = tx?.Element(ChartNs + "v");
        if (directName is not null) series.Name = directName.Value;
        else
        {
            var reference = tx?.Element(ChartNs + "strRef");
            var names = reference is null ? null : ReadStringPoints(reference.Element(ChartNs + "strCache"));
            if (names is null || names.Length != 1) return false;
            series.Name = names[0];
            editable = false;
        }
        if (series.Name.Length > 255 || HasControls(series.Name)) return false;
        if (UsesNumericXAxis(chartType))
        {
            var xValue = source.Element(ChartNs + "xVal");
            var yValue = source.Element(ChartNs + "yVal");
            if (xValue is null || yValue is null ||
                !TryNumericData(xValue, allowMissing: false, out var xValues, out _, out var xFormula) ||
                !TryNumericData(yValue, allowMissing: true, out var values, out var missingValueIndexes, out var valueFormula) ||
                xValues.Length != values.Length || values.Length > MaxPoints) return false;
            series.XValueFormula = xFormula;
            series.ValueFormula = valueFormula;
            series.XValues.Add(xValues);
            series.Values.Add(values);
            series.MissingValueIndexes.Add(missingValueIndexes);
            if (chartType == SpreadsheetChartType.Bubble)
            {
                var bubbleSize = source.Element(ChartNs + "bubbleSize");
                if (bubbleSize is null || !TryNumericData(bubbleSize, allowMissing: false, out var bubbleSizes, out _, out var bubbleFormula) || bubbleSizes.Length != values.Length || bubbleSizes.Any(value => value <= 0)) return false;
                series.BubbleSizeFormula = bubbleFormula;
                series.BubbleSizes.Add(bubbleSizes);
                editable &= ScalarEquals(source, "bubble3D", "0", required: false);
            }
        }
        else
        {
            var category = source.Element(ChartNs + "cat");
            var value = source.Element(ChartNs + "val");
            if (category is null || value is null || !TryStringData(category, out categories, out var categoryFormula) ||
                !TryNumericData(value, allowMissing: true, out var values, out var missingValueIndexes, out var valueFormula) ||
                categories.Length != values.Length || categories.Length > MaxPoints) return false;
            series.CategoryFormula = categoryFormula;
            series.ValueFormula = valueFormula;
            series.Values.Add(values);
            series.MissingValueIndexes.Add(missingValueIndexes);
        }
        editable &= XlsxChartSeriesStyleCodec.TryRead(source, series);
        editable &= XlsxChartSeriesLineStyleCodec.TryRead(source, series, chartType);
        editable &= XlsxChartSeriesMarkerCodec.TryRead(source, series, chartType);
        editable &= XlsxChartSeriesDataLabelsCodec.TryRead(source, series, chartType);
        editable &= XlsxChartPointStyleCodec.TryRead(source, series, chartType);
        editable &= OpenXmlChartTrendlineCodec.TryRead(source, series, chartType);
        editable &= OpenXmlChartErrorBarsCodec.TryRead(source, series, chartType);
        return true;
    }

    internal static XElement SeriesElement(SpreadsheetChartSeriesArtifact series, IEnumerable<string> categories, int index, SpreadsheetChartType chartType)
    {
        var output = new XElement(ChartNs + "ser",
            new XElement(ChartNs + "idx", new XAttribute("val", index)),
            new XElement(ChartNs + "order", new XAttribute("val", index)),
            new XElement(ChartNs + "tx", new XElement(ChartNs + "v", series.Name)),
            XlsxChartSeriesStyleCodec.PropertiesElement(series, markerOnly: chartType == SpreadsheetChartType.Scatter),
            XlsxChartSeriesMarkerCodec.Element(series.Marker),
            XlsxChartPointStyleCodec.Elements(series),
            XlsxChartSeriesDataLabelsCodec.Element(series.DataLabels),
            OpenXmlChartTrendlineCodec.Elements(series.Trendlines),
            OpenXmlChartErrorBarsCodec.Element(series.ErrorBars));
        if (UsesNumericXAxis(chartType))
        {
            output.Add(
                new XElement(ChartNs + "xVal", NumericData(series.XValues, [], series.XValueFormula)),
                new XElement(ChartNs + "yVal", NumericData(series.Values, series.MissingValueIndexes, series.ValueFormula)));
            if (chartType == SpreadsheetChartType.Bubble) output.Add(new XElement(ChartNs + "bubbleSize", NumericData(series.BubbleSizes, [], series.BubbleSizeFormula)));
        }
        else output.Add(
            new XElement(ChartNs + "cat", StringData(categories, series.CategoryFormula)),
            new XElement(ChartNs + "val", NumericData(series.Values, series.MissingValueIndexes, series.ValueFormula)));
        return output;
    }

    internal static void PatchTitle(XElement chart, string title, SpreadsheetChartTextStyleArtifact? style, string errorCode, string subject)
    {
        var existing = chart.Element(ChartNs + "title");
        if (title.Length == 0) { existing?.Remove(); return; }
        if (existing is null)
        {
            var plotArea = chart.Element(ChartNs + "plotArea") ?? throw Topology(errorCode, subject, "is missing c:plotArea");
            plotArea.AddBeforeSelf(XlsxChartTextStyleCodec.TitleElement(title, style));
            return;
        }
        var runs = existing.Descendants(DrawingNs + "t").ToArray();
        if (runs.Length == 0) throw Topology(errorCode, subject, "has a title outside the editable rich-text profile");
        runs[0].Value = title;
        foreach (var run in runs.Skip(1)) run.Value = string.Empty;
        XlsxChartTextStyleCodec.PatchTitle(existing, style);
    }

    internal static void PatchLegend(
        XElement chart,
        bool hasLegend,
        string position = "",
        SpreadsheetChartTextStyleArtifact? textStyle = null,
        bool hasOverlay = false,
        bool overlay = false)
    {
        var legend = chart.Element(ChartNs + "legend");
        if (!hasLegend) { legend?.Remove(); return; }
        if (legend is null)
        {
            chart.Element(ChartNs + "plotArea")!.AddAfterSelf(LegendElement(position, textStyle, hasOverlay, overlay));
            return;
        }
        if (!TryLegendPosition(legend, out _) || !TryLegendOverlay(legend, out _, out _))
            throw Topology("unsupported_chart_edit", "Chart", "has a legend outside the canonical placement profile");
        legend.Element(ChartNs + "legendPos")!.SetAttributeValue("val", LegendPositionToken(position));
        var existingOverlay = legend.Element(ChartNs + "overlay");
        if (hasOverlay)
        {
            var value = overlay ? "1" : "0";
            if (existingOverlay is null)
                legend.Element(ChartNs + "layout")!.AddAfterSelf(new XElement(ChartNs + "overlay", new XAttribute("val", value)));
            else
                existingOverlay.SetAttributeValue("val", value);
        }
        else
            existingOverlay?.Remove();
        XlsxChartTextStyleCodec.PatchTextProperties(legend, textStyle, new HashSet<string>(StringComparer.Ordinal) { "extLst" });
    }

    internal static void PatchSeries(XElement native, SpreadsheetChartSeriesArtifact target, IEnumerable<string> categories, SpreadsheetChartType chartType, string errorCode, string subject)
    {
        var name = native.Element(ChartNs + "tx")?.Element(ChartNs + "v") ?? throw Topology(errorCode, subject, "series name topology changed unexpectedly");
        name.Value = target.Name;
        XlsxChartSeriesStyleCodec.Patch(native, target);
        XlsxChartSeriesLineStyleCodec.Patch(native, target, markerOnly: chartType == SpreadsheetChartType.Scatter);
        XlsxChartSeriesMarkerCodec.Patch(native, target);
        XlsxChartPointStyleCodec.Patch(native, target, chartType, errorCode, subject);
        XlsxChartSeriesDataLabelsCodec.Patch(native, target, chartType, errorCode, subject);
        OpenXmlChartTrendlineCodec.Patch(native, target, errorCode, subject);
        OpenXmlChartErrorBarsCodec.Patch(native, target, errorCode, subject);
        if (UsesNumericXAxis(chartType))
        {
            PatchNumericData(native.Element(ChartNs + "xVal"), target.XValues, [], target.XValueFormula, errorCode, subject);
            PatchNumericData(native.Element(ChartNs + "yVal"), target.Values, target.MissingValueIndexes, target.ValueFormula, errorCode, subject);
            if (chartType == SpreadsheetChartType.Bubble) PatchNumericData(native.Element(ChartNs + "bubbleSize"), target.BubbleSizes, [], target.BubbleSizeFormula, errorCode, subject);
        }
        else
        {
            PatchStringData(native.Element(ChartNs + "cat"), categories, target.CategoryFormula, errorCode, subject);
            PatchNumericData(native.Element(ChartNs + "val"), target.Values, target.MissingValueIndexes, target.ValueFormula, errorCode, subject);
        }
    }

    internal static XElement LegendElement(
        string position = "",
        SpreadsheetChartTextStyleArtifact? textStyle = null,
        bool hasOverlay = false,
        bool overlay = false) => new(
        ChartNs + "legend",
        new XElement(ChartNs + "legendPos", new XAttribute("val", LegendPositionToken(position))),
        new XElement(ChartNs + "layout"),
        hasOverlay ? new XElement(ChartNs + "overlay", new XAttribute("val", overlay ? "1" : "0")) : null,
        textStyle is null ? null : XlsxChartTextStyleCodec.TextPropertiesElement(textStyle));

    internal static bool TryReadLegend(XElement legend, SpreadsheetChartArtifact chart)
    {
        if (!TryLegendPosition(legend, out var position) ||
            !TryLegendOverlay(legend, out var hasOverlay, out var overlay) ||
            !XlsxChartTextStyleCodec.TryReadTextProperties(legend, out var textStyle)) return false;
        chart.LegendPosition = position;
        if (hasOverlay) chart.LegendOverlay = overlay;
        if (textStyle is not null) chart.LegendTextStyle = textStyle;
        return true;
    }

    internal static bool TryReadDisplayBlanksAs(XElement nativeChart, out string? value) =>
        TryScalar(nativeChart, "dispBlanksAs", DisplayBlanksAsValues, required: false, out value);

    internal static void PatchDisplayBlanksAs(
        XElement nativeChart,
        bool present,
        string value,
        string errorCode,
        string subject)
    {
        var existing = nativeChart.Element(ChartNs + "dispBlanksAs");
        if (!present)
        {
            existing?.Remove();
            return;
        }
        if (existing is not null)
        {
            if (!TryReadDisplayBlanksAs(nativeChart, out var current) || current is null)
                throw Topology(errorCode, subject, "has an invalid c:dispBlanksAs");
        }
        var replacement = new XElement(ChartNs + "dispBlanksAs", new XAttribute("val", DisplayBlanksAsToken(value)));
        if (existing is not null)
        {
            existing.ReplaceWith(replacement);
            return;
        }
        var following = nativeChart.Elements().FirstOrDefault(element =>
            element.Name == ChartNs + "showDLbls" ||
            element.Name == ChartNs + "showDLblsOverMax" ||
            element.Name == ChartNs + "extLst");
        if (following is null) nativeChart.Add(replacement);
        else following.AddBeforeSelf(replacement);
    }

    private static bool TryLegendPosition(XElement legend, out string position)
    {
        position = string.Empty;
        if (legend.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
            legend.Elements().Any(element => element.Name != ChartNs + "legendPos" && element.Name != ChartNs + "layout" && element.Name != ChartNs + "overlay" && element.Name != ChartNs + "txPr") ||
            legend.Elements(ChartNs + "legendPos").Take(2).Count() != 1 ||
            legend.Elements(ChartNs + "layout").Take(2).Count() != 1 ||
            legend.Element(ChartNs + "layout")!.HasElements || legend.Element(ChartNs + "layout")!.HasAttributes)
            return false;
        var nativePosition = legend.Element(ChartNs + "legendPos")!;
        if (nativePosition.HasElements || nativePosition.Attributes().Count() != 1 || nativePosition.Attribute("val") is null)
            return false;
        var native = (string?)nativePosition.Attribute("val");
        position = native switch
        {
            "t" => "top",
            "tr" => "topRight",
            "b" => "bottom",
            "l" => "left",
            "r" => "right",
            _ => string.Empty,
        };
        return position.Length > 0;
    }

    private static bool TryLegendOverlay(XElement legend, out bool hasOverlay, out bool overlay)
    {
        hasOverlay = false;
        overlay = false;
        var elements = legend.Elements(ChartNs + "overlay").Take(2).ToArray();
        if (elements.Length == 0) return true;
        if (elements.Length != 1) return false;
        var element = elements[0];
        if (element.HasElements ||
            element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "val") ||
            element.Attribute("val") is null)
            return false;
        switch ((string?)element.Attribute("val"))
        {
            case "1":
            case "true":
                overlay = true;
                break;
            case "0":
            case "false":
                break;
            default:
                return false;
        }
        hasOverlay = true;
        return true;
    }

    private static string LegendPositionToken(string position) => position switch
    {
        "" or "right" => "r",
        "top" => "t",
        "topRight" => "tr",
        "bottom" => "b",
        "left" => "l",
        _ => throw new CodecException("invalid_chart_legend", $"Unsupported chart legend position {position}."),
    };
    internal static bool UsesNumericXAxis(SpreadsheetChartType type) => type is SpreadsheetChartType.Scatter or SpreadsheetChartType.Bubble;

    private static bool TryReadPlotOptions(XElement plot, SpreadsheetChartArtifact chart)
    {
        if (chart.Type == SpreadsheetChartType.Bar)
        {
            if (!TryScalar(plot, "barDir", new[] { "col", "bar" }, required: true, out var direction) ||
                !TryScalar(plot, "grouping", new[] { "clustered", "stacked", "percentStacked" }, required: true, out var grouping) ||
                !TryOptionalUInt(plot, "gapWidth", 0, 500, out var hasGapWidth, out var gapWidth))
                return false;
            chart.BarDirection = direction == "bar" ? "bar" : "column";
            chart.Grouping = NativeGrouping(grouping!);
            if (hasGapWidth) chart.GapWidth = gapWidth;
            return true;
        }

        if (chart.Type is SpreadsheetChartType.Line or SpreadsheetChartType.Area)
        {
            if (!TryScalar(plot, "grouping", new[] { "standard", "stacked", "percentStacked" }, required: true, out var grouping) ||
                plot.Element(ChartNs + "gapWidth") is not null || plot.Element(ChartNs + "barDir") is not null)
                return false;
            chart.Grouping = NativeGrouping(grouping!);
            return true;
        }

        if (chart.Type == SpreadsheetChartType.Pie)
        {
            if (!TryOptionalUInt(plot, "firstSliceAng", 0, 360, out var hasAngle, out var angle) ||
                plot.Element(ChartNs + "holeSize") is not null)
                return false;
            if (hasAngle) chart.FirstSliceAngle = angle;
            return plot.Element(ChartNs + "grouping") is null &&
                   plot.Element(ChartNs + "gapWidth") is null &&
                   plot.Element(ChartNs + "barDir") is null;
        }

        if (chart.Type == SpreadsheetChartType.Doughnut)
        {
            if (!TryOptionalUInt(plot, "firstSliceAng", 0, 360, out var hasAngle, out var angle) ||
                !TryOptionalUInt(plot, "holeSize", 10, 90, out var hasHoleSize, out var holeSize) ||
                !hasHoleSize)
                return false;
            if (hasAngle) chart.FirstSliceAngle = angle;
            chart.DoughnutHoleSize = holeSize;
            return plot.Element(ChartNs + "grouping") is null &&
                   plot.Element(ChartNs + "gapWidth") is null &&
                   plot.Element(ChartNs + "barDir") is null;
        }

        if (chart.Type == SpreadsheetChartType.Bubble)
        {
            if (!TryOptionalUInt(plot, "bubbleScale", 0, 300, out var hasScale, out var scale) ||
                !TryScalar(plot, "sizeRepresents", new[] { "area", "w" }, required: false, out var sizeMode))
                return false;
            if (hasScale) chart.BubbleScale = scale;
            if (sizeMode is not null) chart.BubbleSizeMode = sizeMode == "w" ? "width" : "area";
        }

        return plot.Element(ChartNs + "grouping") is null &&
               plot.Element(ChartNs + "gapWidth") is null &&
               plot.Element(ChartNs + "barDir") is null &&
               plot.Element(ChartNs + "firstSliceAng") is null &&
               plot.Element(ChartNs + "holeSize") is null;
    }

    private static void PatchPlotOptions(XElement plot, SpreadsheetChartArtifact chart)
    {
        if (chart.Type == SpreadsheetChartType.Bar)
        {
            SetRequiredScalar(plot, "barDir", BarDirectionToken(chart.BarDirection));
            SetRequiredScalar(plot, "grouping", GroupingToken(chart.Grouping, clustered: true));
            PatchOptionalUInt(plot, "gapWidth", chart.HasGapWidth, chart.GapWidth);
            return;
        }
        if (chart.Type == SpreadsheetChartType.Area)
        {
            SetRequiredScalar(plot, "grouping", GroupingToken(chart.Grouping, clustered: false));
            return;
        }
        if (chart.Type == SpreadsheetChartType.Line)
        {
            SetRequiredScalar(plot, "grouping", GroupingToken(chart.Grouping, clustered: false));
            return;
        }
        if (chart.Type == SpreadsheetChartType.Pie)
        {
            PatchOptionalUInt(plot, "firstSliceAng", chart.HasFirstSliceAngle, chart.FirstSliceAngle);
            return;
        }
        if (chart.Type == SpreadsheetChartType.Doughnut)
        {
            PatchOptionalUInt(plot, "firstSliceAng", chart.HasFirstSliceAngle, chart.FirstSliceAngle);
            SetRequiredScalar(plot, "holeSize", (chart.HasDoughnutHoleSize ? chart.DoughnutHoleSize : 50U).ToString(CultureInfo.InvariantCulture));
            return;
        }
        if (chart.Type == SpreadsheetChartType.Bubble)
        {
            PatchOptionalUInt(plot, "bubbleScale", chart.HasBubbleScale, chart.BubbleScale);
            if (chart.BubbleSizeMode.Length == 0) plot.Element(ChartNs + "sizeRepresents")?.Remove();
            else SetRequiredScalar(plot, "sizeRepresents", BubbleSizeModeToken(chart.BubbleSizeMode));
            return;
        }
        if (chart.HasGapWidth || chart.Grouping.Length > 0 || chart.BarDirection.Length > 0 || chart.HasFirstSliceAngle || chart.HasDoughnutHoleSize || chart.HasBubbleScale || chart.BubbleSizeMode.Length > 0)
            throw new CodecException("invalid_chart_style", "The selected chart type cannot carry grouping, gap width, bar direction, or circular-plot geometry.");
    }

    private static SpreadsheetChartLineOptionsArtifact? LineOptions(SpreadsheetChartArtifact chart)
    {
        if (chart.Grouping.Length == 0) return chart.LineOptions;
        var output = chart.LineOptions?.Clone() ?? new SpreadsheetChartLineOptionsArtifact();
        output.Grouping = chart.Grouping switch
        {
            "none" => SpreadsheetChartLineGrouping.Standard,
            "stacked" => SpreadsheetChartLineGrouping.Stacked,
            "percent-stacked" => SpreadsheetChartLineGrouping.PercentStacked,
            _ => throw new CodecException("invalid_chart_grouping", $"Unsupported chart grouping {chart.Grouping}."),
        };
        return output;
    }

    private static string NativeGrouping(string value) => value switch
    {
        "clustered" or "standard" => "none",
        "stacked" => "stacked",
        "percentStacked" => "percent-stacked",
        _ => throw new InvalidOperationException("Validated native chart grouping changed unexpectedly."),
    };

    internal static string GroupingToken(string grouping, bool clustered) => grouping switch
    {
        "" or "none" => clustered ? "clustered" : "standard",
        "stacked" => "stacked",
        "percent-stacked" => "percentStacked",
        _ => throw new CodecException("invalid_chart_grouping", $"Unsupported chart grouping {grouping}."),
    };

    internal static string BarDirectionToken(string direction) => direction switch
    {
        "" or "column" => "col",
        "bar" => "bar",
        _ => throw new CodecException("invalid_chart_direction", $"Unsupported bar direction {direction}."),
    };

    private static bool TryScalar(XElement owner, string name, IReadOnlyCollection<string> allowed, bool required, out string? value)
    {
        value = null;
        var matches = owner.Elements(ChartNs + name).Take(2).ToArray();
        if (matches.Length == 0) return !required;
        if (matches.Length != 1) return false;
        var element = matches[0];
        value = (string?)element.Attribute("val");
        return value is not null && allowed.Contains(value) &&
               !element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "val") &&
               !element.Nodes().Any(node => node is XText text ? !string.IsNullOrWhiteSpace(text.Value) : true);
    }

    private static bool TryOptionalUInt(XElement owner, string name, uint minimum, uint maximum, out bool present, out uint value)
    {
        present = false;
        value = 0;
        var matches = owner.Elements(ChartNs + name).Take(2).ToArray();
        if (matches.Length == 0) return true;
        if (matches.Length != 1) return false;
        var element = matches[0];
        var text = (string?)element.Attribute("val");
        if (text is null ||
            element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "val") ||
            element.Nodes().Any(node => node is XText nativeText ? !string.IsNullOrWhiteSpace(nativeText.Value) : true) ||
            !uint.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out value) || value < minimum || value > maximum)
            return false;
        present = true;
        return true;
    }

    private static void SetRequiredScalar(XElement owner, string name, string value)
    {
        var existing = owner.Element(ChartNs + name);
        var replacement = new XElement(ChartNs + name, new XAttribute("val", value));
        if (existing is not null) existing.ReplaceWith(replacement);
        else owner.AddFirst(replacement);
    }

    private static void PatchOptionalUInt(XElement owner, string name, bool present, uint value)
    {
        var existing = owner.Element(ChartNs + name);
        if (!present)
        {
            existing?.Remove();
            return;
        }
        var replacement = new XElement(ChartNs + name, new XAttribute("val", value));
        if (existing is not null) existing.ReplaceWith(replacement);
        else
        {
            var following = owner.Elements().FirstOrDefault(element =>
                element.Name == ChartNs + "axId" || element.Name == ChartNs + "extLst");
            if (following is null) owner.Add(replacement);
            else following.AddBeforeSelf(replacement);
        }
    }

    private static bool PlotProfileEditable(XElement plot, SpreadsheetChartType type)
    {
        if (type == SpreadsheetChartType.Area) return true;
        if (type == SpreadsheetChartType.Doughnut) return true;
        if (type == SpreadsheetChartType.Scatter) return ScalarEquals(plot, "scatterStyle", "marker", required: true);
        if (type == SpreadsheetChartType.Bubble) return ScalarEquals(plot, "varyColors", "0", required: false) && ScalarEquals(plot, "bubble3D", "0", required: false) && ScalarEquals(plot, "showNegBubbles", "0", required: false);
        if (type == SpreadsheetChartType.Radar) return ScalarEquals(plot, "radarStyle", "standard", required: true) && ScalarEquals(plot, "varyColors", "0", required: false);
        return true;
    }

    private static bool ScalarEquals(XElement owner, string name, string expected, bool required)
    {
        var matches = owner.Elements(ChartNs + name).Take(2).ToArray();
        if (matches.Length == 0) return !required;
        if (matches.Length != 1) return false;
        var element = matches[0];
        return (string?)element.Attribute("val") == expected && !element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "val") && !element.Nodes().Any(node => node is XText text ? !string.IsNullOrWhiteSpace(text.Value) : true);
    }

    private static string BubbleSizeModeToken(string value) => value switch
    {
        "" or "area" => "area",
        "width" => "w",
        _ => throw new CodecException("invalid_chart_style", $"Unsupported bubble size mode {value}."),
    };

    internal static string DisplayBlanksAsToken(string value) => value switch
    {
        "zero" => "zero",
        "gap" => "gap",
        "span" => "span",
        _ => throw new CodecException("invalid_chart_style", $"Unsupported display blanks mode {value}."),
    };

    private static bool TryStringData(XElement source, out string[] values, out string formula)
    {
        formula = string.Empty; values = [];
        var literal = source.Element(ChartNs + "strLit");
        var reference = source.Element(ChartNs + "strRef");
        if ((literal is null) == (reference is null)) return false;
        if (reference is not null)
        {
            formula = reference.Element(ChartNs + "f")?.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(formula) || formula.Length > 8_192 || formula.StartsWith('=') || HasControls(formula)) return false;
            values = ReadStringPoints(reference.Element(ChartNs + "strCache")) ?? [];
        }
        else values = ReadStringPoints(literal) ?? [];
        return values.All(value => value.Length <= 32_767 && !HasControls(value));
    }

    private static bool TryNumericData(
        XElement source,
        bool allowMissing,
        out double[] values,
        out uint[] missingIndexes,
        out string formula)
    {
        formula = string.Empty; values = []; missingIndexes = [];
        var literal = source.Element(ChartNs + "numLit");
        var reference = source.Element(ChartNs + "numRef");
        if ((literal is null) == (reference is null)) return false;
        XElement? cache;
        if (reference is not null)
        {
            formula = reference.Element(ChartNs + "f")?.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(formula) || formula.Length > 8_192 || formula.StartsWith('=') || HasControls(formula)) return false;
            cache = reference.Element(ChartNs + "numCache");
        }
        else cache = literal;
        if (cache is null || !TryNumericPoints(cache, allowMissing, out var points, out var count, out missingIndexes)) return false;
        var output = new double[count];
        foreach (var point in points)
        {
            if (!double.TryParse(point.Element(ChartNs + "v")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number)) return false;
            output[checked((int)(uint)point.Attribute("idx")!)] = number;
        }
        values = output;
        return true;
    }

    private static bool TryNumericPoints(
        XElement source,
        bool allowMissing,
        out XElement[] points,
        out int count,
        out uint[] missingIndexes)
    {
        points = source.Elements(ChartNs + "pt").ToArray();
        count = 0;
        missingIndexes = [];
        if (points.Length > MaxPoints) return false;
        var declaredCount = (uint?)source.Element(ChartNs + "ptCount")?.Attribute("val");
        if (declaredCount is null)
        {
            count = points.Length;
            for (var index = 0; index < points.Length; index++)
                if ((uint?)points[index].Attribute("idx") != (uint)index) return false;
            return true;
        }
        if (declaredCount.Value > MaxPoints || (uint)points.Length > declaredCount.Value) return false;
        count = checked((int)declaredCount.Value);
        uint? previous = null;
        foreach (var point in points)
        {
            var index = (uint?)point.Attribute("idx");
            if (index is null || index.Value >= declaredCount.Value || previous is not null && index.Value <= previous.Value) return false;
            previous = index;
        }
        if (!allowMissing) return points.Length == count;
        var missing = new List<uint>(count - points.Length);
        var pointOffset = 0;
        for (uint index = 0; index < declaredCount.Value; index++)
        {
            if (pointOffset < points.Length && (uint)points[pointOffset].Attribute("idx")! == index) pointOffset++;
            else missing.Add(index);
        }
        missingIndexes = missing.ToArray();
        return true;
    }

    private static string[]? ReadStringPoints(XElement? source)
    {
        if (source is null || !TryOrderedPoints(source, out var points)) return null;
        return points.Select(item => item.Element(ChartNs + "v")?.Value ?? string.Empty).ToArray();
    }

    private static bool TryOrderedPoints(XElement source, out XElement[] points)
    {
        points = source.Elements(ChartNs + "pt").ToArray();
        if (points.Length > MaxPoints) return false;
        for (var index = 0; index < points.Length; index++) if ((uint?)points[index].Attribute("idx") != (uint)index) return false;
        var count = (uint?)source.Element(ChartNs + "ptCount")?.Attribute("val");
        return count is null || count.Value == points.Length;
    }

    private static XElement StringData(IEnumerable<string> values, string formula)
    {
        var cache = new XElement(ChartNs + (formula.Length > 0 ? "strCache" : "strLit"));
        AppendPoints(cache, values, value => value);
        return formula.Length > 0 ? new XElement(ChartNs + "strRef", new XElement(ChartNs + "f", formula), cache) : cache;
    }

    private static XElement NumericData(IEnumerable<double> values, IEnumerable<uint> missingIndexes, string formula)
    {
        var cache = new XElement(ChartNs + (formula.Length > 0 ? "numCache" : "numLit"), new XElement(ChartNs + "formatCode", "General"));
        AppendNumericPoints(cache, values, missingIndexes);
        return formula.Length > 0 ? new XElement(ChartNs + "numRef", new XElement(ChartNs + "f", formula), cache) : cache;
    }

    private static void AppendNumericPoints(XElement cache, IEnumerable<double> values, IEnumerable<uint> missingIndexes)
    {
        var array = values.ToArray();
        var missing = missingIndexes.ToHashSet();
        cache.Add(new XElement(ChartNs + "ptCount", new XAttribute("val", array.Length)));
        for (var index = 0; index < array.Length; index++)
        {
            if (missing.Contains((uint)index)) continue;
            cache.Add(new XElement(
                ChartNs + "pt",
                new XAttribute("idx", index),
                new XElement(ChartNs + "v", array[index].ToString("R", CultureInfo.InvariantCulture))));
        }
    }

    private static void AppendPoints<T>(XElement cache, IEnumerable<T> values, Func<T, string> format)
    {
        var array = values.ToArray();
        cache.Add(new XElement(ChartNs + "ptCount", new XAttribute("val", array.Length)));
        for (var index = 0; index < array.Length; index++) cache.Add(new XElement(ChartNs + "pt", new XAttribute("idx", index), new XElement(ChartNs + "v", format(array[index]))));
    }

    private static void PatchStringData(XElement? holder, IEnumerable<string> values, string formula, string errorCode, string subject)
    {
        if (holder is null) throw Topology(errorCode, subject, "category cache topology changed unexpectedly");
        var branch = holder.Element(ChartNs + (formula.Length > 0 ? "strRef" : "strLit")) ?? throw Topology(errorCode, subject, "category literal/reference topology changed unexpectedly");
        if (formula.Length > 0) (branch.Element(ChartNs + "f") ?? throw Topology(errorCode, subject, "category formula topology changed unexpectedly")).Value = formula;
        PatchPoints(formula.Length > 0 ? branch.Element(ChartNs + "strCache") : branch, values, value => value, errorCode, subject);
    }

    private static void PatchNumericData(
        XElement? holder,
        IEnumerable<double> values,
        IEnumerable<uint> missingIndexes,
        string formula,
        string errorCode,
        string subject)
    {
        if (holder is null) throw Topology(errorCode, subject, "numeric cache topology changed unexpectedly");
        var branch = holder.Element(ChartNs + (formula.Length > 0 ? "numRef" : "numLit")) ?? throw Topology(errorCode, subject, "numeric literal/reference topology changed unexpectedly");
        if (formula.Length > 0) (branch.Element(ChartNs + "f") ?? throw Topology(errorCode, subject, "numeric formula topology changed unexpectedly")).Value = formula;
        PatchNumericPoints(formula.Length > 0 ? branch.Element(ChartNs + "numCache") : branch, values, missingIndexes, errorCode, subject);
    }

    private static void PatchNumericPoints(
        XElement? cache,
        IEnumerable<double> values,
        IEnumerable<uint> missingIndexes,
        string errorCode,
        string subject)
    {
        if (cache is null) throw Topology(errorCode, subject, "numeric cache topology changed unexpectedly");
        var requested = values.ToArray();
        var requestedMissing = missingIndexes.ToArray();
        if (!TryNumericPoints(cache, allowMissing: true, out var points, out var count, out var existingMissing) ||
            count != requested.Length || !existingMissing.SequenceEqual(requestedMissing))
            throw Topology(errorCode, subject, "numeric point topology changed unexpectedly");
        cache.Element(ChartNs + "ptCount")?.SetAttributeValue("val", requested.Length);
        foreach (var point in points)
        {
            var index = checked((int)(uint)point.Attribute("idx")!);
            (point.Element(ChartNs + "v") ?? throw Topology(errorCode, subject, "point value topology changed unexpectedly")).Value =
                requested[index].ToString("R", CultureInfo.InvariantCulture);
        }
    }

    private static void PatchPoints<T>(XElement? cache, IEnumerable<T> values, Func<T, string> format, string errorCode, string subject)
    {
        if (cache is null) throw Topology(errorCode, subject, "cache topology changed unexpectedly");
        var requested = values.ToArray();
        var points = cache.Elements(ChartNs + "pt").ToArray();
        if (points.Length != requested.Length) throw Topology(errorCode, subject, "point topology changed unexpectedly");
        cache.Element(ChartNs + "ptCount")?.SetAttributeValue("val", requested.Length);
        for (var index = 0; index < points.Length; index++) (points[index].Element(ChartNs + "v") ?? throw Topology(errorCode, subject, "point value topology changed unexpectedly")).Value = format(requested[index]);
    }

    private static bool HasControls(string value) => value.Any(char.IsControl);
    private static CodecException Topology(string code, string subject, string message) => new(code, $"{subject} {message}.");
}
