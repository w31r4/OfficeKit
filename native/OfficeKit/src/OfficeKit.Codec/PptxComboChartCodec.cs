using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns one deliberate categorical mixed-plot profile: literal column, line and
// area families over a shared category domain, with one optional secondary
// axis pair. Keeping numeric scatter/bubble combinations outside this codec
// prevents incompatible category/value-axis semantics from being coerced.
internal static partial class PptxChartCodec
{
    private sealed record ComboNativeSeries(SpreadsheetChartType Type, PresentationChartAxisGroup AxisGroup, XElement Element, uint Order);

    private static XDocument BuildPresentationChartDocument(
        PresentationChart chart,
        string id,
        string name,
        PptxPartContext chartContext)
    {
        var document = chart.Type == SpreadsheetChartType.Combo
            ? BuildComboChartDocument(chart, id, name)
            : BuildChartDocument(ToSpreadsheet(chart, id, name));
        PptxChartFrameCodec.Patch(document.Root!, chart.Frame, chart.ChartAreaFill, $"Presentation chart {id} frame", chartContext);
        PptxChartTitleTextCodec.Apply(document, chart);
        return document;
    }

    private static void PatchPresentationChart(
        XDocument document,
        PresentationChart chart,
        string id,
        string name,
        PptxPartContext chartContext)
    {
        var rewritePlainTitle = chart.TitleBody is null && PptxChartTitleTextCodec.RequiresPlainRewrite(document);
        var patchTitle = chart.TitleBody is null && !rewritePlainTitle;
        if (chart.Type == SpreadsheetChartType.Combo) PatchComboChart(document, chart, id, name, patchTitle);
        else PatchChart(document, ToSpreadsheet(chart, id, name), patchTitle);
        PptxChartFrameCodec.Patch(document.Root!, chart.Frame, chart.ChartAreaFill, $"Presentation chart {id} frame", chartContext);
        if (rewritePlainTitle) PptxChartTitleTextCodec.ApplyPlain(document, chart);
        PptxChartTitleTextCodec.Apply(document, chart);
    }

    private static bool PresentationChartTopologyMatches(PresentationChart requested, PresentationChart original)
    {
        if (requested.Type != original.Type || requested.Categories.Count != original.Categories.Count) return false;
        if (requested.Type != SpreadsheetChartType.Combo)
        {
            if (requested.ComboSeries.Count != 0 || original.ComboSeries.Count != 0 || requested.Series.Count != original.Series.Count) return false;
            return requested.Series.Zip(original.Series).All(pair =>
                pair.First.Values.Count == pair.Second.Values.Count &&
                pair.First.MissingValueIndexes.SequenceEqual(pair.Second.MissingValueIndexes) &&
                pair.First.Trendlines.Count == pair.Second.Trendlines.Count &&
                (pair.First.ErrorBars is null) == (pair.Second.ErrorBars is null));
        }
        if (requested.Series.Count != 0 || original.Series.Count != 0 || requested.ComboSeries.Count != original.ComboSeries.Count) return false;
        if ((requested.SecondaryXAxis is null) != (original.SecondaryXAxis is null) || (requested.SecondaryYAxis is null) != (original.SecondaryYAxis is null)) return false;
        return requested.ComboSeries.Zip(original.ComboSeries).All(pair =>
            pair.First.Type == pair.Second.Type &&
            ComboAxisGroup(pair.First) == ComboAxisGroup(pair.Second) &&
            pair.First.Series is not null && pair.Second.Series is not null &&
            pair.First.Series.Values.Count == pair.Second.Series.Values.Count &&
            pair.First.Series.MissingValueIndexes.SequenceEqual(pair.Second.Series.MissingValueIndexes) &&
            pair.First.Series.Trendlines.Count == pair.Second.Series.Trendlines.Count &&
            (pair.First.Series.ErrorBars is null) == (pair.Second.Series.ErrorBars is null));
    }

    private static void ValidateComboChart(PresentationChart chart, string elementId, string name, bool allowFormulas = false)
    {
        if (chart.Series.Count != 0) throw Invalid(elementId, "must keep series empty when type is combo");
        if (chart.ComboSeries.Count is < 2 or > MaxSeries) throw Invalid(elementId, $"must contain 2 through {MaxSeries} combo_series entries");
        if (chart.Categories.Count > MaxPoints || chart.Categories.Any(value => value.Length > 32_767 || HasControls(value))) throw Invalid(elementId, "contains invalid categories");

        var families = new Dictionary<SpreadsheetChartType, List<PresentationComboSeriesArtifact>>();
        foreach (var entry in chart.ComboSeries)
        {
            if (entry.Series is null) throw Invalid(elementId, "contains a combo series without payload");
            if (entry.Type is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line or SpreadsheetChartType.Area))
                throw Invalid(elementId, "combo series type must be column, line, or area");
            if (!families.TryGetValue(entry.Type, out var family)) families[entry.Type] = family = [];
            family.Add(entry);
            var hasFormulas = !string.IsNullOrWhiteSpace(entry.Series.CategoryFormula) ||
                !string.IsNullOrWhiteSpace(entry.Series.ValueFormula) ||
                !string.IsNullOrWhiteSpace(entry.Series.XValueFormula) ||
                !string.IsNullOrWhiteSpace(entry.Series.BubbleSizeFormula) ||
                ErrorBarsUseFormula(entry.Series);
            if (hasFormulas && !allowFormulas)
                throw Invalid(elementId, "must use literal categories and values without workbook formulas");
            if (hasFormulas &&
                (!FormulaProfileIsSafe(entry.Series.CategoryFormula) ||
                 !FormulaProfileIsSafe(entry.Series.ValueFormula) ||
                 !FormulaProfileIsSafe(entry.Series.XValueFormula) ||
                 !FormulaProfileIsSafe(entry.Series.BubbleSizeFormula) ||
                 !FormulaProfileIsSafe(entry.Series.ErrorBars?.Plus?.Formula ?? string.Empty) ||
                 !FormulaProfileIsSafe(entry.Series.ErrorBars?.Minus?.Formula ?? string.Empty)))
                throw Invalid(elementId, "contains a formula outside the local worksheet range profile");
        }
        if (families.Count < 2) throw Invalid(elementId, "must contain at least two distinct column, line, or area plot families");
        if (families.ContainsKey(SpreadsheetChartType.Bar) && chart.BarDirection == "bar")
            throw Invalid(elementId, "horizontal bars cannot share the categorical combo-axis profile; use columns");
        if (chart.HasGapWidth && !families.ContainsKey(SpreadsheetChartType.Bar))
            throw Invalid(elementId, "gap_width requires a column plot family");
        if (chart.HasOverlap && !families.ContainsKey(SpreadsheetChartType.Bar))
            throw Invalid(elementId, "overlap requires a column plot family");
        if (chart.HasVaryColors && !families.ContainsKey(SpreadsheetChartType.Bar))
            throw Invalid(elementId, "vary_colors requires a column plot family");
        foreach (var (type, family) in families)
            if (family.Select(ComboAxisGroup).Distinct().Count() != 1)
                throw Invalid(elementId, $"cannot split one {type.ToString().ToLowerInvariant()} plot family across primary and secondary axes");

        var hasPrimaryPlot = chart.ComboSeries.Any(entry => ComboAxisGroup(entry) == PresentationChartAxisGroup.Primary);
        var hasSecondaryPlot = HasSecondaryComboPlot(chart);
        if (!hasPrimaryPlot) throw Invalid(elementId, "requires at least one primary-axis plot family");
        if (!hasSecondaryPlot && (chart.SecondaryXAxis is not null || chart.SecondaryYAxis is not null))
            throw Invalid(elementId, "cannot carry secondary axes without a secondary plot family");
        if ((chart.SecondaryXAxis is null) != (chart.SecondaryYAxis is null))
            throw Invalid(elementId, "must carry both secondary axes or neither");

        foreach (var (type, family) in families)
        {
            var series = family.Select(entry => entry.Series).ToArray();
            var probe = ComboSpreadsheetChart(chart, elementId, name, type, series);
            if (ComboAxisGroup(family[0]) == PresentationChartAxisGroup.Secondary)
            {
                probe.XAxis = chart.SecondaryXAxis?.Clone();
                probe.YAxis = chart.SecondaryYAxis?.Clone();
            }
            try { XlsxChartCodec.Validate([probe], $"presentation/{elementId}"); }
            catch (CodecException error) when (error.Code == "invalid_spreadsheet_chart") { throw Invalid(elementId, error.Message); }
        }
    }

    private static SpreadsheetChartArtifact ComboSpreadsheetChart(PresentationChart source, string id, string name, SpreadsheetChartType type, IEnumerable<SpreadsheetChartSeriesArtifact> series)
    {
        var output = new SpreadsheetChartArtifact
        {
            Id = id,
            Name = name,
            Title = source.Title,
            Type = type,
            HasLegend = source.HasLegend,
            LegendPosition = source.LegendPosition,
            Grouping = source.Grouping,
            TitleTextStyle = source.TitleTextStyle?.Clone(),
            AbsoluteAnchor = new SpreadsheetAbsoluteAnchorArtifact
            {
                XEmu = source.LeftEmu,
                YEmu = source.TopEmu,
                WidthEmu = source.WidthEmu,
                HeightEmu = source.HeightEmu,
            },
        };
        if (source.LegendTextStyle is not null) output.LegendTextStyle = source.LegendTextStyle.Clone();
        if (source.LegendFill is not null) output.LegendFill = source.LegendFill.Clone();
        if (source.LegendLine is not null) output.LegendLine = source.LegendLine.Clone();
        output.Categories.Add(source.Categories);
        output.Series.Add(series.Select(item => item.Clone()));
        if (source.XAxis is not null) output.XAxis = source.XAxis.Clone();
        if (source.YAxis is not null) output.YAxis = source.YAxis.Clone();
        if (source.HasShowCategoryAxis)
        {
            output.XAxis ??= new SpreadsheetChartAxisArtifact();
            output.XAxis.Visible = source.ShowCategoryAxis;
        }
        if (source.HasShowValueAxis)
        {
            output.YAxis ??= new SpreadsheetChartAxisArtifact();
            output.YAxis.Visible = source.ShowValueAxis;
        }
        if (source.HasShowGridlines)
        {
            output.YAxis ??= new SpreadsheetChartAxisArtifact();
            output.YAxis.ShowMajorGridlines = source.ShowGridlines;
        }
        if (type == SpreadsheetChartType.Bar)
        {
            output.BarDirection = source.BarDirection;
            if (source.HasGapWidth) output.GapWidth = source.GapWidth;
            if (source.HasOverlap) output.Overlap = source.Overlap;
            if (source.HasVaryColors) output.VaryColors = source.VaryColors;
        }
        if (source.ChartAreaFill is not null) output.ChartAreaFill = source.ChartAreaFill.Clone();
        if (source.PlotAreaFill is not null) output.PlotAreaFill = source.PlotAreaFill.Clone();
        if (source.DataLabels is not null) output.DataLabels = source.DataLabels.Clone();
        if (source.HasDisplayBlanksAs) output.DisplayBlanksAs = source.DisplayBlanksAs;
        if (source.HasLegendOverlay) output.LegendOverlay = source.LegendOverlay;
        return output;
    }

    private static SpreadsheetChartArtifact ComboAxisCarrier(PresentationChart source, string id, string name, bool secondary = false)
    {
        var output = ComboSpreadsheetChart(source, id, name, SpreadsheetChartType.Bar, []);
        if (secondary)
        {
            output.XAxis = source.SecondaryXAxis?.Clone();
            output.YAxis = source.SecondaryYAxis?.Clone();
        }
        return output;
    }

    private static PresentationChartAxisGroup ComboAxisGroup(PresentationComboSeriesArtifact entry) =>
        entry.AxisGroup == PresentationChartAxisGroup.Secondary
            ? PresentationChartAxisGroup.Secondary
            : PresentationChartAxisGroup.Primary;

    private static bool HasSecondaryComboPlot(PresentationChart chart) =>
        chart.ComboSeries.Any(entry => ComboAxisGroup(entry) == PresentationChartAxisGroup.Secondary);

    private static bool TryReadComboChart(
        string xml,
        out PresentationChart chart,
        out XDocument document,
        out bool editable,
        bool allowChartFrameDecorations = false)
    {
        chart = new PresentationChart();
        editable = true;
        try { document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace); }
        catch (XmlException) { document = new XDocument(); return false; }
        var root = document.Root;
        var nativeChart = root?.Element(ChartNs + "chart");
        var plotArea = nativeChart?.Element(ChartNs + "plotArea");
        if (root?.Name != ChartNs + "chartSpace" || nativeChart is null || plotArea is null || root.Element(ChartNs + "externalData") is not null) return false;
        var plots = plotArea.Elements().Where(item => item.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToArray();
        if (plots.Length is < 2 or > 3 ||
            plots.Any(plot => !TryComboPlotType(plot, out _)) ||
            plots.Select(plot => { TryComboPlotType(plot, out var type); return type; }).Distinct().Count() != plots.Length)
            return false;
        XElement? primaryPlot = null;
        XElement? secondaryPlot = null;
        var nativePlots = new List<(XElement Plot, SpreadsheetChartType Type, PresentationChartAxisGroup AxisGroup)>();
        string? commonGrouping = null;
        foreach (var plot in plots)
        {
            if (!TryComboPlotType(plot, out var type) ||
                !TryReadComboPlotStyle(plot, type, out var grouping, out var barDirection, out var hasGapWidth, out var gapWidth, out var hasOverlap, out var overlap, out var hasVaryColors, out var varyColors)) return false;
            commonGrouping ??= grouping;
            if (!string.Equals(commonGrouping, grouping, StringComparison.Ordinal)) return false;
            if (!TryComboAxisGroup(plotArea, plot, out var axisGroup)) return false;
            if (axisGroup == PresentationChartAxisGroup.Primary)
            {
                if (primaryPlot is not null && !SharesComboAxes(primaryPlot, plot)) return false;
                primaryPlot ??= plot;
            }
            else
            {
                if (secondaryPlot is not null && !SharesComboAxes(secondaryPlot, plot)) return false;
                secondaryPlot ??= plot;
            }
            if (type == SpreadsheetChartType.Bar)
            {
                if (barDirection != "column") return false;
                chart.BarDirection = barDirection;
                if (hasGapWidth) chart.GapWidth = gapWidth;
                if (hasOverlap) chart.Overlap = overlap;
                if (hasVaryColors) chart.VaryColors = varyColors;
            }
            nativePlots.Add((plot, type, axisGroup));
        }
        if (primaryPlot is null) return false;

        var title = nativeChart.Element(ChartNs + "title");
        if (title is not null)
        {
            var richText = title.Descendants(DrawingNs + "t").ToArray();
            var directValue = title.Descendants(ChartNs + "v").FirstOrDefault();
            chart.Title = richText.Length > 0 ? string.Concat(richText.Select(item => item.Value)) : directValue?.Value ?? string.Empty;
            if (richText.Length == 0) editable = false;
            var titleProbe = new SpreadsheetChartArtifact();
            _ = XlsxChartTextStyleCodec.TryReadTitle(title, titleProbe);
            if (titleProbe.TitleTextStyle is not null) chart.TitleTextStyle = titleProbe.TitleTextStyle.Clone();
        }
        chart.Type = SpreadsheetChartType.Combo;
        chart.Grouping = commonGrouping ?? string.Empty;
        if (!OpenXmlChartSpaceCodec.TryReadDisplayBlanksAs(nativeChart, out var displayBlanksAs)) editable = false;
        else if (displayBlanksAs is not null) chart.DisplayBlanksAs = displayBlanksAs;
        var legend = nativeChart.Element(ChartNs + "legend");
        chart.HasLegend = legend is not null;
        if (legend is not null)
        {
            var legendProbe = new SpreadsheetChartArtifact();
            if (!OpenXmlChartSpaceCodec.TryReadLegend(legend, legendProbe)) editable = false;
            else
            {
                chart.LegendPosition = legendProbe.LegendPosition;
                if (legendProbe.HasLegendOverlay) chart.LegendOverlay = legendProbe.LegendOverlay;
                if (legendProbe.LegendTextStyle is not null) chart.LegendTextStyle = legendProbe.LegendTextStyle.Clone();
                if (legendProbe.LegendFill is not null) chart.LegendFill = legendProbe.LegendFill.Clone();
                if (legendProbe.LegendLine is not null) chart.LegendLine = legendProbe.LegendLine.Clone();
            }
        }

        var collectedSeries = new List<ComboNativeSeries>();
        foreach (var nativePlot in nativePlots)
        {
            if (!TryReadComboSeries(nativePlot.Plot, nativePlot.Type, nativePlot.AxisGroup, out var plotSeries)) return false;
            collectedSeries.AddRange(plotSeries);
        }
        var orderedSeries = collectedSeries.OrderBy(item => item.Order).ToArray();
        if (orderedSeries.Length is < 2 or > MaxSeries || orderedSeries.Select(item => item.Order).Distinct().Count() != orderedSeries.Length ||
            !orderedSeries.Select(item => item.Order).SequenceEqual(Enumerable.Range(0, orderedSeries.Length).Select(index => (uint)index))) return false;

        string[]? commonCategories = null;
        foreach (var native in orderedSeries)
        {
            if (!OpenXmlChartSpaceCodec.TrySeries(native.Element, native.Type, out var series, out var categories, out var seriesEditable) ||
                !FormulaProfileIsSafe(series.CategoryFormula) || !FormulaProfileIsSafe(series.ValueFormula) ||
                !FormulaProfileIsSafe(series.XValueFormula) || !FormulaProfileIsSafe(series.BubbleSizeFormula)) return false;
            editable &= seriesEditable;
            if (commonCategories is null) commonCategories = categories;
            else if (!commonCategories.SequenceEqual(categories, StringComparer.Ordinal)) return false;
            chart.ComboSeries.Add(new PresentationComboSeriesArtifact { Type = native.Type, AxisGroup = native.AxisGroup, Series = series });
        }
        chart.Categories.Add(commonCategories ?? []);
        if (chart.Categories.Count > MaxPoints) return false;

        string? labelSemantics = null;
        foreach (var plot in plots)
        {
            var labelProbe = new SpreadsheetChartArtifact();
            if (!XlsxChartDataLabelsCodec.TryRead(plot, labelProbe)) return false;
            var semantics = XlsxChartDataLabelsCodec.Semantics(labelProbe.DataLabels);
            labelSemantics ??= semantics;
            if (!string.Equals(labelSemantics, semantics, StringComparison.Ordinal)) return false;
            if (chart.DataLabels is null && labelProbe.DataLabels is not null) chart.DataLabels = labelProbe.DataLabels;
        }

        var axisCarrier = ComboAxisCarrier(chart, "combo", "combo");
        if (!XlsxChartAxisCodec.TryRead(plotArea, primaryPlot, axisCarrier, out var axesEditable)) return false;
        chart.XAxis = axisCarrier.XAxis;
        chart.YAxis = axisCarrier.YAxis;
        if (chart.XAxis?.HasVisible == true) chart.ShowCategoryAxis = chart.XAxis.Visible;
        if (chart.YAxis?.HasVisible == true) chart.ShowValueAxis = chart.YAxis.Visible;
        if (chart.YAxis?.HasShowMajorGridlines == true) chart.ShowGridlines = chart.YAxis.ShowMajorGridlines;
        editable &= axesEditable;
        if (secondaryPlot is not null)
        {
            var secondaryAxisCarrier = ComboAxisCarrier(chart, "combo", "combo", secondary: true);
            if (!XlsxChartAxisCodec.TryReadPresentationSecondary(plotArea, secondaryPlot, secondaryAxisCarrier, out var secondaryAxesEditable)) return false;
            chart.SecondaryXAxis = secondaryAxisCarrier.XAxis;
            chart.SecondaryYAxis = secondaryAxisCarrier.YAxis;
            editable &= secondaryAxesEditable;
        }
        var chartSpaceProperties = root.Element(ChartNs + "spPr");
        if (!XlsxChartSurfaceFillCodec.TryRead(chartSpaceProperties, out var chartAreaFill, allowChartFrameDecorations))
        {
            // The Presentation-only frame codec resolves a ChartPart image
            // relationship after this package-agnostic combo parser returns.
            if (!(allowChartFrameDecorations && chartSpaceProperties?.Elements(DrawingNs + "blipFill").Any() == true))
                editable = false;
        }
        else if (chartAreaFill is not null) chart.ChartAreaFill = chartAreaFill;
        if (!XlsxChartSurfaceFillCodec.TryRead(plotArea.Element(ChartNs + "spPr"), out var plotAreaFill)) editable = false;
        else if (plotAreaFill is not null) chart.PlotAreaFill = plotAreaFill;
        return chart.Title.Length <= 32_767 && !HasControls(chart.Title);
    }

    private static bool TryReadComboSeries(XElement plot, SpreadsheetChartType type, PresentationChartAxisGroup axisGroup, out ComboNativeSeries[] result)
    {
        result = [];
        var nativeSeries = plot.Elements(ChartNs + "ser").ToArray();
        if (nativeSeries.Length == 0) return false;
        var output = new List<ComboNativeSeries>();
        foreach (var series in nativeSeries)
        {
            if (!TryComboSeriesOrder(series, "idx", out var index) || !TryComboSeriesOrder(series, "order", out var order) || index != order) return false;
            output.Add(new ComboNativeSeries(type, axisGroup, series, order));
        }
        result = output.ToArray();
        return true;
    }

    private static bool TryComboSeriesOrder(XElement series, string name, out uint value)
    {
        value = 0;
        var elements = series.Elements(ChartNs + name).Take(2).ToArray();
        if (elements.Length != 1 || elements[0].Elements().Any() || elements[0].Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value))) return false;
        var attributes = elements[0].Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        return attributes.Length == 1 && attributes[0].Name == "val" && uint.TryParse(attributes[0].Value, NumberStyles.None, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryReadComboPlotStyle(
        XElement plot,
        SpreadsheetChartType type,
        out string grouping,
        out string barDirection,
        out bool hasGapWidth,
        out uint gapWidth,
        out bool hasOverlap,
        out int overlap,
        out bool hasVaryColors,
        out bool varyColors)
    {
        grouping = string.Empty;
        barDirection = string.Empty;
        hasGapWidth = false;
        gapWidth = 0;
        hasOverlap = false;
        overlap = 0;
        hasVaryColors = false;
        varyColors = false;
        var allowed = type == SpreadsheetChartType.Bar
            ? new HashSet<XName> { ChartNs + "barDir", ChartNs + "grouping", ChartNs + "varyColors", ChartNs + "ser", ChartNs + "dLbls", ChartNs + "gapWidth", ChartNs + "overlap", ChartNs + "axId" }
            : new HashSet<XName> { ChartNs + "grouping", ChartNs + "ser", ChartNs + "dLbls", ChartNs + "axId" };
        if (plot.Elements().Any(item => !allowed.Contains(item.Name))) return false;
        if (!ComboScalar(plot.Element(ChartNs + "grouping"), out var nativeGrouping) ||
            nativeGrouping is not ("clustered" or "standard" or "stacked" or "percentStacked")) return false;
        if ((type == SpreadsheetChartType.Bar && nativeGrouping == "standard") ||
            (type is SpreadsheetChartType.Line or SpreadsheetChartType.Area && nativeGrouping == "clustered")) return false;
        grouping = nativeGrouping switch
        {
            "clustered" or "standard" => "none",
            "stacked" => "stacked",
            "percentStacked" => "percent-stacked",
            _ => string.Empty,
        };
        if (type == SpreadsheetChartType.Bar)
        {
            if (!ComboScalar(plot.Element(ChartNs + "barDir"), out var nativeDirection) || nativeDirection is not ("col" or "bar")) return false;
            barDirection = nativeDirection == "bar" ? "bar" : "column";
            var vary = plot.Elements(ChartNs + "varyColors").Take(2).ToArray();
            if (vary.Length > 1 || vary.Length == 1 &&
                (!ComboScalar(vary[0], out var varyText) || varyText is not ("0" or "1" or "false" or "true"))) return false;
            hasVaryColors = vary.Length == 1;
            if (hasVaryColors) varyColors = vary[0].Attribute("val")!.Value is "1" or "true";
            var gap = plot.Elements(ChartNs + "gapWidth").Take(2).ToArray();
            if (gap.Length > 1 || gap.Length == 1 &&
                (!ComboScalar(gap[0], out var text) || !uint.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out gapWidth) || gapWidth > 500)) return false;
            hasGapWidth = gap.Length == 1;
            var overlapElements = plot.Elements(ChartNs + "overlap").Take(2).ToArray();
            if (overlapElements.Length > 1 || overlapElements.Length == 1 &&
                (!ComboScalar(overlapElements[0], out var overlapText) || !int.TryParse(overlapText, NumberStyles.Integer, CultureInfo.InvariantCulture, out overlap) || overlap is < -100 or > 100)) return false;
            hasOverlap = overlapElements.Length == 1;
        }
        var axisIds = plot.Elements(ChartNs + "axId").ToArray();
        return axisIds.Length == 2 && axisIds.All(item => ComboScalar(item, out _));
    }

    private static bool TryComboPlotType(XElement plot, out SpreadsheetChartType type)
    {
        type = plot.Name.LocalName switch
        {
            "barChart" => SpreadsheetChartType.Bar,
            "lineChart" => SpreadsheetChartType.Line,
            "areaChart" => SpreadsheetChartType.Area,
            _ => SpreadsheetChartType.Unspecified,
        };
        return type != SpreadsheetChartType.Unspecified;
    }

    private static XName ComboPlotName(SpreadsheetChartType type) => type switch
    {
        SpreadsheetChartType.Bar => ChartNs + "barChart",
        SpreadsheetChartType.Line => ChartNs + "lineChart",
        SpreadsheetChartType.Area => ChartNs + "areaChart",
        _ => throw new ArgumentOutOfRangeException(nameof(type)),
    };

    private static int ComboPlotOrder(SpreadsheetChartType type) => type switch
    {
        SpreadsheetChartType.Area => 0,
        SpreadsheetChartType.Bar => 1,
        SpreadsheetChartType.Line => 2,
        _ => 3,
    };

    private static bool SharesComboAxes(XElement firstPlot, XElement secondPlot)
    {
        var firstIds = firstPlot.Elements(ChartNs + "axId").Select(item => ComboScalar(item, out var value) ? value : string.Empty).ToArray();
        var secondIds = secondPlot.Elements(ChartNs + "axId").Select(item => ComboScalar(item, out var value) ? value : string.Empty).ToArray();
        return firstIds.Length == 2 && secondIds.Length == 2 && firstIds.Distinct(StringComparer.Ordinal).Count() == 2 && firstIds.SequenceEqual(secondIds, StringComparer.Ordinal);
    }

    private static bool TryComboAxisGroup(XElement plotArea, XElement plot, out PresentationChartAxisGroup axisGroup)
    {
        axisGroup = PresentationChartAxisGroup.Primary;
        var ids = plot.Elements(ChartNs + "axId")
            .Select(item => ComboScalar(item, out var value) ? value : string.Empty)
            .ToHashSet(StringComparer.Ordinal);
        if (ids.Count != 2 || ids.Contains(string.Empty)) return false;
        var categoryAxes = plotArea.Elements(ChartNs + "catAx")
            .Where(axis => ComboScalar(axis.Element(ChartNs + "axId"), out var id) && ids.Contains(id))
            .ToArray();
        var valueAxes = plotArea.Elements(ChartNs + "valAx")
            .Where(axis => ComboScalar(axis.Element(ChartNs + "axId"), out var id) && ids.Contains(id))
            .ToArray();
        if (categoryAxes.Length != 1 || valueAxes.Length != 1 ||
            !ComboScalar(categoryAxes[0].Element(ChartNs + "axPos"), out var categoryPosition) ||
            !ComboScalar(valueAxes[0].Element(ChartNs + "axPos"), out var valuePosition)) return false;
        if (categoryPosition == "b" && valuePosition == "l") return true;
        if (categoryPosition == "t" && valuePosition == "r")
        {
            axisGroup = PresentationChartAxisGroup.Secondary;
            return true;
        }
        return false;
    }

    private static bool ComboScalar(XElement? element, out string value)
    {
        value = string.Empty;
        if (element is null || element.Elements().Any() || element.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value))) return false;
        var attributes = element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        if (attributes.Length != 1 || attributes[0].Name != "val") return false;
        value = attributes[0].Value;
        return value.Length > 0;
    }

    private static XDocument BuildComboChartDocument(PresentationChart chart, string id, string name)
    {
        var indexed = chart.ComboSeries.Select((item, index) => new ComboNativeSeries(item.Type, ComboAxisGroup(item), OpenXmlChartSpaceCodec.SeriesElement(item.Series, chart.Categories, index, item.Type), checked((uint)index))).ToArray();
        var plots = indexed
            .GroupBy(item => item.Type)
            .OrderBy(group => ComboPlotOrder(group.Key))
            .Select(group => BuildComboPlot(chart, group.Key, group.First().AxisGroup, group.Select(item => item.Element)))
            .ToArray();
        var plotArea = new XElement(ChartNs + "plotArea", new XElement(ChartNs + "layout"), plots);
        XlsxChartAxisCodec.AppendAuthored(plotArea, ComboAxisCarrier(chart, id, name));
        if (HasSecondaryComboPlot(chart)) XlsxChartAxisCodec.AppendAuthoredPresentationSecondary(plotArea, ComboAxisCarrier(chart, id, name, secondary: true));
        if (XlsxChartSurfaceFillCodec.Element(chart.PlotAreaFill, "Presentation combo chart plot area") is { } plotFill) plotArea.Add(plotFill);
        var nativeChart = new XElement(ChartNs + "chart");
        if (chart.Title.Length > 0) nativeChart.Add(XlsxChartTextStyleCodec.TitleElement(chart.Title, chart.TitleTextStyle));
        nativeChart.Add(plotArea);
        if (chart.HasLegend) nativeChart.Add(OpenXmlChartSpaceCodec.LegendElement(chart.LegendPosition, chart.LegendTextStyle, chart.HasLegendOverlay, chart.LegendOverlay, chart.LegendFill, chart.LegendLine));
        nativeChart.Add(new XElement(ChartNs + "plotVisOnly", new XAttribute("val", "1")));
        if (chart.HasDisplayBlanksAs)
            nativeChart.Add(new XElement(ChartNs + "dispBlanksAs", new XAttribute("val", OpenXmlChartSpaceCodec.DisplayBlanksAsToken(chart.DisplayBlanksAs))));
        var chartSpace = new XElement(ChartNs + "chartSpace", new XAttribute(XNamespace.Xmlns + "c", ChartNs), new XAttribute(XNamespace.Xmlns + "a", DrawingNs), nativeChart);
        if (XlsxChartSurfaceFillCodec.Element(chart.ChartAreaFill, "Presentation combo chart area") is { } chartFill) chartSpace.Add(chartFill);
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), chartSpace);
    }

    private static XElement BuildComboPlot(
        PresentationChart chart,
        SpreadsheetChartType type,
        PresentationChartAxisGroup axisGroup,
        IEnumerable<XElement> series)
    {
        var plot = new XElement(ComboPlotName(type));
        if (type == SpreadsheetChartType.Bar)
            plot.Add(new XElement(ChartNs + "barDir", new XAttribute("val", OpenXmlChartSpaceCodec.BarDirectionToken(chart.BarDirection))));
        plot.Add(new XElement(ChartNs + "grouping", new XAttribute("val", OpenXmlChartSpaceCodec.GroupingToken(chart.Grouping, clustered: type == SpreadsheetChartType.Bar))));
        if (type == SpreadsheetChartType.Bar && chart.HasVaryColors)
            plot.Add(new XElement(ChartNs + "varyColors", new XAttribute("val", chart.VaryColors ? "1" : "0")));
        plot.Add(series);
        plot.Add(XlsxChartDataLabelsCodec.Element(chart.DataLabels));
        if (type == SpreadsheetChartType.Bar && chart.HasGapWidth)
            plot.Add(new XElement(ChartNs + "gapWidth", new XAttribute("val", chart.GapWidth)));
        if (type == SpreadsheetChartType.Bar && chart.HasOverlap)
            plot.Add(new XElement(ChartNs + "overlap", new XAttribute("val", chart.Overlap)));
        var secondary = axisGroup == PresentationChartAxisGroup.Secondary;
        plot.Add(new XElement(ChartNs + "axId", new XAttribute("val", secondary ? "3" : "1")));
        plot.Add(new XElement(ChartNs + "axId", new XAttribute("val", secondary ? "4" : "2")));
        return plot;
    }

    private static void PatchComboChart(XDocument document, PresentationChart target, string id, string name, bool patchTitle)
    {
        var nativeChart = document.Root!.Element(ChartNs + "chart")!;
        if (patchTitle)
            OpenXmlChartSpaceCodec.PatchTitle(nativeChart, target.Title, target.TitleTextStyle, "unsupported_presentation_edit", "Presentation combo chart");
        OpenXmlChartSpaceCodec.PatchLegend(nativeChart, target.HasLegend, target.LegendPosition, target.LegendTextStyle, target.HasLegendOverlay, target.LegendOverlay, target.LegendFill, target.LegendLine);
        OpenXmlChartSpaceCodec.PatchDisplayBlanksAs(nativeChart, target.HasDisplayBlanksAs, target.DisplayBlanksAs, "unsupported_presentation_edit", "Presentation combo chart");
        var plotArea = nativeChart.Element(ChartNs + "plotArea")!;
        var plots = plotArea.Elements().Where(item => item.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToArray();
        if (plots.Length is < 2 or > 3) throw new CodecException("unsupported_presentation_edit", "Presentation combo chart no longer has two or three bounded plots.");
        XElement? primaryPlot = null;
        XElement? secondaryPlot = null;
        var collectedSeries = new List<ComboNativeSeries>();
        foreach (var plot in plots)
        {
            if (!TryComboPlotType(plot, out var type) || !TryComboAxisGroup(plotArea, plot, out var axisGroup) ||
                !TryReadComboSeries(plot, type, axisGroup, out var plotSeries))
                throw new CodecException("unsupported_presentation_edit", "Presentation combo chart no longer matches the bounded native plot profile.");
            if (axisGroup == PresentationChartAxisGroup.Primary) primaryPlot ??= plot;
            else secondaryPlot ??= plot;
            if (type == SpreadsheetChartType.Bar)
            {
                PatchComboScalar(plot, "barDir", OpenXmlChartSpaceCodec.BarDirectionToken(target.BarDirection));
                PatchComboVaryColors(plot, target);
                PatchComboGapWidth(plot, target);
                PatchComboOverlap(plot, target);
            }
            PatchComboScalar(plot, "grouping", OpenXmlChartSpaceCodec.GroupingToken(target.Grouping, clustered: type == SpreadsheetChartType.Bar));
            collectedSeries.AddRange(plotSeries);
        }
        if (primaryPlot is null || (secondaryPlot is not null) != HasSecondaryComboPlot(target))
            throw new CodecException("unsupported_presentation_edit", "Presentation combo chart axis topology no longer matches the bounded profile.");
        var nativeSeries = collectedSeries.OrderBy(item => item.Order).ToArray();
        if (nativeSeries.Length != target.ComboSeries.Count || nativeSeries.Select(item => item.Order).Distinct().Count() != nativeSeries.Length ||
            !nativeSeries.Select(item => item.Order).SequenceEqual(Enumerable.Range(0, nativeSeries.Length).Select(index => (uint)index)))
            throw new CodecException("presentation_chart_topology_changed", "Presentation combo chart series topology changed unexpectedly.");
        for (var index = 0; index < nativeSeries.Length; index++)
        {
            var requested = target.ComboSeries[index];
            if (requested.Type != nativeSeries[index].Type || ComboAxisGroup(requested) != nativeSeries[index].AxisGroup || requested.Series is null) throw new CodecException("presentation_chart_topology_changed", "Presentation combo chart series type or axis group changed unexpectedly.");
            OpenXmlChartSpaceCodec.PatchSeries(nativeSeries[index].Element, requested.Series, target.Categories, requested.Type, "presentation_chart_topology_changed", "Presentation combo chart");
        }
        foreach (var plot in plots) XlsxChartDataLabelsCodec.Patch(plot, target.DataLabels);
        XlsxChartAxisCodec.Patch(plotArea, primaryPlot, ComboAxisCarrier(target, id, name));
        if (secondaryPlot is not null) XlsxChartAxisCodec.PatchPresentationSecondary(plotArea, secondaryPlot, ComboAxisCarrier(target, id, name, secondary: true));
        XlsxChartSurfaceFillCodec.Patch(document.Root!, target.ChartAreaFill, "Presentation combo chart area", allowFrameDecorations: true);
        XlsxChartSurfaceFillCodec.Patch(plotArea, target.PlotAreaFill, "Presentation combo chart plot area");
    }

    private static void PatchComboScalar(XElement plot, string name, string value)
    {
        var existing = plot.Element(ChartNs + name) ?? throw new CodecException("unsupported_presentation_edit", $"Presentation combo chart is missing c:{name}.");
        if (!ComboScalar(existing, out _)) throw new CodecException("unsupported_presentation_edit", $"Presentation combo chart c:{name} is outside the bounded scalar profile.");
        existing.SetAttributeValue("val", value);
    }

    private static void PatchComboGapWidth(XElement plot, PresentationChart chart)
    {
        var existing = plot.Element(ChartNs + "gapWidth");
        if (!chart.HasGapWidth)
        {
            existing?.Remove();
            return;
        }
        if (existing is not null)
        {
            if (!ComboScalar(existing, out _)) throw new CodecException("unsupported_presentation_edit", "Presentation combo chart gapWidth is outside the bounded scalar profile.");
            existing.SetAttributeValue("val", chart.GapWidth);
            return;
        }
        var axis = plot.Elements(ChartNs + "axId").First();
        axis.AddBeforeSelf(new XElement(ChartNs + "gapWidth", new XAttribute("val", chart.GapWidth)));
    }

    private static void PatchComboVaryColors(XElement plot, PresentationChart chart)
    {
        var existing = plot.Element(ChartNs + "varyColors");
        if (!chart.HasVaryColors)
        {
            existing?.Remove();
            return;
        }
        var replacement = new XElement(ChartNs + "varyColors", new XAttribute("val", chart.VaryColors ? "1" : "0"));
        if (existing is not null)
        {
            if (!ComboScalar(existing, out var text) || text is not ("0" or "1" or "false" or "true"))
                throw new CodecException("unsupported_presentation_edit", "Presentation combo chart varyColors is outside the bounded scalar profile.");
            existing.ReplaceWith(replacement);
            return;
        }
        var series = plot.Element(ChartNs + "ser");
        if (series is null) plot.Add(replacement);
        else series.AddBeforeSelf(replacement);
    }

    private static void PatchComboOverlap(XElement plot, PresentationChart chart)
    {
        var existing = plot.Element(ChartNs + "overlap");
        if (!chart.HasOverlap)
        {
            existing?.Remove();
            return;
        }
        if (existing is not null)
        {
            if (!ComboScalar(existing, out var text) || !int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var current) || current is < -100 or > 100)
                throw new CodecException("unsupported_presentation_edit", "Presentation combo chart overlap is outside the bounded scalar profile.");
            existing.SetAttributeValue("val", chart.Overlap);
            return;
        }
        var axis = plot.Elements(ChartNs + "axId").First();
        axis.AddBeforeSelf(new XElement(ChartNs + "overlap", new XAttribute("val", chart.Overlap)));
    }
}
