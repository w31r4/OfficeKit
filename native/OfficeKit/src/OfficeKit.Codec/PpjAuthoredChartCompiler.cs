using System.Globalization;
using System.Text.Json;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal static partial class PpjAuthoredPresentationCompiler
{
    /// <summary>
    /// Lowers every authored chart family behind one content-level boundary.
    /// The presentation compiler chooses where an element belongs; this module
    /// owns whether a chart becomes native DrawingML or editable vector shapes.
    /// </summary>
    private static class ChartCompiler
    {
        internal static void BuildInto(
            PresentationElement output,
            PpjChartElementModel chart,
            JsonElement raw,
            Catalog catalog)
        {
            if (chart.Data.Series.Any(series => series.Levels is not null) &&
                chart.ChartType is not ("treemap" or "sunburst"))
                throw Unsupported(chart.Id, "levels applies only to treemap and sunburst charts");

            switch (chart.ChartType)
            {
                case "heatmap":
                    output.Group = BuildHeatmap(chart, raw, catalog);
                    break;
                case "candlestick":
                    output.Group = BuildCandlestick(chart, raw, catalog);
                    break;
                case "treemap":
                    output.Group = BuildTreemap(chart, raw, catalog);
                    break;
                case "sunburst":
                    output.Group = BuildSunburst(chart, raw, catalog);
                    break;
                case "sankey":
                    output.Group = BuildSankey(chart, raw, catalog);
                    break;
                case "area" when IsStreamgraph(chart, raw, catalog):
                    output.Group = BuildStreamgraph(chart, raw, catalog);
                    break;
                default:
                    if (IsPictographicChart(chart))
                        output.Group = BuildPictographicChart(chart, raw, catalog);
                    else if (IsNumericCombo(chart) || IsExplicitlySizedBubble(chart, raw, catalog))
                        output.Group = BuildVectorNumericChart(chart, raw, catalog);
                    else
                        output.Chart = BuildChart(chart, raw, catalog);
                    break;
            }
        }

    private static PresentationChart BuildChart(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        var isWaterfall = element.ChartType == "waterfall";
        var chart = new PresentationChart
        {
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
            Type = isWaterfall ? SpreadsheetChartType.Bar : ChartType(element.ChartType),
            Title = element.Title is null ? string.Empty : Flatten(element.Title),
            BarDirection = element.ChartType == "bar" ? "bar" : element.ChartType is "column" or "waterfall" ? "column" : string.Empty,
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) chart.FrameTransform = frameTransform;
        chart.Categories.Add(element.Data.Categories.Select(CategoryText));
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        if (isWaterfall) ValidateWaterfallCompileProfile(element, raw, namedStyle, inlineStyle);
        ApplyChartStyle(chart, namedStyle, inlineStyle, catalog, element.Id, element.Frame.Width, element.Frame.Height, raw);
        if (raw.TryGetProperty("title", out var title) && title.ValueKind != JsonValueKind.String)
        {
            chart.TitleBody = BuildTextBody(title, null, null, catalog);
            ApplyChartTitleDefaults(chart.TitleBody, chart.TitleTextStyle);
        }

        var rawXAxis = Property(raw, "xAxis");
        var rawYAxis = Property(raw, "yAxis");
        var rawSpokeAxis = Property(raw, "spokeAxis");
        if (rawSpokeAxis is { } spokeAxis)
        {
            if (chart.Type != SpreadsheetChartType.Radar)
                throw Unsupported(element.Id, "spokeAxis applies only to radar charts");
            if (rawXAxis is not null || rawYAxis is not null || Property(raw, "secondaryXAxis") is not null || Property(raw, "secondaryYAxis") is not null)
                throw Unsupported(element.Id, "spokeAxis cannot be combined with generic or secondary chart axes");
            if (FirstProperty(inlineStyle, namedStyle, "showCategoryAxis") is not null ||
                FirstProperty(inlineStyle, namedStyle, "showValueAxis") is not null ||
                FirstProperty(inlineStyle, namedStyle, "showGridlines") is not null)
                throw Unsupported(element.Id, "spokeAxis cannot be combined with legacy chart-axis visibility fields");
            (chart.XAxis, chart.YAxis) = BuildRadarSpokeAxes(spokeAxis, catalog);
        }
        else if (rawXAxis is not null || rawYAxis is not null)
        {
            if (chart.Type is SpreadsheetChartType.Pie or SpreadsheetChartType.Doughnut)
                throw Unsupported(element.Id, "pie and doughnut charts cannot define axes");
            if (FirstProperty(inlineStyle, namedStyle, "showCategoryAxis") is not null ||
                FirstProperty(inlineStyle, namedStyle, "showValueAxis") is not null)
                throw Unsupported(element.Id, "structured axes cannot be combined with legacy axis visibility fields");
            chart.XAxis = rawXAxis is { } xAxis ? BuildChartAxis(xAxis, catalog) : new SpreadsheetChartAxisArtifact();
            chart.YAxis = rawYAxis is { } yAxis ? BuildChartAxis(yAxis, catalog) : new SpreadsheetChartAxisArtifact();
        }
        var rawSecondaryXAxis = Property(raw, "secondaryXAxis");
        var rawSecondaryYAxis = Property(raw, "secondaryYAxis");
        if (rawSecondaryXAxis is not null || rawSecondaryYAxis is not null)
        {
            if (chart.Type != SpreadsheetChartType.Combo)
                throw Unsupported(element.Id, "secondary axes require a combo chart");
            chart.SecondaryXAxis = rawSecondaryXAxis is { } xAxis
                ? BuildChartAxis(xAxis, catalog)
                : new SpreadsheetChartAxisArtifact();
            chart.SecondaryYAxis = rawSecondaryYAxis is { } yAxis
                ? BuildChartAxis(yAxis, catalog)
                : new SpreadsheetChartAxisArtifact();
        }

        var seriesJson = raw.GetProperty("data").GetProperty("series").EnumerateArray().ToArray();
        if (isWaterfall)
        {
            BuildWaterfallSeries(chart, element, namedStyle, inlineStyle, catalog);
            ApplyAccessibility(chart, element.Accessibility);
            return chart;
        }
        for (var index = 0; index < element.Data.Series.Count; index++)
        {
            var source = element.Data.Series[index];
            if (element.ChartType != "combo") RejectProperties(seriesJson[index], element.Id, "chartType", "axis");
            var effectiveType = ChartType(source.ChartType ?? element.ChartType);
            var series = BuildSeries(source, seriesJson[index], catalog, effectiveType);
            if (chart.Type == SpreadsheetChartType.Combo)
            {
                if (source.ChartType is "bar" or "column")
                {
                    var direction = source.ChartType == "bar" ? "bar" : "column";
                    if (chart.BarDirection.Length > 0 && chart.BarDirection != direction)
                        throw Unsupported(element.Id, "combo charts cannot mix horizontal bars and vertical columns");
                    chart.BarDirection = direction;
                }
                chart.ComboSeries.Add(new PresentationComboSeriesArtifact
                {
                    Type = ChartType(source.ChartType!),
                    AxisGroup = source.Axis == "secondary" || source.XAxisIndex == 1 || source.YAxisIndex == 1
                        ? PresentationChartAxisGroup.Secondary
                        : PresentationChartAxisGroup.Primary,
                    Series = series,
                });
            }
            else chart.Series.Add(series);
        }
        ApplyAccessibility(chart, element.Accessibility);
        return chart;
    }

    private static bool IsNumericCombo(PpjChartElementModel element) =>
        element.ChartType == "combo" &&
        element.Data.Series.Any(series => series.ChartType is "scatter" or "bubble");

    private static bool IsExplicitlySizedBubble(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        if (element.ChartType != "bubble") return false;
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        return FirstProperty(inlineStyle, namedStyle, "bubbleSizeScale") is not null ||
               FirstProperty(inlineStyle, namedStyle, "bubbleRadiusRange") is not null;
    }

    private sealed record VectorAxisRange(
        double Minimum,
        double Maximum,
        double MajorUnit,
        int TickCount,
        bool Reverse,
        string? NumberFormat);

    private static PresentationGroup BuildVectorNumericChart(
        PpjChartElementModel element,
        JsonElement raw,
        Catalog catalog)
    {
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ValidateVectorNumericCompileProfile(element, raw, namedStyle, inlineStyle);

        var seriesJson = raw.GetProperty("data").GetProperty("series").EnumerateArray().ToArray();
        var xAxis = Property(raw, "xAxis");
        var yAxis = Property(raw, "yAxis");
        var showXAxis = xAxis is null || OptionalBoolean(xAxis.Value, "visible") != false;
        var showYAxis = yAxis is null || OptionalBoolean(yAxis.Value, "visible") != false;
        var legend = FirstProperty(inlineStyle, namedStyle, "legend")?.GetString() ?? "right";
        var titleText = element.Title is null ? string.Empty : Flatten(element.Title);
        var titleHeight = titleText.Length == 0 ? 0 : Math.Clamp(element.Frame.Height * 0.12, 22, 38);
        var legendWidth = legend == "right" ? Math.Clamp(element.Frame.Width * 0.18, 92, 138) : 0;
        var legendGap = legendWidth > 0 ? 12 : 0;
        var leftInset = showYAxis ? Math.Clamp(element.Frame.Width * 0.1, 42, 66) : 8;
        var bottomInset = showXAxis
            ? (xAxis is { } xAxisValue && xAxisValue.TryGetProperty("title", out _) ? 38 : 24)
            : 8;
        var topInset = titleHeight + 8;
        var rightInset = legendWidth + legendGap + 8;
        var plotX = element.Frame.X + leftInset;
        var plotY = element.Frame.Y + topInset;
        var plotWidth = element.Frame.Width - leftInset - rightInset;
        var plotHeight = element.Frame.Height - topInset - bottomInset;
        if (plotWidth < 120 || plotHeight < 90)
            throw Unsupported(element.Id, "numeric combo frame is too small for editable axes and marks");

        var xValues = element.Data.Series.SelectMany(series => series.XValues).ToArray();
        var yValues = element.Data.Series.SelectMany(series => series.Values).Select(value => value!.Value).ToList();
        var requiresZero = element.Data.Series.Any(series => NumericSeriesType(element, series) is "area" or "column");
        var xRange = BuildVectorAxisRange(xValues, xAxis, includeZero: false, element.Id, "X");
        var yRange = BuildVectorAxisRange(yValues, yAxis, requiresZero, element.Id, "Y");
        if (requiresZero && (yRange.Minimum > 0 || yRange.Maximum < 0))
            throw Unsupported(element.Id, "numeric combo area and column series require zero inside the Y domain");

        double PlotX(double value)
        {
            var ratio = (value - xRange.Minimum) / (xRange.Maximum - xRange.Minimum);
            if (xRange.Reverse) ratio = 1 - ratio;
            return plotX + ratio * plotWidth;
        }
        double PlotY(double value)
        {
            var ratio = (value - yRange.Minimum) / (yRange.Maximum - yRange.Minimum);
            if (yRange.Reverse) ratio = 1 - ratio;
            return plotY + (1 - ratio) * plotHeight;
        }

        var group = new PresentationGroup
        {
            LeftEmu = Emu(element.Frame.X),
            TopEmu = Emu(element.Frame.Y),
            WidthEmu = Emu(element.Frame.Width),
            HeightEmu = Emu(element.Frame.Height),
            ChildLeftEmu = Emu(element.Frame.X),
            ChildTopEmu = Emu(element.Frame.Y),
            ChildWidthEmu = Emu(element.Frame.Width),
            ChildHeightEmu = Emu(element.Frame.Height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;

        AddVectorGridlines(group, element.Id, "numeric", plotX, plotY, plotWidth, plotHeight, xRange, yRange, xAxis, yAxis, catalog);

        var ordered = element.Data.Series
            .Select((series, index) => (Series: series, Json: seriesJson[index], Index: index, Type: NumericSeriesType(element, series)))
            .OrderBy(item => NumericSeriesOrder(item.Type))
            .ThenBy(item => item.Index)
            .ToArray();
        var columnSeries = ordered.Where(item => item.Type == "column").ToArray();
        var columnPositions = columnSeries.Select((item, index) => (item.Series.Id, index))
            .ToDictionary(item => item.Id, item => item.index, StringComparer.Ordinal);
        var xSpacing = element.Data.Series
            .SelectMany(series => series.XValues.Zip(series.XValues.Skip(1)).Select(pair => pair.Second - pair.First))
            .DefaultIfEmpty((xRange.Maximum - xRange.Minimum) / 8)
            .Min();
        var slotWidth = Math.Clamp(Math.Abs(PlotX(xRange.Minimum + xSpacing) - PlotX(xRange.Minimum)) * 0.7, 4, 54);
        var columnWidth = columnSeries.Length == 0 ? 0 : Math.Max(2, slotWidth / columnSeries.Length);
        var baselineY = PlotY(0);
        var bubbleSizes = ordered
            .Where(item => item.Type == "bubble")
            .SelectMany(item => item.Series.BubbleSizes)
            .ToArray();
        var bubbleScale = FirstProperty(inlineStyle, namedStyle, "bubbleScale")?.GetInt32() ?? 100;
        var bubbleSizeMode = FirstProperty(inlineStyle, namedStyle, "bubbleSizeMode")?.GetString() ?? "area";
        var bubbleSizeScale = FirstProperty(inlineStyle, namedStyle, "bubbleSizeScale")?.GetString();
        var bubbleRadiusRange = FirstProperty(inlineStyle, namedStyle, "bubbleRadiusRange") is { } radiusRange
            ? radiusRange.EnumerateArray().Select(item => item.GetDouble()).ToArray()
            : null;

        foreach (var item in ordered)
        {
            var series = item.Series;
            var points = series.XValues.Zip(series.Values, (xValue, value) => (X: PlotX(xValue), Y: PlotY(value!.Value), Value: value.Value)).ToArray();
            switch (item.Type)
            {
                case "area":
                {
                    var shape = ShapeFrame(new PpjFrameModel(plotX, plotY, plotWidth, plotHeight, 0, false, false), "custom");
                    ApplyVectorSeriesFill(shape, item.Json, catalog, item.Index, element.Id, defaultOpacity: 0.32);
                    shape.CustomPaths.Add(BuildVectorAreaPath(plotX, plotY, plotWidth, plotHeight, points.Select(point => (point.X, point.Y)).ToArray(), baselineY));
                    group.Children.Add(new PresentationElement
                    {
                        Id = NumericComboNativeId(element.Id, $"area/{series.Id}"),
                        Name = $"numeric area {series.Name}",
                        Shape = shape,
                    });
                    break;
                }
                case "column":
                {
                    var seriesOffset = columnPositions[series.Id] - (columnSeries.Length - 1) / 2d;
                    for (var pointIndex = 0; pointIndex < points.Length; pointIndex++)
                    {
                        var point = points[pointIndex];
                        var top = Math.Min(point.Y, baselineY);
                        var height = Math.Max(1.5, Math.Abs(point.Y - baselineY));
                        var shape = ShapeFrame(new PpjFrameModel(point.X + seriesOffset * columnWidth - columnWidth * 0.42, top, columnWidth * 0.84, height, 0, false, false), "rect");
                        ApplyVectorSeriesFill(shape, item.Json, catalog, item.Index, element.Id, defaultOpacity: 0.82);
                        group.Children.Add(new PresentationElement
                        {
                            Id = NumericComboNativeId(element.Id, $"column/{series.Id}/{pointIndex}"),
                            Name = $"numeric column {series.Name} {pointIndex + 1}",
                            Shape = shape,
                        });
                    }
                    break;
                }
                case "line":
                    AddVectorLineSeries(group, element.Id, "numeric", series, item.Json, item.Index, points, catalog, drawDefaultMarkers: false);
                    break;
                case "scatter":
                    AddVectorPointSeries(group, element.Id, "numeric", series, item.Json, item.Index, points, null, plotWidth, plotHeight, catalog);
                    break;
                case "bubble":
                    AddVectorPointSeries(group, element.Id, "numeric", series, item.Json, item.Index, points, series.BubbleSizes, plotWidth, plotHeight, catalog,
                        bubbleScale: bubbleScale,
                        bubbleSizeMode: bubbleSizeMode,
                        bubbleSizeScale: bubbleSizeScale,
                        bubbleDomainMinimum: bubbleSizes.Min(),
                        bubbleDomainMaximum: bubbleSizes.Max(),
                        bubbleRadiusRange: bubbleRadiusRange);
                    break;
            }
        }

        AddVectorAxesAndLabels(group, element.Id, "numeric", plotX, plotY, plotWidth, plotHeight, xRange, yRange, xAxis, yAxis, catalog);

        if (titleText.Length > 0)
            group.Children.Add(VectorChartTitleElement(
                NumericComboNativeId(element.Id, "title"),
                "numeric chart title",
                element.Frame.X,
                element.Frame.Y,
                element.Frame.Width,
                titleHeight,
                raw.GetProperty("title"),
                FirstProperty(inlineStyle, namedStyle, "titleTextStyle"),
                catalog,
                Math.Clamp(titleHeight * 0.48, 11, 18),
                "left"));
        if (legend == "right")
            AddVectorLegend(group, element.Id, "numeric", ordered.Select(item => (item.Series, item.Json, item.Index, item.Type)).ToArray(),
                plotX + plotWidth + legendGap, plotY, legendWidth, plotHeight,
                FirstProperty(inlineStyle, namedStyle, "legendTextStyle"), catalog);

        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static void ValidateVectorNumericCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        var numericCombo = IsNumericCombo(element);
        var explicitlySizedBubble = element.ChartType == "bubble" &&
            (FirstProperty(inlineStyle, namedStyle, "bubbleSizeScale") is not null ||
             FirstProperty(inlineStyle, namedStyle, "bubbleRadiusRange") is not null);
        if (!numericCombo && !explicitlySizedBubble)
            throw Unsupported(element.Id, "vector numeric charts require a numeric combo or explicit bubble sizing");
        var minimumSeries = numericCombo ? 2 : 1;
        if (element.Data.Categories.Count != 0 || element.Data.Series.Count < minimumSeries || element.Data.Series.Count > 8)
            throw Unsupported(element.Id, $"vector numeric charts require {minimumSeries}..8 series and an empty categories array");
        var families = element.Data.Series.Select(series => series.ChartType).Distinct(StringComparer.Ordinal).ToArray();
        if (numericCombo && (!families.Any(family => family is "scatter" or "bubble") || families.Length < 2))
            throw Unsupported(element.Id, "numeric combo charts require scatter or bubble evidence and a different plot family");
        foreach (var series in element.Data.Series)
        {
            var seriesType = NumericSeriesType(element, series);
            if (seriesType is not ("scatter" or "bubble" or "line" or "area" or "column"))
                throw Unsupported(element.Id, "numeric combo series types are scatter, bubble, line, area and column");
            if (series.Values.Count is < 2 or > 64 || series.Values.Any(value => value is null || !double.IsFinite(value.Value)) ||
                series.XValues.Count != series.Values.Count || series.XValues.Any(value => !double.IsFinite(value)) ||
                series.XValues.Zip(series.XValues.Skip(1)).Any(pair => pair.First >= pair.Second))
                throw Unsupported(element.Id, $"numeric combo series {series.Id} requires 2..64 finite values with strictly increasing aligned xValues");
            if (seriesType == "bubble")
            {
                if (series.BubbleSizes.Count != series.Values.Count || series.BubbleSizes.Any(value => !double.IsFinite(value) || value <= 0))
                    throw Unsupported(element.Id, $"bubble series {series.Id} requires one finite positive size per point");
            }
            else if (series.BubbleSizes.Count != 0)
                throw Unsupported(element.Id, "bubbleSizes apply only to a bubble series");
            if (series.Raw.TryGetProperty("marker", out var marker))
            {
                if (seriesType is "area" or "column" or "bubble")
                    throw Unsupported(element.Id, $"numeric combo {seriesType} series do not render markers");
                if (seriesType == "scatter" &&
                    (marker.ValueKind == JsonValueKind.String && marker.GetString() == "none" ||
                     marker.ValueKind == JsonValueKind.Object && OptionalString(marker, "symbol") == "none"))
                    throw Unsupported(element.Id, "numeric combo scatter series cannot use marker none");
            }
            if (series.Axis is not null and not "primary")
                throw Unsupported(element.Id, "numeric combo charts use one shared primary axis pair");
            RejectProperties(series.Raw, element.Id, "pointRoles", "openValues", "highValues", "lowValues", "parents", "sources", "targets", "symbol", "trendlines", "errorBars");
            if (series.Raw.TryGetProperty("fill", out _) && series.Raw.TryGetProperty("color", out _))
                throw Unsupported(element.Id, "numeric combo series color and fill are aliases and cannot both be present");
        }
        foreach (var property in new[] { "secondaryXAxis", "secondaryYAxis" })
            if (raw.TryGetProperty(property, out _)) throw Unsupported(element.Id, "numeric combo charts do not support secondary axes");
        foreach (var style in new[] { namedStyle, inlineStyle })
        {
            if (style is not { ValueKind: JsonValueKind.Object } value) continue;
            foreach (var property in value.EnumerateObject())
                if (property.Name is not ("legend" or "titleTextStyle" or "legendTextStyle" or "bubbleScale" or "bubbleSizeMode" or "bubbleSizeScale" or "bubbleRadiusRange"))
                    throw Unsupported(element.Id, $"numeric combo charts do not support chart style field {property.Name}");
        }
        if (FirstProperty(inlineStyle, namedStyle, "legend") is { } legend && legend.GetString() is not ("none" or "right"))
            throw Unsupported(element.Id, "numeric combo legends support only none or right");
        if (FirstProperty(inlineStyle, namedStyle, "bubbleScale") is { } bubbleScale && bubbleScale.GetInt32() < 10)
            throw Unsupported(element.Id, "generated numeric combo bubbles require bubbleScale between 10 and 300");
        if (FirstProperty(inlineStyle, namedStyle, "bubbleRadiusRange") is { } radiusRange)
        {
            var radii = radiusRange.EnumerateArray().Select(item => item.GetDouble()).ToArray();
            if (radii.Length != 2 || radii[0] >= radii[1])
                throw Unsupported(element.Id, "bubbleRadiusRange requires a strictly increasing [minimum, maximum] pair");
        }
        ValidateVectorAxisCompileProfile(Property(raw, "xAxis"), element.Id, "X");
        ValidateVectorAxisCompileProfile(Property(raw, "yAxis"), element.Id, "Y");
    }

    private static void ValidateVectorAxisCompileProfile(JsonElement? axis, string elementId, string name)
    {
        if (axis is not { ValueKind: JsonValueKind.Object } value) return;
        foreach (var property in value.EnumerateObject())
            if (property.Name is not ("visible" or "title" or "numberFormat" or "min" or "max" or "majorUnit" or "textStyle" or "titleTextStyle" or "reverse" or "axisLine" or "gridLine"))
                throw Unsupported(elementId, $"generated numeric {name} axis does not support {property.Name}");
        foreach (var property in value.EnumerateObject())
            if (IsTokenReference(property.Value))
                throw Unsupported(elementId, $"generated numeric {name} axis cannot resolve tokenized {property.Name}; use a native ChartPart chart for typed axis tokens");
        if (OptionalString(value, "numberFormat") is { } numberFormat && !VectorChartNumberFormats.Contains(numberFormat, StringComparer.Ordinal))
            throw Unsupported(elementId, $"generated numeric {name} axis numberFormat {numberFormat} is outside the bounded profile");
    }

    private static VectorAxisRange BuildVectorAxisRange(
        IReadOnlyCollection<double> values,
        JsonElement? axis,
        bool includeZero,
        string elementId,
        string name)
    {
        var dataMinimum = values.Min();
        var dataMaximum = values.Max();
        if (includeZero)
        {
            dataMinimum = Math.Min(dataMinimum, 0);
            dataMaximum = Math.Max(dataMaximum, 0);
        }
        var dataRange = dataMaximum - dataMinimum;
        if (dataRange <= 0) dataRange = Math.Max(1, Math.Abs(dataMaximum) * 0.1);
        var majorUnit = axis is { } explicitAxis && OptionalDouble(explicitAxis, "majorUnit") is { } explicitMajor
            ? explicitMajor
            : NiceVectorChartStep(dataRange / 5);
        var pad = dataRange * 0.05;
        var minimum = axis is { } minimumAxis && OptionalDouble(minimumAxis, "min") is { } explicitMinimum
            ? explicitMinimum
            : Math.Floor((dataMinimum - pad) / majorUnit) * majorUnit;
        var maximum = axis is { } maximumAxis && OptionalDouble(maximumAxis, "max") is { } explicitMaximum
            ? explicitMaximum
            : Math.Ceiling((dataMaximum + pad) / majorUnit) * majorUnit;
        if (minimum > dataMinimum || maximum < dataMaximum || maximum <= minimum)
            throw Unsupported(elementId, $"numeric combo {name} axis must contain every data value and have maximum greater than minimum");
        var tickCount = checked((int)Math.Floor((maximum - minimum) / majorUnit + 1e-9)) + 1;
        if (tickCount is < 2 or > 12)
            throw Unsupported(elementId, $"numeric combo {name} axis must expand to 2..12 major ticks");
        return new VectorAxisRange(
            minimum,
            maximum,
            majorUnit,
            tickCount,
            axis is { } reversed && OptionalBoolean(reversed, "reverse") == true,
            axis is { } formatted ? OptionalString(formatted, "numberFormat") : null);
    }

    private static int NumericSeriesOrder(string chartType) => chartType switch
    {
        "area" => 0,
        "column" => 1,
        "line" => 2,
        "scatter" => 3,
        "bubble" => 4,
        _ => 5,
    };

    private static string NumericSeriesType(PpjChartElementModel chart, PpjChartSeriesModel series) =>
        chart.ChartType == "bubble" ? "bubble" : series.ChartType!;

    private static string NumericComboNativeId(string elementId, string suffix) =>
        $"{elementId}/numeric-combo/{suffix}";

    private static readonly string[] VectorChartNumberFormats = ["0", "0.0", "0.00", "#,##0", "#,##0.0", "#,##0.00"];

    private static void AddVectorGridlines(
        PresentationGroup group,
        string elementId,
        string family,
        double plotX,
        double plotY,
        double plotWidth,
        double plotHeight,
        VectorAxisRange xRange,
        VectorAxisRange yRange,
        JsonElement? xAxis,
        JsonElement? yAxis,
        Catalog catalog)
    {
        if (AxisLineSetting(yAxis, "gridLine") is { Visible: true } yGrid)
            for (var index = 0; index < yRange.TickCount; index++)
            {
                var ratio = index * yRange.MajorUnit / (yRange.Maximum - yRange.Minimum);
                if (yRange.Reverse) ratio = 1 - ratio;
                var y = plotY + (1 - ratio) * plotHeight;
                group.Children.Add(VectorStyledLineElement(
                    $"{elementId}/{family}/y-grid/{index}",
                    $"{family} Y gridline {index + 1}",
                    plotX,
                    y,
                    plotX + plotWidth,
                    y,
                    yGrid.Style,
                    catalog,
                    "D9DFE8",
                    0.75));
            }
        if (AxisLineSetting(xAxis, "gridLine") is { Visible: true } xGrid)
            for (var index = 0; index < xRange.TickCount; index++)
            {
                var ratio = index * xRange.MajorUnit / (xRange.Maximum - xRange.Minimum);
                if (xRange.Reverse) ratio = 1 - ratio;
                var x = plotX + ratio * plotWidth;
                group.Children.Add(VectorStyledLineElement(
                    $"{elementId}/{family}/x-grid/{index}",
                    $"{family} X gridline {index + 1}",
                    x,
                    plotY,
                    x,
                    plotY + plotHeight,
                    xGrid.Style,
                    catalog,
                    "D9DFE8",
                    0.75));
            }
    }

    private static void AddVectorAxesAndLabels(
        PresentationGroup group,
        string elementId,
        string family,
        double plotX,
        double plotY,
        double plotWidth,
        double plotHeight,
        VectorAxisRange xRange,
        VectorAxisRange yRange,
        JsonElement? xAxis,
        JsonElement? yAxis,
        Catalog catalog)
    {
        var showXAxis = xAxis is null || OptionalBoolean(xAxis.Value, "visible") != false;
        var showYAxis = yAxis is null || OptionalBoolean(yAxis.Value, "visible") != false;
        if (showXAxis)
        {
            if (AxisLineSetting(xAxis, "axisLine", defaultVisible: true) is { Visible: true } axisLine)
                group.Children.Add(VectorStyledLineElement(
                    $"{elementId}/{family}/x-axis",
                    $"{family} X axis",
                    plotX,
                    plotY + plotHeight,
                    plotX + plotWidth,
                    plotY + plotHeight,
                    axisLine.Style,
                    catalog,
                    "52606D",
                    1));
            var labelStyle = Property(xAxis, "textStyle");
            var labelWidth = Math.Clamp(plotWidth / Math.Min(xRange.TickCount, 8), 42, 88);
            for (var index = 0; index < xRange.TickCount; index++)
            {
                var ratio = index * xRange.MajorUnit / (xRange.Maximum - xRange.Minimum);
                var value = xRange.Minimum + index * xRange.MajorUnit;
                if (xRange.Reverse) ratio = 1 - ratio;
                var x = plotX + ratio * plotWidth;
                group.Children.Add(VectorChartTextElement(
                    $"{elementId}/{family}/x-label/{index}",
                    $"{family} X label {index + 1}",
                    Math.Clamp(x - labelWidth / 2, plotX, plotX + plotWidth - labelWidth),
                    plotY + plotHeight + 2,
                    labelWidth,
                    16,
                    FormatVectorChartValue(value, xRange.NumberFormat),
                    labelStyle,
                    catalog,
                    7,
                    index == 0 ? "left" : index == xRange.TickCount - 1 ? "right" : "center",
                    "52606D"));
            }
            if (xAxis is { } xAxisWithTitle && xAxisWithTitle.TryGetProperty("title", out var title))
                group.Children.Add(VectorChartTextElement(
                    $"{elementId}/{family}/x-title",
                    $"{family} X axis title",
                    plotX,
                    plotY + plotHeight + 18,
                    plotWidth,
                    16,
                    title.GetString()!,
                    Property(xAxis, "titleTextStyle"),
                    catalog,
                    8,
                    "center",
                    "52606D"));
        }
        if (showYAxis)
        {
            if (AxisLineSetting(yAxis, "axisLine", defaultVisible: true) is { Visible: true } axisLine)
                group.Children.Add(VectorStyledLineElement(
                    $"{elementId}/{family}/y-axis",
                    $"{family} Y axis",
                    plotX,
                    plotY,
                    plotX,
                    plotY + plotHeight,
                    axisLine.Style,
                    catalog,
                    "52606D",
                    1));
            var labelStyle = Property(yAxis, "textStyle");
            for (var index = 0; index < yRange.TickCount; index++)
            {
                var ratio = index * yRange.MajorUnit / (yRange.Maximum - yRange.Minimum);
                var value = yRange.Minimum + index * yRange.MajorUnit;
                if (yRange.Reverse) ratio = 1 - ratio;
                var y = plotY + (1 - ratio) * plotHeight;
                group.Children.Add(VectorChartTextElement(
                    $"{elementId}/{family}/y-label/{index}",
                    $"{family} Y label {index + 1}",
                    plotX - 52,
                    y - 7,
                    46,
                    14,
                    FormatVectorChartValue(value, yRange.NumberFormat),
                    labelStyle,
                    catalog,
                    7,
                    "right",
                    "52606D"));
            }
            if (yAxis is { } yAxisWithTitle && yAxisWithTitle.TryGetProperty("title", out var title))
                group.Children.Add(VectorChartTextElement(
                    $"{elementId}/{family}/y-title",
                    $"{family} Y axis title",
                    plotX - 52,
                    plotY - 18,
                    Math.Min(140, plotWidth * 0.35),
                    16,
                    title.GetString()!,
                    Property(yAxis, "titleTextStyle"),
                    catalog,
                    8,
                    "left",
                    "52606D"));
        }
    }

    private sealed record VectorLineSetting(bool Visible, JsonElement? Style);

    private static VectorLineSetting? AxisLineSetting(JsonElement? axis, string property, bool defaultVisible = false)
    {
        if (axis is null || !axis.Value.TryGetProperty(property, out var value))
            return defaultVisible ? new VectorLineSetting(true, null) : null;
        return value.ValueKind == JsonValueKind.False
            ? new VectorLineSetting(false, null)
            : value.ValueKind == JsonValueKind.True
                ? new VectorLineSetting(true, null)
                : new VectorLineSetting(true, value);
    }

    private static PresentationElement VectorStyledLineElement(
        string id,
        string name,
        double startX,
        double startY,
        double endX,
        double endY,
        JsonElement? style,
        Catalog catalog,
        string defaultRgb,
        double defaultWidth)
    {
        var connector = new PresentationConnector
        {
            ConnectorType = "straight",
            StartXEmu = Emu(startX),
            StartYEmu = Emu(startY),
            EndXEmu = Emu(endX),
            EndYEmu = Emu(endY),
        };
        if (style is { } stroke) ApplyLine(connector, stroke, catalog);
        else
        {
            connector.LineRgb = defaultRgb;
            connector.LineWidthEmu = Emu(defaultWidth);
            connector.LineStyle = "solid";
        }
        return new PresentationElement { Id = id, Name = name, Connector = connector };
    }

    private static PresentationCustomGeometryPath BuildVectorAreaPath(
        double plotX,
        double plotY,
        double plotWidth,
        double plotHeight,
        IReadOnlyList<(double X, double Y)> points,
        double baselineY)
    {
        const long viewport = 100_000;
        PresentationCustomGeometryPoint Point((double X, double Y) source) => new()
        {
            X = checked((long)Math.Round((source.X - plotX) / plotWidth * viewport, MidpointRounding.AwayFromZero)),
            Y = checked((long)Math.Round((source.Y - plotY) / plotHeight * viewport, MidpointRounding.AwayFromZero)),
        };
        var path = new PresentationCustomGeometryPath
        {
            Width = viewport,
            Height = viewport,
            FillMode = PresentationCustomGeometryPath.Types.FillMode.Normal,
            Stroke = true,
        };
        path.Commands.Add(MoveTo(Point((points[0].X, baselineY))));
        foreach (var point in points) path.Commands.Add(LineTo(Point(point)));
        path.Commands.Add(LineTo(Point((points[^1].X, baselineY))));
        path.Commands.Add(new PresentationCustomGeometryCommand { Close = true });
        return path;
    }

    private static void ApplyVectorSeriesFill(
        PresentationShape target,
        JsonElement series,
        Catalog catalog,
        int seriesIndex,
        string elementId,
        double defaultOpacity)
    {
        if (series.TryGetProperty("fill", out var fill))
        {
            if (fill.GetProperty("type").GetString() == "none")
                throw Unsupported(elementId, "filled numeric marks cannot use fill none");
            ApplyTextBoxFill(target, fill, catalog);
        }
        else if (series.TryGetProperty("color", out var color))
        {
            var resolved = catalog.Color(color);
            target.FillRgb = resolved.Rgb;
            if (resolved.Alpha < 1) target.FillOpacityThousandthPercent = Opacity(resolved.Alpha);
        }
        else target.FillRgb = catalog.Theme.AccentRgb[seriesIndex % catalog.Theme.AccentRgb.Count];
        if (defaultOpacity < 1 && !target.HasFillOpacityThousandthPercent && target.GradientFill is null)
            target.FillOpacityThousandthPercent = Opacity(defaultOpacity);
        if (series.TryGetProperty("stroke", out var stroke)) ApplyLine(target, stroke, catalog);
        else target.LineStyle = "none";
    }

    private static (string Rgb, double Alpha) VectorSeriesColor(JsonElement series, Catalog catalog, int seriesIndex)
    {
        if (series.TryGetProperty("color", out var color))
        {
            var resolved = catalog.Color(color);
            return (resolved.Rgb, resolved.Alpha);
        }
        if (series.TryGetProperty("fill", out var fill) && fill.GetProperty("type").GetString() == "solid")
        {
            var resolved = FillColor(fill, catalog);
            if (resolved is { } value) return (value.Rgb, value.Opacity);
        }
        return (catalog.Theme.AccentRgb[seriesIndex % catalog.Theme.AccentRgb.Count], 1);
    }

    private static void ApplyVectorSeriesLine(
        PresentationConnector connector,
        JsonElement series,
        Catalog catalog,
        int seriesIndex)
    {
        if (series.TryGetProperty("stroke", out var stroke))
        {
            ApplyLine(connector, stroke, catalog);
            return;
        }
        var color = VectorSeriesColor(series, catalog, seriesIndex);
        connector.LineRgb = color.Rgb;
        connector.LineWidthEmu = Emu(1.5);
        connector.LineStyle = "solid";
        connector.LineCap = "round";
        connector.LineJoin = "round";
        if (color.Alpha < 1) connector.LineOpacityThousandthPercent = Opacity(color.Alpha);
    }

    private static void AddVectorLineSeries(
        PresentationGroup group,
        string elementId,
        string family,
        PpjChartSeriesModel series,
        JsonElement seriesJson,
        int seriesIndex,
        IReadOnlyList<(double X, double Y, double Value)> points,
        Catalog catalog,
        bool drawDefaultMarkers)
    {
        for (var index = 0; index < points.Count - 1; index++)
        {
            var connector = new PresentationConnector
            {
                ConnectorType = "straight",
                StartXEmu = Emu(points[index].X),
                StartYEmu = Emu(points[index].Y),
                EndXEmu = Emu(points[index + 1].X),
                EndYEmu = Emu(points[index + 1].Y),
            };
            ApplyVectorSeriesLine(connector, seriesJson, catalog, seriesIndex);
            group.Children.Add(new PresentationElement
            {
                Id = $"{elementId}/{family}/line/{series.Id}/{index}",
                Name = $"{family} line {series.Name} {index + 1}",
                Connector = connector,
            });
        }
        var marker = seriesJson.TryGetProperty("marker", out var configuredMarker) ? configuredMarker : default;
        if (drawDefaultMarkers || marker.ValueKind != JsonValueKind.Undefined)
            AddVectorMarkers(group, elementId, family, series, seriesJson, seriesIndex, points, marker, catalog);
    }

    private static void AddVectorPointSeries(
        PresentationGroup group,
        string elementId,
        string family,
        PpjChartSeriesModel series,
        JsonElement seriesJson,
        int seriesIndex,
        IReadOnlyList<(double X, double Y, double Value)> points,
        IReadOnlyList<double>? bubbleSizes,
        double plotWidth,
        double plotHeight,
        Catalog catalog,
        int bubbleScale = 100,
        string bubbleSizeMode = "area",
        string? bubbleSizeScale = null,
        double? bubbleDomainMinimum = null,
        double? bubbleDomainMaximum = null,
        IReadOnlyList<double>? bubbleRadiusRange = null)
    {
        if (bubbleSizes is null)
        {
            var marker = seriesJson.TryGetProperty("marker", out var configuredMarker) ? configuredMarker : default;
            AddVectorMarkers(group, elementId, family, series, seriesJson, seriesIndex, points, marker, catalog, defaultSymbol: "circle");
            return;
        }
        var minimum = bubbleDomainMinimum ?? bubbleSizes.Min();
        var maximum = bubbleDomainMaximum ?? bubbleSizes.Max();
        var maximumDiameter = Math.Clamp(Math.Min(plotWidth, plotHeight) * 0.12 * bubbleScale / 100d, 10, 48);
        var explicitSizing = bubbleSizeScale is not null || bubbleRadiusRange is not null;
        var scale = bubbleSizeScale ?? (bubbleSizeMode == "width" ? "linear" : "sqrt");
        var minimumRadius = bubbleRadiusRange is null ? 2d : bubbleRadiusRange[0];
        var maximumRadius = bubbleRadiusRange is null ? maximumDiameter / 2 : bubbleRadiusRange[1];
        double Transform(double value) => scale switch
        {
            "linear" => value,
            "log" => Math.Log(value),
            _ => Math.Sqrt(value),
        };
        var transformedMinimum = Transform(minimum);
        var transformedMaximum = Transform(maximum);
        for (var index = 0; index < points.Count; index++)
        {
            double diameter;
            if (explicitSizing)
            {
                var ratio = transformedMaximum == transformedMinimum
                    ? 0.5
                    : (Transform(bubbleSizes[index]) - transformedMinimum) / (transformedMaximum - transformedMinimum);
                var radius = minimumRadius + Math.Clamp(ratio, 0, 1) * (maximumRadius - minimumRadius);
                diameter = radius * 2;
            }
            else
            {
                var ratio = bubbleSizes[index] / maximum;
                diameter = Math.Max(4, maximumDiameter * (bubbleSizeMode == "width" ? ratio : Math.Sqrt(ratio)));
            }
            group.Children.Add(BuildVectorMarkerElement(
                $"{elementId}/{family}/bubble/{series.Id}/{index}",
                $"{family} bubble {series.Name} {index + 1}",
                points[index].X,
                points[index].Y,
                diameter,
                "circle",
                seriesJson,
                default,
                catalog,
                seriesIndex,
                defaultOpacity: 0.65));
        }
    }

    private static void AddVectorMarkers(
        PresentationGroup group,
        string elementId,
        string family,
        PpjChartSeriesModel series,
        JsonElement seriesJson,
        int seriesIndex,
        IReadOnlyList<(double X, double Y, double Value)> points,
        JsonElement marker,
        Catalog catalog,
        string? defaultSymbol = null)
    {
        var symbol = marker.ValueKind == JsonValueKind.String
            ? marker.GetString()
            : marker.ValueKind == JsonValueKind.Object
                ? OptionalString(marker, "symbol") ?? defaultSymbol ?? "circle"
                : defaultSymbol;
        if (symbol is null or "none") return;
        var size = marker.ValueKind == JsonValueKind.Object && marker.TryGetProperty("size", out var markerSize)
            ? markerSize.GetDouble()
            : symbol == "dot" ? 4 : 7;
        for (var index = 0; index < points.Count; index++)
            group.Children.Add(BuildVectorMarkerElement(
                $"{elementId}/{family}/marker/{series.Id}/{index}",
                $"{family} marker {series.Name} {index + 1}",
                points[index].X,
                points[index].Y,
                size,
                symbol,
                seriesJson,
                marker,
                catalog,
                seriesIndex,
                defaultOpacity: 1));
    }

    private static PresentationElement BuildVectorMarkerElement(
        string id,
        string name,
        double centerX,
        double centerY,
        double size,
        string symbol,
        JsonElement series,
        JsonElement marker,
        Catalog catalog,
        int seriesIndex,
        double defaultOpacity)
    {
        var geometry = symbol switch
        {
            "dot" or "circle" => "ellipse",
            "square" => "rect",
            "diamond" => "diamond",
            "triangle" => "triangle",
            "x" => "mathMultiply",
            "star" => "star5",
            "plus" => "plus",
            "dash" => "rect",
            _ => "ellipse",
        };
        var width = symbol == "dash" ? size * 1.6 : size;
        var height = symbol == "dash" ? Math.Max(1.5, size * 0.28) : size;
        var shape = ShapeFrame(new PpjFrameModel(centerX - width / 2, centerY - height / 2, width, height, 0, false, false), geometry);
        ApplyVectorSeriesFill(shape, series, catalog, seriesIndex, id, defaultOpacity);
        if (marker.ValueKind == JsonValueKind.Object)
        {
            if (marker.TryGetProperty("fill", out var markerFill))
            {
                var color = catalog.Color(markerFill);
                shape.FillRgb = color.Rgb;
                if (color.Alpha < 1) shape.FillOpacityThousandthPercent = Opacity(color.Alpha);
            }
            if (marker.TryGetProperty("stroke", out var markerStroke)) ApplyLine(shape, markerStroke, catalog);
        }
        return new PresentationElement { Id = id, Name = name, Shape = shape };
    }

    private static void AddVectorLegend(
        PresentationGroup group,
        string elementId,
        string family,
        IReadOnlyList<(PpjChartSeriesModel Series, JsonElement Json, int Index, string Type)> series,
        double x,
        double y,
        double width,
        double height,
        JsonElement? textStyle,
        Catalog catalog)
    {
        var rowHeight = Math.Min(20, height / series.Count);
        var startY = y + Math.Max(0, (height - rowHeight * series.Count) / 2);
        for (var index = 0; index < series.Count; index++)
        {
            var item = series[index];
            var rowY = startY + index * rowHeight;
            var swatch = ShapeFrame(new PpjFrameModel(x, rowY + rowHeight * 0.35, 12, Math.Max(2, rowHeight * 0.3), 0, false, false), item.Type is "scatter" or "bubble" ? "ellipse" : "rect");
            ApplyVectorSeriesFill(swatch, item.Json, catalog, item.Index, elementId, defaultOpacity: item.Type == "area" ? 0.45 : 1);
            group.Children.Add(new PresentationElement
            {
                Id = $"{elementId}/{family}/legend-swatch/{item.Series.Id}",
                Name = $"{family} legend swatch {item.Series.Name}",
                Shape = swatch,
            });
            group.Children.Add(VectorChartTextElement(
                $"{elementId}/{family}/legend-label/{item.Series.Id}",
                $"{family} legend label {item.Series.Name}",
                x + 18,
                rowY,
                width - 18,
                rowHeight,
                item.Series.Name,
                textStyle,
                catalog,
                8,
                "left",
                "52606D"));
        }
    }

    private static string FormatVectorChartValue(double value, string? numberFormat) =>
        value.ToString(numberFormat ?? "0.##", CultureInfo.InvariantCulture);

    private static bool IsStreamgraph(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        return FirstProperty(inlineStyle, namedStyle, "stacking") is { } stacking &&
               stacking.GetString() == "stream";
    }

    private static PresentationGroup BuildStreamgraph(
        PpjChartElementModel element,
        JsonElement raw,
        Catalog catalog)
    {
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ValidateStreamgraphCompileProfile(element, raw, namedStyle, inlineStyle);

        var x = element.Frame.X;
        var y = element.Frame.Y;
        var width = element.Frame.Width;
        var height = element.Frame.Height;
        var titleText = element.Title is null ? string.Empty : Flatten(element.Title);
        var titleHeight = titleText.Length == 0 ? 0 : Math.Clamp(height * 0.12, 22, 38);
        var legend = FirstProperty(inlineStyle, namedStyle, "legend")?.GetString() ?? "right";
        var legendWidth = legend == "right" ? Math.Clamp(width * 0.18, 88, 132) : 0;
        var legendGap = legendWidth > 0 ? 14 : 0;
        var rawXAxis = Property(raw, "xAxis");
        var showCategories = rawXAxis is null || OptionalBoolean(rawXAxis.Value, "visible") is not false;
        var categoryHeight = showCategories ? 18 : 0;
        var plotX = x;
        var plotY = y + titleHeight + 6;
        var plotWidth = width - legendWidth - legendGap;
        var plotHeight = height - titleHeight - categoryHeight - 12;
        if (plotWidth < 220 || plotHeight < 100)
            throw Unsupported(element.Id, "streamgraph frame is too small for editable bands and labels");

        var categories = element.Data.Categories.Select(CategoryText).ToArray();
        var seriesJson = raw.GetProperty("data").GetProperty("series").EnumerateArray().ToArray();
        var totals = new double[categories.Length];
        foreach (var series in element.Data.Series)
            for (var categoryIndex = 0; categoryIndex < categories.Length; categoryIndex++)
                totals[categoryIndex] += series.Values[categoryIndex]!.Value;
        var maximumTotal = totals.Max();
        var scale = plotHeight * 0.86 / maximumTotal;
        var cumulative = new double[categories.Length];
        var upper = new (double X, double Y)[categories.Length];
        var lower = new (double X, double Y)[categories.Length];

        var group = new PresentationGroup
        {
            LeftEmu = Emu(x),
            TopEmu = Emu(y),
            WidthEmu = Emu(width),
            HeightEmu = Emu(height),
            ChildLeftEmu = Emu(x),
            ChildTopEmu = Emu(y),
            ChildWidthEmu = Emu(width),
            ChildHeightEmu = Emu(height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;

        for (var seriesIndex = 0; seriesIndex < element.Data.Series.Count; seriesIndex++)
        {
            var series = element.Data.Series[seriesIndex];
            for (var categoryIndex = 0; categoryIndex < categories.Length; categoryIndex++)
            {
                var pointX = plotX + categoryIndex * plotWidth / (categories.Length - 1);
                var totalHeight = totals[categoryIndex] * scale;
                var bottom = plotY + (plotHeight + totalHeight) / 2;
                lower[categoryIndex] = (pointX, bottom - cumulative[categoryIndex] * scale);
                cumulative[categoryIndex] += series.Values[categoryIndex]!.Value;
                upper[categoryIndex] = (pointX, bottom - cumulative[categoryIndex] * scale);
            }

            var shape = ShapeFrame(new PpjFrameModel(plotX, plotY, plotWidth, plotHeight, 0, false, false), "custom");
            ApplyStreamgraphSeriesPaint(shape, seriesJson[seriesIndex], catalog, seriesIndex, element.Id);
            shape.CustomPaths.Add(BuildStreamgraphBandPath(plotX, plotY, plotWidth, plotHeight, upper, lower));
            group.Children.Add(new PresentationElement
            {
                Id = StreamgraphNativeId(element.Id, $"band/{series.Id}"),
                Name = $"stream band {series.Name}",
                Shape = shape,
            });
        }

        if (titleText.Length > 0)
            group.Children.Add(VectorChartTitleElement(
                StreamgraphNativeId(element.Id, "title"),
                "streamgraph title",
                x,
                y,
                width,
                titleHeight,
                raw.GetProperty("title"),
                FirstProperty(inlineStyle, namedStyle, "titleTextStyle"),
                catalog,
                Math.Clamp(titleHeight * 0.48, 11, 18),
                "left"));

        if (showCategories)
        {
            var labelStyle = Property(rawXAxis, "textStyle");
            var step = Math.Max(1, (int)Math.Ceiling(categories.Length / 8d));
            var labelWidth = Math.Clamp(plotWidth / Math.Min(categories.Length, 8), 42, 92);
            for (var categoryIndex = 0; categoryIndex < categories.Length; categoryIndex++)
            {
                if (categoryIndex != 0 && categoryIndex != categories.Length - 1 && categoryIndex % step != 0) continue;
                var pointX = plotX + categoryIndex * plotWidth / (categories.Length - 1);
                group.Children.Add(VectorChartTextElement(
                    StreamgraphNativeId(element.Id, $"category/{categoryIndex}"),
                    $"stream category {categories[categoryIndex]}",
                    Math.Clamp(pointX - labelWidth / 2, plotX, plotX + plotWidth - labelWidth),
                    plotY + plotHeight + 2,
                    labelWidth,
                    categoryHeight,
                    categories[categoryIndex],
                    labelStyle,
                    catalog,
                    7,
                    categoryIndex == 0 ? "left" : categoryIndex == categories.Length - 1 ? "right" : "center",
                    "52606D"));
            }
        }

        if (legend == "right")
        {
            var legendStyle = FirstProperty(inlineStyle, namedStyle, "legendTextStyle");
            var rowHeight = Math.Min(18, plotHeight / element.Data.Series.Count);
            var legendY = plotY + (plotHeight - rowHeight * element.Data.Series.Count) / 2;
            for (var seriesIndex = 0; seriesIndex < element.Data.Series.Count; seriesIndex++)
            {
                var series = element.Data.Series[seriesIndex];
                var rowY = legendY + seriesIndex * rowHeight;
                var swatch = ShapeFrame(new PpjFrameModel(plotX + plotWidth + legendGap, rowY + rowHeight * 0.38, 12, 3, 0, false, false), "rect");
                ApplyStreamgraphSeriesPaint(swatch, seriesJson[seriesIndex], catalog, seriesIndex, element.Id, includeStroke: false);
                group.Children.Add(new PresentationElement
                {
                    Id = StreamgraphNativeId(element.Id, $"legend-swatch/{series.Id}"),
                    Name = $"stream legend swatch {series.Name}",
                    Shape = swatch,
                });
                group.Children.Add(VectorChartTextElement(
                    StreamgraphNativeId(element.Id, $"legend-label/{series.Id}"),
                    $"stream legend label {series.Name}",
                    plotX + plotWidth + legendGap + 18,
                    rowY,
                    legendWidth - 18,
                    rowHeight,
                    series.Name,
                    legendStyle,
                    catalog,
                    Math.Clamp(rowHeight * 0.48, 6.5, 9),
                    "left",
                    "16324F"));
            }
        }

        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static void ValidateStreamgraphCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        var categories = element.Data.Categories.Select(CategoryText).ToArray();
        if (categories.Length is < 3 or > 64 || categories.Any(string.IsNullOrWhiteSpace) ||
            categories.Distinct(StringComparer.Ordinal).Count() != categories.Length)
            throw Unsupported(element.Id, "streamgraph categories must be 3..64 unique non-empty labels");
        if (element.Data.Series.Count is < 2 or > 12)
            throw Unsupported(element.Id, "streamgraphs require 2..12 series");
        if (element.Data.Series.Any(series => string.IsNullOrWhiteSpace(series.Name)) ||
            element.Data.Series.Select(series => series.Name).Distinct(StringComparer.Ordinal).Count() != element.Data.Series.Count)
            throw Unsupported(element.Id, "streamgraph series names must be unique and non-empty");
        if (element.Data.Series.Any(series => series.Values.Count != categories.Length || series.Values.Any(value => value is null or < 0)))
            throw Unsupported(element.Id, "streamgraph series require complete aligned non-negative values");
        for (var categoryIndex = 0; categoryIndex < categories.Length; categoryIndex++)
            if (element.Data.Series.Sum(series => series.Values[categoryIndex]!.Value) <= 0)
                throw Unsupported(element.Id, $"streamgraph category {categories[categoryIndex]} has no positive magnitude");

        var seriesJson = raw.GetProperty("data").GetProperty("series").EnumerateArray().ToArray();
        for (var seriesIndex = 0; seriesIndex < seriesJson.Length; seriesIndex++)
        {
            var series = seriesJson[seriesIndex];
            RejectProperties(series, element.Id, "pointRoles", "xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "parents", "sources", "targets", "chartType", "axis", "marker", "trendlines", "errorBars");
            if (series.TryGetProperty("color", out _) && series.TryGetProperty("fill", out _))
                throw Unsupported(element.Id, "streamgraph series color and fill are aliases and cannot both be present");
            if (series.TryGetProperty("fill", out var fill) && fill.GetProperty("type").GetString() is not ("solid" or "gradient"))
                throw Unsupported(element.Id, "streamgraph series fill must be solid or gradient");
        }
        foreach (var property in new[] { "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, "streamgraphs use one generated centered value scale and do not accept Y or secondary axes");
        if (raw.TryGetProperty("xAxis", out var xAxis))
            foreach (var property in xAxis.EnumerateObject())
                if (property.Name is not ("visible" or "textStyle"))
                    throw Unsupported(element.Id, $"streamgraph xAxis does not support {property.Name}");

        foreach (var style in new[] { namedStyle, inlineStyle })
        {
            if (style is not { ValueKind: JsonValueKind.Object } value) continue;
            foreach (var property in value.EnumerateObject())
                if (property.Name is not ("stacking" or "legend" or "titleTextStyle" or "legendTextStyle"))
                    throw Unsupported(element.Id, $"streamgraphs do not support chart style field {property.Name}");
        }
        if (FirstProperty(inlineStyle, namedStyle, "stacking") is not { } stacking || stacking.GetString() != "stream")
            throw Unsupported(element.Id, "streamgraphs require style.stacking stream");
        if (FirstProperty(inlineStyle, namedStyle, "legend") is { } legend && legend.GetString() is not ("none" or "right"))
            throw Unsupported(element.Id, "streamgraph legend supports only none or right");
    }

    private static void ApplyStreamgraphSeriesPaint(
        PresentationShape target,
        JsonElement series,
        Catalog catalog,
        int seriesIndex,
        string elementId,
        bool includeStroke = true)
    {
        if (series.TryGetProperty("fill", out var fill))
        {
            if (fill.GetProperty("type").GetString() == "gradient")
                target.GradientFill = BuildGradientFill(fill, color => catalog.Color(color));
            else
            {
                var resolved = FillColor(fill, catalog) ??
                    throw Unsupported(elementId, "streamgraph series cannot use a no-fill paint");
                target.FillRgb = resolved.Rgb;
                if (resolved.Opacity < 1) target.FillOpacityThousandthPercent = Opacity(resolved.Opacity);
            }
        }
        else if (series.TryGetProperty("color", out var color))
        {
            var resolved = catalog.Color(color);
            target.FillRgb = resolved.Rgb;
            if (resolved.Alpha < 1) target.FillOpacityThousandthPercent = Opacity(resolved.Alpha);
        }
        else target.FillRgb = catalog.Theme.AccentRgb[seriesIndex % catalog.Theme.AccentRgb.Count];

        if (includeStroke && series.TryGetProperty("stroke", out var stroke)) ApplyLine(target, stroke, catalog);
        else target.LineStyle = "none";
    }

    private static PresentationCustomGeometryPath BuildStreamgraphBandPath(
        double plotX,
        double plotY,
        double plotWidth,
        double plotHeight,
        IReadOnlyList<(double X, double Y)> upper,
        IReadOnlyList<(double X, double Y)> lower)
    {
        const long viewport = 100_000;
        PresentationCustomGeometryPoint Point((double X, double Y) source) => new()
        {
            X = checked((long)Math.Round((source.X - plotX) / plotWidth * viewport, MidpointRounding.AwayFromZero)),
            Y = checked((long)Math.Round((source.Y - plotY) / plotHeight * viewport, MidpointRounding.AwayFromZero)),
        };
        PresentationCustomGeometryCommand Curve((double X, double Y) start, (double X, double Y) end) => new()
        {
            CubicBezierTo = new PresentationCustomGeometryCubicBezier
            {
                Control1 = Point((start.X + (end.X - start.X) / 3, start.Y)),
                Control2 = Point((end.X - (end.X - start.X) / 3, end.Y)),
                End = Point(end),
            },
        };

        var path = new PresentationCustomGeometryPath
        {
            Width = viewport,
            Height = viewport,
            FillMode = PresentationCustomGeometryPath.Types.FillMode.Normal,
            Stroke = true,
        };
        path.Commands.Add(MoveTo(Point(upper[0])));
        for (var index = 0; index < upper.Count - 1; index++) path.Commands.Add(Curve(upper[index], upper[index + 1]));
        path.Commands.Add(LineTo(Point(lower[^1])));
        for (var index = lower.Count - 1; index > 0; index--) path.Commands.Add(Curve(lower[index], lower[index - 1]));
        path.Commands.Add(new PresentationCustomGeometryCommand { Close = true });
        return path;
    }

    private static string StreamgraphNativeId(string elementId, string suffix) =>
        $"{elementId}/stream/{suffix}";

    private static bool IsPictographicChart(PpjChartElementModel element) =>
        element.ChartType is "bar" or "column" &&
        element.Data.Series.Any(series => series.Raw.TryGetProperty("symbol", out _));

    private static PresentationGroup BuildPictographicChart(
        PpjChartElementModel element,
        JsonElement raw,
        Catalog catalog)
    {
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ValidatePictographicCompileProfile(element, raw, namedStyle, inlineStyle);

        var x = element.Frame.X;
        var y = element.Frame.Y;
        var width = element.Frame.Width;
        var height = element.Frame.Height;
        var titleText = element.Title is null ? string.Empty : Flatten(element.Title);
        var titleHeight = titleText.Length == 0 ? 0 : Math.Clamp(height * 0.12, 22, 38);
        const double unitCaptionHeight = 16;
        var contentY = y + titleHeight + unitCaptionHeight + 8;
        var contentHeight = height - titleHeight - unitCaptionHeight - 12;
        if (contentHeight < 76 || width < 220)
            throw Unsupported(element.Id, "pictographic frame is too small for editable symbols and labels");

        var series = element.Data.Series[0];
        var seriesJson = raw.GetProperty("data").GetProperty("series")[0];
        var symbol = seriesJson.GetProperty("symbol");
        var unit = symbol.GetProperty("unit").GetDouble();
        var gap = OptionalDouble(symbol, "gap") ?? 2;
        var showValue = OptionalBoolean(symbol, "showValue") is not false;
        var unitLabel = OptionalString(symbol, "unitLabel");
        var categories = element.Data.Categories.Select(CategoryText).ToArray();
        var counts = series.Values.Select(value => checked((int)Math.Round(value!.Value / unit))).ToArray();
        var maximumCount = counts.Max();

        var group = new PresentationGroup
        {
            LeftEmu = Emu(x),
            TopEmu = Emu(y),
            WidthEmu = Emu(width),
            HeightEmu = Emu(height),
            ChildLeftEmu = Emu(x),
            ChildTopEmu = Emu(y),
            ChildWidthEmu = Emu(width),
            ChildHeightEmu = Emu(height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;

        if (titleText.Length > 0)
            group.Children.Add(VectorChartTitleElement(
                PictographNativeId(element.Id, "title"),
                "pictographic chart title",
                x,
                y,
                width,
                titleHeight,
                raw.GetProperty("title"),
                FirstProperty(inlineStyle, namedStyle, "titleTextStyle"),
                catalog,
                Math.Clamp(titleHeight * 0.48, 11, 18),
                "left"));

        var unitText = $"1 symbol = {FormatPictographValue(unit)}{(unitLabel is null ? string.Empty : " " + unitLabel)}";
        group.Children.Add(VectorChartTextElement(
            PictographNativeId(element.Id, "unit"),
            "pictographic unit",
            x,
            y + titleHeight,
            width,
            unitCaptionHeight,
            unitText,
            null,
            catalog,
            7.5,
            "right",
            "52606D"));

        if (element.ChartType == "bar")
            BuildHorizontalPictograph(
                group, element, symbol, seriesJson, catalog, categories, counts,
                x, contentY, width, contentHeight, maximumCount, gap, showValue, unitLabel);
        else
            BuildVerticalPictograph(
                group, element, symbol, seriesJson, catalog, categories, counts,
                x, contentY, width, contentHeight, maximumCount, gap, showValue, unitLabel);

        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static void BuildHorizontalPictograph(
        PresentationGroup group,
        PpjChartElementModel element,
        JsonElement symbol,
        JsonElement series,
        Catalog catalog,
        IReadOnlyList<string> categories,
        IReadOnlyList<int> counts,
        double x,
        double y,
        double width,
        double height,
        int maximumCount,
        double gap,
        bool showValue,
        string? unitLabel)
    {
        var categoryWidth = Math.Clamp(width * 0.2, 70, 130);
        var valueWidth = showValue ? Math.Clamp(width * 0.12, 48, 86) : 0;
        const double plotGap = 10;
        var plotX = x + categoryWidth + plotGap;
        var plotWidth = width - categoryWidth - valueWidth - plotGap * 2;
        var rowHeight = height / categories.Count;
        var symbolSize = Math.Min(rowHeight * 0.62, (plotWidth - gap * (maximumCount - 1)) / maximumCount);
        if (symbolSize < 6)
            throw Unsupported(element.Id, "pictographic bar symbols would be smaller than 6 points");

        for (var categoryIndex = 0; categoryIndex < categories.Count; categoryIndex++)
        {
            var rowY = y + categoryIndex * rowHeight;
            group.Children.Add(VectorChartTextElement(
                PictographNativeId(element.Id, $"category/{categoryIndex}"),
                $"pictographic category {categories[categoryIndex]}",
                x,
                rowY,
                categoryWidth,
                rowHeight,
                categories[categoryIndex],
                null,
                catalog,
                8,
                "left",
                "16324F"));
            for (var symbolIndex = 0; symbolIndex < counts[categoryIndex]; symbolIndex++)
            {
                var symbolX = plotX + symbolIndex * (symbolSize + gap);
                var symbolY = rowY + (rowHeight - symbolSize) / 2;
                group.Children.Add(PictographSymbolElement(
                    element.Id, symbol, series, catalog, categoryIndex, symbolIndex,
                    new PpjFrameModel(symbolX, symbolY, symbolSize, symbolSize, 0, false, false)));
            }
            if (showValue)
                group.Children.Add(VectorChartTextElement(
                    PictographNativeId(element.Id, $"value/{categoryIndex}"),
                    $"pictographic value {categories[categoryIndex]}",
                    x + width - valueWidth,
                    rowY,
                    valueWidth,
                    rowHeight,
                    PictographValueLabel(element.Data.Series[0].Values[categoryIndex]!.Value, unitLabel),
                    null,
                    catalog,
                    8,
                    "right",
                    "16324F"));
        }
    }

    private static void BuildVerticalPictograph(
        PresentationGroup group,
        PpjChartElementModel element,
        JsonElement symbol,
        JsonElement series,
        Catalog catalog,
        IReadOnlyList<string> categories,
        IReadOnlyList<int> counts,
        double x,
        double y,
        double width,
        double height,
        int maximumCount,
        double gap,
        bool showValue,
        string? unitLabel)
    {
        const double categoryHeight = 20;
        var valueHeight = showValue ? 18 : 0;
        var plotY = y + valueHeight;
        var plotHeight = height - categoryHeight - valueHeight;
        var columnWidth = width / categories.Count;
        var symbolSize = Math.Min(columnWidth * 0.62, (plotHeight - gap * (maximumCount - 1)) / maximumCount);
        if (symbolSize < 6)
            throw Unsupported(element.Id, "pictographic column symbols would be smaller than 6 points");
        var baseline = plotY + plotHeight;

        for (var categoryIndex = 0; categoryIndex < categories.Count; categoryIndex++)
        {
            var symbolX = x + categoryIndex * columnWidth + (columnWidth - symbolSize) / 2;
            for (var symbolIndex = 0; symbolIndex < counts[categoryIndex]; symbolIndex++)
            {
                var symbolY = baseline - (symbolIndex + 1) * symbolSize - symbolIndex * gap;
                group.Children.Add(PictographSymbolElement(
                    element.Id, symbol, series, catalog, categoryIndex, symbolIndex,
                    new PpjFrameModel(symbolX, symbolY, symbolSize, symbolSize, 0, false, false)));
            }
            if (showValue)
            {
                var stackTop = baseline - counts[categoryIndex] * symbolSize - Math.Max(0, counts[categoryIndex] - 1) * gap;
                group.Children.Add(VectorChartTextElement(
                    PictographNativeId(element.Id, $"value/{categoryIndex}"),
                    $"pictographic value {categories[categoryIndex]}",
                    x + categoryIndex * columnWidth,
                    Math.Max(y, stackTop - valueHeight),
                    columnWidth,
                    valueHeight,
                    PictographValueLabel(element.Data.Series[0].Values[categoryIndex]!.Value, unitLabel),
                    null,
                    catalog,
                    8,
                    "center",
                    "16324F"));
            }
            group.Children.Add(VectorChartTextElement(
                PictographNativeId(element.Id, $"category/{categoryIndex}"),
                $"pictographic category {categories[categoryIndex]}",
                x + categoryIndex * columnWidth,
                baseline + 2,
                columnWidth,
                categoryHeight,
                categories[categoryIndex],
                null,
                catalog,
                8,
                "center",
                "16324F"));
        }
    }

    private static PresentationElement PictographSymbolElement(
        string elementId,
        JsonElement symbol,
        JsonElement series,
        Catalog catalog,
        int categoryIndex,
        int symbolIndex,
        PpjFrameModel frame)
    {
        var kind = symbol.GetProperty("kind").GetString();
        PresentationShape shape;
        if (kind == "icon")
        {
            shape = ShapeFrame(frame, "custom");
            shape.CatalogIconName = symbol.GetProperty("iconName").GetString()!;
            ApplyIconGeometry(shape, PpjIconCatalog.Resolve(shape.CatalogIconName), frame, elementId);
        }
        else
        {
            var preset = symbol.GetProperty("preset").GetString()!;
            PptxPresetGeometryAdjustmentCodec.Validate(preset, [], elementId);
            shape = ShapeFrame(frame, preset);
        }
        ApplyPictographPaint(shape, series, catalog, elementId);
        return new PresentationElement
        {
            Id = PictographNativeId(elementId, $"category/{categoryIndex}/symbol/{symbolIndex}"),
            Name = $"pictographic symbol {categoryIndex + 1}.{symbolIndex + 1}",
            Shape = shape,
        };
    }

    private static void ApplyPictographPaint(
        PresentationShape target,
        JsonElement series,
        Catalog catalog,
        string elementId)
    {
        if (series.TryGetProperty("fill", out var fill))
        {
            if (fill.GetProperty("type").GetString() == "gradient")
                target.GradientFill = BuildGradientFill(fill, color => catalog.Color(color));
            else
            {
                var resolved = FillColor(fill, catalog) ??
                    throw Unsupported(elementId, "pictographic symbols cannot use a no-fill paint");
                target.FillRgb = resolved.Rgb;
                if (resolved.Opacity < 1) target.FillOpacityThousandthPercent = Opacity(resolved.Opacity);
            }
        }
        else if (series.TryGetProperty("color", out var color))
        {
            var resolved = catalog.Color(color);
            target.FillRgb = resolved.Rgb;
            if (resolved.Alpha < 1) target.FillOpacityThousandthPercent = Opacity(resolved.Alpha);
        }
        else target.FillRgb = catalog.Theme.AccentRgb[0];

        if (series.TryGetProperty("stroke", out var stroke)) ApplyLine(target, stroke, catalog);
        else target.LineStyle = "none";
    }

    private static void ValidatePictographicCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        if (element.ChartType is not ("bar" or "column") || element.Data.Series.Count != 1 ||
            !element.Data.Series[0].Raw.TryGetProperty("symbol", out var symbol))
            throw Unsupported(element.Id, "pictographic charts require one bar or column series with a symbol");
        if (element.Data.Categories.Count is < 2 or > 12 ||
            element.Data.Categories.Any(category => category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString())) ||
            element.Data.Categories.Select(category => category.GetString()).Distinct(StringComparer.Ordinal).Count() != element.Data.Categories.Count)
            throw Unsupported(element.Id, "pictographic categories must be 2..12 unique non-empty labels");
        var kind = symbol.GetProperty("kind").GetString();
        if (kind == "icon") _ = PpjIconCatalog.Resolve(symbol.GetProperty("iconName").GetString()!);
        else
        {
            var preset = symbol.GetProperty("preset").GetString()!;
            if (preset is "textbox" or "line" || !PptxPresetGeometryAdjustmentCodec.HasProfile(preset))
                throw Unsupported(element.Id, $"pictographic preset {preset} is not a closed authored preset");
        }
        if (symbol.TryGetProperty("unitLabel", out var unitLabel) && string.IsNullOrWhiteSpace(unitLabel.GetString()))
            throw Unsupported(element.Id, "pictographic unitLabel must be non-empty");
        var unit = symbol.GetProperty("unit").GetDouble();
        var total = 0;
        foreach (var value in element.Data.Series[0].Values)
        {
            if (value is null || !double.IsFinite(value.Value) || value.Value < 0)
                throw Unsupported(element.Id, "pictographic values must be finite, complete and non-negative");
            var quotient = value.Value / unit;
            var rounded = Math.Round(quotient);
            if (Math.Abs(quotient - rounded) > 1e-9 * Math.Max(1, Math.Abs(quotient)))
                throw Unsupported(element.Id, $"pictographic value {value.Value} is not an exact multiple of unit {unit}");
            if (rounded > 32) throw Unsupported(element.Id, "one pictographic category exceeds the 32-symbol budget");
            total = checked(total + (int)rounded);
        }
        if (total is < 1 or > 192)
            throw Unsupported(element.Id, "pictographic charts require 1..192 total symbols");

        foreach (var property in element.Data.Series[0].Raw.EnumerateObject())
            if (property.Name is not ("id" or "name" or "values" or "color" or "fill" or "stroke" or "symbol"))
                throw Unsupported(element.Id, $"pictographic series do not support {property.Name}");
        var series = element.Data.Series[0].Raw;
        if (series.TryGetProperty("color", out _) && series.TryGetProperty("fill", out _))
            throw Unsupported(element.Id, "pictographic series color and fill are aliases and cannot both be present");
        if (series.TryGetProperty("fill", out var fill) && fill.GetProperty("type").GetString() is not ("solid" or "gradient"))
            throw Unsupported(element.Id, "pictographic symbol fill must be solid or gradient");
        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, "pictographic charts generate their own categorical layout and do not accept axes");
        foreach (var style in new[] { namedStyle, inlineStyle })
        {
            if (style is not { ValueKind: JsonValueKind.Object } value) continue;
            foreach (var property in value.EnumerateObject())
                if (property.Name != "titleTextStyle")
                    throw Unsupported(element.Id, $"pictographic charts do not support chart style field {property.Name}");
        }
    }

    private static string PictographValueLabel(double value, string? unitLabel) =>
        $"{FormatPictographValue(value)}{(unitLabel is null ? string.Empty : " " + unitLabel)}";

    private static string FormatPictographValue(double value) =>
        value.ToString("0.###", CultureInfo.InvariantCulture);

    private static string PictographNativeId(string elementId, string suffix) =>
        $"{elementId}/pictograph/{suffix}";

    internal static PresentationTextBody BuildTitleBody(
        PpjProgramModel program,
        PpjChartElementModel element)
    {
        if (!element.Raw.TryGetProperty("title", out var title))
            throw Unsupported(element.Id, "structured chart title is missing");
        var catalog = new Catalog(program.Root);
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(element.Raw, "style");
        var body = BuildTextBody(title, null, null, catalog);
        if (FirstProperty(inlineStyle, namedStyle, "titleTextStyle") is { } titleStyle)
            ApplyChartTitleDefaults(body, BuildChartTextStyle(titleStyle, catalog));
        return body;
    }

    private static PresentationGroup BuildHeatmap(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ValidateHeatmapCompileProfile(element, raw, namedStyle, inlineStyle);
        var style = FirstProperty(inlineStyle, namedStyle, "heatmap")!.Value;
        var scale = OptionalString(style, "scale") ?? "linear";
        var colors = style.GetProperty("colors").EnumerateArray()
            .Select(color => HeatmapColor(catalog.Color(color)))
            .ToArray();
        var values = element.Data.Series.SelectMany(series => series.Values)
            .Where(value => value is not null)
            .Select(value => value!.Value)
            .ToArray();
        var inferredMinimum = values.Min();
        var inferredMaximum = values.Max();
        double minimum;
        double maximum;
        if (style.TryGetProperty("domain", out var domain))
        {
            minimum = domain[0].GetDouble();
            maximum = domain[1].GetDouble();
        }
        else if (scale == "diverging")
        {
            var extent = Math.Max(Math.Abs(inferredMinimum), Math.Abs(inferredMaximum));
            if (extent == 0) extent = 1;
            minimum = -extent;
            maximum = extent;
        }
        else if (inferredMinimum == inferredMaximum)
        {
            var extent = Math.Max(1, Math.Abs(inferredMinimum) * 0.05);
            minimum = inferredMinimum - extent;
            maximum = inferredMaximum + extent;
        }
        else
        {
            minimum = inferredMinimum;
            maximum = inferredMaximum;
        }
        var midpoint = OptionalDouble(style, "midpoint") ?? 0;

        var showValues = OptionalBoolean(style, "showValues") ?? false;
        var showColorBar = OptionalBoolean(style, "showColorBar") ?? true;
        var cellGap = OptionalDouble(style, "cellGap") ?? 1.5;
        var axisStyle = Property(style, "axisTextStyle");
        var valueStyle = Property(style, "valueTextStyle");
        var titleStyle = FirstProperty(inlineStyle, namedStyle, "titleTextStyle");

        var x = element.Frame.X;
        var y = element.Frame.Y;
        var width = element.Frame.Width;
        var height = element.Frame.Height;
        var titleText = element.Title is null ? string.Empty : Flatten(element.Title);
        var titleHeight = titleText.Length == 0 ? 0 : Math.Clamp(height * 0.11, 20, 36);
        var longestRowLabel = element.Data.Series.Max(series => series.Name.Length);
        var maximumLabelWidth = Math.Max(52, Math.Min(150, width * 0.24));
        var leftLabelWidth = Math.Clamp(24 + longestRowLabel * 4.8, 52, maximumLabelWidth);
        var bottomLabelHeight = Math.Clamp(height * 0.09, 20, 36);
        var colorBarWidth = showColorBar ? Math.Clamp(width * 0.09, 46, 72) : 0;
        var gridX = x + leftLabelWidth;
        var gridY = y + titleHeight;
        var gridWidth = width - leftLabelWidth - colorBarWidth;
        var gridHeight = height - titleHeight - bottomLabelHeight;
        var columnCount = element.Data.Categories.Count;
        var rowCount = element.Data.Series.Count;
        var cellWidth = gridWidth / columnCount;
        var cellHeight = gridHeight / rowCount;
        if (cellWidth < 8 || cellHeight < 8 || cellGap >= cellWidth || cellGap >= cellHeight)
            throw Unsupported(element.Id, "vector heatmap frame is too small for its matrix and cell gap");

        var group = new PresentationGroup
        {
            LeftEmu = Emu(x),
            TopEmu = Emu(y),
            WidthEmu = Emu(width),
            HeightEmu = Emu(height),
            ChildLeftEmu = Emu(x),
            ChildTopEmu = Emu(y),
            ChildWidthEmu = Emu(width),
            ChildHeightEmu = Emu(height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;

        if (titleText.Length > 0)
            group.Children.Add(VectorChartTitleElement(
                HeatmapNativeId(element.Id, "title"),
                "heatmap title",
                x,
                y,
                width,
                titleHeight,
                raw.GetProperty("title"),
                titleStyle,
                catalog,
                Math.Clamp(titleHeight * 0.48, 11, 18),
                "left"));

        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            group.Children.Add(VectorChartTextElement(
                HeatmapNativeId(element.Id, $"row/{rowIndex}"),
                $"heatmap row {rowIndex + 1}",
                x,
                gridY + rowIndex * cellHeight,
                leftLabelWidth - 6,
                cellHeight,
                element.Data.Series[rowIndex].Name,
                axisStyle,
                catalog,
                Math.Clamp(cellHeight * 0.34, 6, 10),
                "right"));

            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                var value = element.Data.Series[rowIndex].Values[columnIndex];
                var cellX = gridX + columnIndex * cellWidth + cellGap / 2;
                var cellY = gridY + rowIndex * cellHeight + cellGap / 2;
                var cellShape = ShapeFrame(
                    new PpjFrameModel(cellX, cellY, cellWidth - cellGap, cellHeight - cellGap, 0, false, false),
                    "rect");
                cellShape.LineStyle = "none";
                (byte R, byte G, byte B, double Alpha)? cellColor = null;
                if (value is not null)
                {
                    cellColor = HeatmapInterpolate(value.Value, minimum, maximum, midpoint, scale, colors);
                    cellShape.FillRgb = HeatmapHex(cellColor.Value);
                    if (cellColor.Value.Alpha < 1) cellShape.FillOpacityThousandthPercent = Opacity(cellColor.Value.Alpha);
                }
                else if (style.TryGetProperty("missingFill", out var missingFill))
                {
                    cellColor = HeatmapColor(catalog.Color(missingFill));
                    cellShape.FillRgb = HeatmapHex(cellColor.Value);
                    if (cellColor.Value.Alpha < 1) cellShape.FillOpacityThousandthPercent = Opacity(cellColor.Value.Alpha);
                }
                if (style.TryGetProperty("cellStroke", out var cellStroke)) ApplyLine(cellShape, cellStroke, catalog);
                group.Children.Add(new PresentationElement
                {
                    Id = HeatmapNativeId(element.Id, $"cell/{rowIndex}/{columnIndex}"),
                    Name = $"heatmap cell {rowIndex + 1},{columnIndex + 1}",
                    Shape = cellShape,
                });

                if (showValues && value is not null)
                {
                    var contrast = HeatmapLuminance(cellColor!.Value) > 0.52 ? "111827" : "FFFFFF";
                    group.Children.Add(VectorChartTextElement(
                        HeatmapNativeId(element.Id, $"value/{rowIndex}/{columnIndex}"),
                        $"heatmap value {rowIndex + 1},{columnIndex + 1}",
                        cellX + 1,
                        cellY + 1,
                        cellWidth - cellGap - 2,
                        cellHeight - cellGap - 2,
                        value.Value.ToString("0.##", CultureInfo.InvariantCulture),
                        valueStyle,
                        catalog,
                        Math.Clamp(Math.Min(cellWidth, cellHeight) * 0.28, 6, 10),
                        "center",
                        contrast));
                }
            }
        }

        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            group.Children.Add(VectorChartTextElement(
                HeatmapNativeId(element.Id, $"column/{columnIndex}"),
                $"heatmap column {columnIndex + 1}",
                gridX + columnIndex * cellWidth,
                gridY + gridHeight + 3,
                cellWidth,
                bottomLabelHeight - 3,
                element.Data.Categories[columnIndex].GetString()!,
                axisStyle,
                catalog,
                Math.Clamp(bottomLabelHeight * 0.3, 6, 9),
                "center"));

        if (showColorBar)
            AddHeatmapColorBar(group, element.Id, gridX + gridWidth + 15, gridY, colorBarWidth - 15, gridHeight, minimum, maximum, midpoint, scale, colors, axisStyle, catalog);

        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static void ValidateHeatmapCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        if (element.Data.Categories.Count is < 1 or > 32 || element.Data.Series.Count is < 1 or > 32)
            throw Unsupported(element.Id, "vector heatmaps support a 1..32 by 1..32 matrix");
        if (element.Data.Categories.Any(category => category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString())) ||
            element.Data.Categories.Select(category => category.GetString()).Distinct(StringComparer.Ordinal).Count() != element.Data.Categories.Count)
            throw Unsupported(element.Id, "vector heatmap categories must be unique non-empty strings");
        if (element.Data.Series.Any(series => string.IsNullOrWhiteSpace(series.Name)) ||
            element.Data.Series.Select(series => series.Name).Distinct(StringComparer.Ordinal).Count() != element.Data.Series.Count)
            throw Unsupported(element.Id, "vector heatmap series names must be unique and non-empty");
        if (element.Data.Series.Any(series => series.Values.Count != element.Data.Categories.Count) ||
            element.Data.Series.SelectMany(series => series.Values).All(value => value is null))
            throw Unsupported(element.Id, "vector heatmap rows must match the category count and contain at least one numeric value");
        foreach (var series in element.Data.Series)
            foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
                if (series.Raw.TryGetProperty(property, out _))
                    throw Unsupported(element.Id, $"vector heatmap series do not support {property}");
        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (raw.TryGetProperty(property, out _)) throw Unsupported(element.Id, "vector heatmaps generate their own matrix labels");
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "startAngle", "holeSize", "bubbleScale", "bubbleSizeMode", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall" })
            if (FirstProperty(inlineStyle, namedStyle, property) is not null)
                throw Unsupported(element.Id, $"vector heatmaps do not support chart style field {property}");
        if (FirstProperty(inlineStyle, namedStyle, "heatmap") is not { } heatmap)
            throw Unsupported(element.Id, "vector heatmaps require style.heatmap");
        var scale = OptionalString(heatmap, "scale") ?? "linear";
        var colorCount = heatmap.GetProperty("colors").GetArrayLength();
        if ((scale == "linear" && colorCount != 2) || (scale == "diverging" && colorCount != 3))
            throw Unsupported(element.Id, scale == "diverging" ? "diverging heatmaps require three colors" : "linear heatmaps require two colors");
        if (heatmap.TryGetProperty("domain", out var domain))
        {
            if (domain[0].GetDouble() >= domain[1].GetDouble())
                throw Unsupported(element.Id, "vector heatmap domain minimum must be smaller than its maximum");
            var effectiveMidpoint = OptionalDouble(heatmap, "midpoint") ?? 0;
            if (scale == "diverging" && (effectiveMidpoint <= domain[0].GetDouble() || effectiveMidpoint >= domain[1].GetDouble()))
                throw Unsupported(element.Id, "diverging heatmap midpoint must lie inside its explicit domain");
        }
        if (heatmap.TryGetProperty("midpoint", out var midpoint))
        {
            if (scale != "diverging") throw Unsupported(element.Id, "heatmap midpoint requires a diverging scale");
        }
    }

    private static PresentationElement VectorChartTextElement(
        string id,
        string name,
        double x,
        double y,
        double width,
        double height,
        string text,
        JsonElement? style,
        Catalog catalog,
        double defaultFontSize,
        string alignment,
        string? fallbackColor = null)
    {
        var shape = ShapeFrame(new PpjFrameModel(x, y, width, height, 0, false, false), "textbox");
        shape.LineStyle = "none";
        shape.Text = text;
        shape.TextBody = new PresentationTextBody
        {
            BodyProperties = new PresentationTextBodyProperties
            {
                NoLeftInset = true,
                NoTopInset = true,
                NoRightInset = true,
                NoBottomInset = true,
                VerticalAnchor = "center",
                Wrap = "none",
                AutoFitMode = "shrinkText",
            },
        };
        var paragraph = new PresentationTextParagraph { Alignment = alignment };
        paragraph.Runs.Add(BuildVectorChartRun(text, style, catalog, defaultFontSize, fallbackColor));
        shape.TextBody.Paragraphs.Add(paragraph);
        return new PresentationElement { Id = id, Name = name, Shape = shape };
    }

    private static PresentationElement VectorChartTitleElement(
        string id,
        string name,
        double x,
        double y,
        double width,
        double height,
        JsonElement title,
        JsonElement? style,
        Catalog catalog,
        double defaultFontSize,
        string alignment,
        string? fallbackColor = null)
    {
        var body = BuildTextBody(title, null, null, catalog);
        var defaults = style is { } value ? BuildChartTextStyle(value, catalog) : null;
        ApplyChartTitleDefaults(body, defaults);
        foreach (var paragraph in body.Paragraphs)
        {
            if (!paragraph.HasAlignment) paragraph.Alignment = alignment;
            foreach (var run in paragraph.Runs)
            {
                if (!run.HasFontSizePoints) run.FontSizePoints = defaultFontSize;
                if (!run.HasColorRgb && !run.HasColorScheme && run.GradientFill is null && fallbackColor is not null)
                    run.ColorRgb = fallbackColor;
            }
        }
        body.BodyProperties = new PresentationTextBodyProperties
        {
            NoLeftInset = true,
            NoTopInset = true,
            NoRightInset = true,
            NoBottomInset = true,
            VerticalAnchor = "center",
            Wrap = "none",
            AutoFitMode = "shrinkText",
        };
        var shape = ShapeFrame(new PpjFrameModel(x, y, width, height, 0, false, false), "textbox");
        shape.LineStyle = "none";
        shape.TextBody = body;
        shape.Text = Flatten(body);
        return new PresentationElement { Id = id, Name = name, Shape = shape };
    }

    private static void ApplyChartTitleDefaults(
        PresentationTextBody body,
        SpreadsheetChartTextStyleArtifact? defaults)
    {
        if (defaults is null) return;
        foreach (var run in body.Paragraphs.SelectMany(paragraph => paragraph.Runs))
        {
            if (!run.HasFontSizePoints && defaults.HasFontSizePoints) run.FontSizePoints = defaults.FontSizePoints;
            if (!run.HasFontFamily && defaults.FontFamily.Length > 0) run.FontFamily = defaults.FontFamily;
            if (!run.HasFontFamilyEastAsia && defaults.FontFamilyEastAsia.Length > 0) run.FontFamilyEastAsia = defaults.FontFamilyEastAsia;
            if (!run.HasFontFamilyComplexScript && defaults.FontFamilyComplexScript.Length > 0) run.FontFamilyComplexScript = defaults.FontFamilyComplexScript;
            if (!run.HasBold && defaults.HasBold) run.Bold = defaults.Bold;
            if (!run.HasItalic && defaults.HasItalic) run.Italic = defaults.Italic;
            if (!run.HasColorRgb && !run.HasColorScheme && run.GradientFill is null && defaults.ColorRgb.Length > 0)
            {
                run.ColorRgb = defaults.ColorRgb;
                if (defaults.HasOpacityThousandthPercent)
                    run.ColorOpacityThousandthPercent = defaults.OpacityThousandthPercent;
            }
        }
    }

    private static PresentationTextRun BuildVectorChartRun(
        string text,
        JsonElement? style,
        Catalog catalog,
        double defaultFontSize,
        string? fallbackColor)
    {
        var run = new PresentationTextRun { Text = text, FontSizePoints = defaultFontSize };
        if (style is not { } value)
        {
            if (fallbackColor is not null) run.ColorRgb = fallbackColor;
            return run;
        }
        if (value.TryGetProperty("fontSize", out var fontSize)) run.FontSizePoints = fontSize.GetDouble();
        if (value.TryGetProperty("fontFamily", out var fontFamily)) run.FontFamily = fontFamily.GetString()!;
        if (value.TryGetProperty("fontFamilyEastAsia", out var eastAsia)) run.FontFamilyEastAsia = eastAsia.GetString()!;
        else if (run.HasFontFamily) run.FontFamilyEastAsia = run.FontFamily;
        if (value.TryGetProperty("fontFamilyComplexScript", out var complexScript)) run.FontFamilyComplexScript = complexScript.GetString()!;
        if (value.TryGetProperty("bold", out var bold)) run.Bold = bold.GetBoolean();
        if (value.TryGetProperty("italic", out var italic)) run.Italic = italic.GetBoolean();
        if (value.TryGetProperty("color", out var color))
        {
            var resolved = catalog.Color(color);
            run.ColorRgb = resolved.Rgb;
            if (resolved.Alpha < 1) run.ColorOpacityThousandthPercent = Opacity(resolved.Alpha);
        }
        else if (fallbackColor is not null) run.ColorRgb = fallbackColor;
        return run;
    }

    private static void AddHeatmapColorBar(
        PresentationGroup group,
        string elementId,
        double x,
        double y,
        double width,
        double height,
        double minimum,
        double maximum,
        double midpoint,
        string scale,
        IReadOnlyList<(byte R, byte G, byte B, double Alpha)> colors,
        JsonElement? axisStyle,
        Catalog catalog)
    {
        var barWidth = Math.Min(12, Math.Max(6, width * 0.32));
        var bar = ShapeFrame(new PpjFrameModel(x, y, barWidth, height, 0, false, false), "rect");
        bar.LineStyle = "none";
        bar.GradientFill = new PresentationGradientFill
        {
            Kind = PresentationGradientFill.Types.Kind.Linear,
            Angle60000 = Angle(90),
        };
        for (var index = 0; index < colors.Count; index++)
        {
            var color = colors[index];
            var stop = new PresentationGradientStop
            {
                PositionThousandthPercent = colors.Count == 1 ? 0U : checked((uint)Math.Round(index * 100_000d / (colors.Count - 1))),
                ColorRgb = HeatmapHex(color),
            };
            if (color.Alpha < 1) stop.OpacityThousandthPercent = Opacity(color.Alpha);
            bar.GradientFill.Stops.Add(stop);
        }
        group.Children.Add(new PresentationElement { Id = HeatmapNativeId(elementId, "colorbar"), Name = "heatmap color scale", Shape = bar });
        var labelX = x + barWidth + 4;
        var labelWidth = Math.Max(20, width - barWidth - 4);
        group.Children.Add(VectorChartTextElement(
            HeatmapNativeId(elementId, "colorbar/max"),
            "heatmap scale maximum",
            labelX,
            y,
            labelWidth,
            14,
            maximum.ToString("0.##", CultureInfo.InvariantCulture),
            axisStyle,
            catalog,
            7,
            "left"));
        group.Children.Add(VectorChartTextElement(
            HeatmapNativeId(elementId, "colorbar/min"),
            "heatmap scale minimum",
            labelX,
            y + height - 14,
            labelWidth,
            14,
            minimum.ToString("0.##", CultureInfo.InvariantCulture),
            axisStyle,
            catalog,
            7,
            "left"));
        if (scale == "diverging")
            group.Children.Add(VectorChartTextElement(
                HeatmapNativeId(elementId, "colorbar/mid"),
                "heatmap scale midpoint",
                labelX,
                y + height / 2 - 7,
                labelWidth,
                14,
                midpoint.ToString("0.##", CultureInfo.InvariantCulture),
                axisStyle,
                catalog,
                7,
                "left"));
    }

    private static (byte R, byte G, byte B, double Alpha) HeatmapColor((string Rgb, double Alpha) color) => (
        byte.Parse(color.Rgb.AsSpan(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
        byte.Parse(color.Rgb.AsSpan(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
        byte.Parse(color.Rgb.AsSpan(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
        color.Alpha);

    private static string HeatmapNativeId(string elementId, string suffix) =>
        $"{elementId}/heatmap/{suffix}";

    private static (byte R, byte G, byte B, double Alpha) HeatmapInterpolate(
        double value,
        double minimum,
        double maximum,
        double midpoint,
        string scale,
        IReadOnlyList<(byte R, byte G, byte B, double Alpha)> colors)
    {
        if (scale == "diverging")
        {
            if (value <= midpoint)
                return HeatmapMix(colors[0], colors[1], (Math.Clamp(value, minimum, midpoint) - minimum) / (midpoint - minimum));
            return HeatmapMix(colors[1], colors[2], (Math.Clamp(value, midpoint, maximum) - midpoint) / (maximum - midpoint));
        }
        return HeatmapMix(colors[0], colors[1], (Math.Clamp(value, minimum, maximum) - minimum) / (maximum - minimum));
    }

    private static (byte R, byte G, byte B, double Alpha) HeatmapMix(
        (byte R, byte G, byte B, double Alpha) from,
        (byte R, byte G, byte B, double Alpha) to,
        double amount) => (
            checked((byte)Math.Round(from.R + (to.R - from.R) * amount)),
            checked((byte)Math.Round(from.G + (to.G - from.G) * amount)),
            checked((byte)Math.Round(from.B + (to.B - from.B) * amount)),
            from.Alpha + (to.Alpha - from.Alpha) * amount);

    private static string HeatmapHex((byte R, byte G, byte B, double Alpha) color) =>
        $"{color.R:X2}{color.G:X2}{color.B:X2}";

    private static double HeatmapLuminance((byte R, byte G, byte B, double Alpha) color)
    {
        static double Channel(byte value)
        {
            var normalized = value / 255d;
            return normalized <= 0.03928 ? normalized / 12.92 : Math.Pow((normalized + 0.055) / 1.055, 2.4);
        }
        return 0.2126 * Channel(color.R) + 0.7152 * Channel(color.G) + 0.0722 * Channel(color.B);
    }

    private static PresentationGroup BuildCandlestick(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ValidateCandlestickCompileProfile(element, raw, namedStyle, inlineStyle);
        var candlestick = FirstProperty(inlineStyle, namedStyle, "candlestick")!.Value;
        var series = element.Data.Series[0];
        var categories = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var closeValues = series.Values.Select(value => value!.Value).ToArray();
        var openValues = series.OpenValues;
        var highValues = series.HighValues;
        var lowValues = series.LowValues;
        var isOhlc = openValues.Count != 0;
        var seriesJson = raw.GetProperty("data").GetProperty("series").EnumerateArray().ToArray();
        var overlays = element.Data.Series.Skip(1)
            .Select((item, index) => (Series: item, Json: seriesJson[index + 1], Index: index + 1))
            .ToArray();

        var xAxis = Property(raw, "xAxis");
        var yAxis = Property(raw, "yAxis");
        var showXAxis = xAxis is null || OptionalBoolean(xAxis.Value, "visible") != false;
        var showYAxis = yAxis is null || OptionalBoolean(yAxis.Value, "visible") != false;
        var titleText = element.Title is null ? string.Empty : Flatten(element.Title);
        var titleStyle = FirstProperty(inlineStyle, namedStyle, "titleTextStyle");
        var axisStyle = Property(candlestick, "axisTextStyle");
        var valueStyle = Property(candlestick, "valueTextStyle");
        var xAxisTextStyle = Property(xAxis, "textStyle") ?? axisStyle;
        var yAxisTextStyle = Property(yAxis, "textStyle") ?? axisStyle;
        var xAxisTitleStyle = Property(xAxis, "titleTextStyle") ?? axisStyle;
        var yAxisTitleStyle = Property(yAxis, "titleTextStyle") ?? axisStyle;
        var showCloseValues = OptionalBoolean(candlestick, "showCloseValues") ?? false;
        if (showCloseValues && categories.Length > 16)
            throw Unsupported(element.Id, "showCloseValues is limited to 16 candlesticks to keep labels readable");

        var overlayValues = overlays.SelectMany(item => item.Series.Values).Select(value => value!.Value).ToArray();
        var dataMinimum = Math.Min(lowValues.Min(), overlayValues.DefaultIfEmpty(double.PositiveInfinity).Min());
        var dataMaximum = Math.Max(highValues.Max(), overlayValues.DefaultIfEmpty(double.NegativeInfinity).Max());
        if (overlays.Any(item => item.Series.ChartType is "area" or "column"))
        {
            dataMinimum = Math.Min(dataMinimum, 0);
            dataMaximum = Math.Max(dataMaximum, 0);
        }
        var dataRange = dataMaximum - dataMinimum;
        if (dataRange <= 0) dataRange = Math.Max(1, Math.Abs(dataMaximum) * 0.1);
        var majorUnit = yAxis is { } y && OptionalDouble(y, "majorUnit") is { } explicitMajor
            ? explicitMajor
            : NiceVectorChartStep(dataRange / 4);
        var pad = dataRange * 0.05;
        var minimum = yAxis is { } minimumAxis && OptionalDouble(minimumAxis, "min") is { } explicitMinimum
            ? explicitMinimum
            : Math.Floor((dataMinimum - pad) / majorUnit) * majorUnit;
        var maximum = yAxis is { } maximumAxis && OptionalDouble(maximumAxis, "max") is { } explicitMaximum
            ? explicitMaximum
            : Math.Ceiling((dataMaximum + pad) / majorUnit) * majorUnit;
        if (maximum <= minimum) throw Unsupported(element.Id, "candlestick yAxis maximum must be greater than its minimum");

        var tickCount = checked((int)Math.Floor((maximum - minimum) / majorUnit + 1e-9)) + 1;
        if (tickCount is < 2 or > 12)
            throw Unsupported(element.Id, "candlestick yAxis must expand to between 2 and 12 major ticks");

        var x = element.Frame.X;
        var yPosition = element.Frame.Y;
        var width = element.Frame.Width;
        var height = element.Frame.Height;
        var titleHeight = titleText.Length == 0 ? 0 : Math.Clamp(height * 0.12, 22, 38);
        var leftInset = showYAxis ? Math.Clamp(width * 0.1, 42, 68) : 8;
        var rightInset = showCloseValues ? Math.Clamp(width * 0.08, 32, 48) : 10;
        var bottomInset = showXAxis
            ? (xAxis is { } xAxisValue && xAxisValue.TryGetProperty("title", out _) ? 38 : 24)
            : 8;
        var topInset = titleHeight + 8;
        var plotX = x + leftInset;
        var plotY = yPosition + topInset;
        var plotWidth = width - leftInset - rightInset;
        var plotHeight = height - topInset - bottomInset;
        if (plotWidth < 120 || plotHeight < 80)
            throw Unsupported(element.Id, "candlestick frame is too small for native axes and editable marks");
        var slotWidth = plotWidth / categories.Length;
        if (slotWidth < 4.5)
            throw Unsupported(element.Id, "candlestick frame is too narrow for the requested observation count");

        var bodyWidthRatio = OptionalDouble(candlestick, "bodyWidthRatio") ?? 0.55;
        var bodyWidth = Math.Max(1.25, slotWidth * bodyWidthRatio);
        var labelInterval = xAxis is { } xAxisLabel && OptionalDouble(xAxisLabel, "tickLabelInterval") is { } interval
            ? checked((int)interval)
            : Math.Max(1, checked((int)Math.Ceiling(categories.Length / 12d)));
        var numberFormat = yAxis is { } formattedAxis ? OptionalString(formattedAxis, "numberFormat") : null;
        if (numberFormat is not null && !VectorChartNumberFormats.Contains(numberFormat, StringComparer.Ordinal))
            throw Unsupported(element.Id, $"candlestick yAxis numberFormat {numberFormat} is outside the bounded profile");

        double PlotY(double value) => plotY + (maximum - value) * plotHeight / (maximum - minimum);
        var group = new PresentationGroup
        {
            LeftEmu = Emu(x),
            TopEmu = Emu(yPosition),
            WidthEmu = Emu(width),
            HeightEmu = Emu(height),
            ChildLeftEmu = Emu(x),
            ChildTopEmu = Emu(yPosition),
            ChildWidthEmu = Emu(width),
            ChildHeightEmu = Emu(height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;

        if (titleText.Length > 0)
            group.Children.Add(VectorChartTitleElement(
                CandlestickNativeId(element.Id, "title"),
                "candlestick title",
                x,
                yPosition,
                width,
                titleHeight,
                raw.GetProperty("title"),
                titleStyle,
                catalog,
                Math.Clamp(titleHeight * 0.48, 11, 18),
                "left"));

        if (showYAxis)
        {
            var gridlineStroke = Property(candlestick, "gridlineStroke");
            for (var tickIndex = 0; tickIndex < tickCount; tickIndex++)
            {
                var tickValue = minimum + tickIndex * majorUnit;
                var tickY = PlotY(tickValue);
                if (gridlineStroke is { } gridline)
                    group.Children.Add(VectorChartLineElement(
                        CandlestickNativeId(element.Id, $"grid/{tickIndex}"),
                        $"candlestick gridline {tickIndex + 1}",
                        plotX,
                        tickY,
                        plotX + plotWidth,
                        tickY,
                        gridline,
                        catalog));
                group.Children.Add(VectorChartTextElement(
                    CandlestickNativeId(element.Id, $"y-label/{tickIndex}"),
                    $"candlestick value-axis label {tickIndex + 1}",
                    x,
                    tickY - 7,
                    leftInset - 7,
                    14,
                    FormatCandlestickValue(tickValue, numberFormat),
                    yAxisTextStyle,
                    catalog,
                    8,
                    "right"));
            }
            if (yAxis is { } yAxisWithTitle && yAxisWithTitle.TryGetProperty("title", out var yTitle))
                group.Children.Add(VectorChartTextElement(
                    CandlestickNativeId(element.Id, "y-title"),
                    "candlestick value-axis title",
                    x,
                    plotY - 18,
                    leftInset + Math.Min(120, plotWidth * 0.3),
                    16,
                    yTitle.GetString()!,
                    yAxisTitleStyle,
                    catalog,
                    8,
                    "left"));
        }

        var overlayPoints = overlays.ToDictionary(
            item => item.Series.Id,
            item => item.Series.Values.Select((value, index) => (
                X: plotX + (index + 0.5) * slotWidth,
                Y: PlotY(value!.Value),
                Value: value.Value)).ToArray(),
            StringComparer.Ordinal);
        var overlayBaselineY = PlotY(0);
        foreach (var item in overlays.Where(item => item.Series.ChartType == "area"))
        {
            var points = overlayPoints[item.Series.Id];
            var shape = ShapeFrame(new PpjFrameModel(plotX, plotY, plotWidth, plotHeight, 0, false, false), "custom");
            ApplyVectorSeriesFill(shape, item.Json, catalog, item.Index, element.Id, defaultOpacity: 0.28);
            shape.CustomPaths.Add(BuildVectorAreaPath(plotX, plotY, plotWidth, plotHeight, points.Select(point => (point.X, point.Y)).ToArray(), overlayBaselineY));
            group.Children.Add(new PresentationElement
            {
                Id = CandlestickNativeId(element.Id, $"overlay-area/{item.Series.Id}"),
                Name = $"candlestick area overlay {item.Series.Name}",
                Shape = shape,
            });
        }
        var columnOverlays = overlays.Where(item => item.Series.ChartType == "column").ToArray();
        var overlayColumnWidth = columnOverlays.Length == 0 ? 0 : Math.Max(1.5, slotWidth * 0.7 / columnOverlays.Length);
        for (var seriesIndex = 0; seriesIndex < columnOverlays.Length; seriesIndex++)
        {
            var item = columnOverlays[seriesIndex];
            var offset = seriesIndex - (columnOverlays.Length - 1) / 2d;
            var points = overlayPoints[item.Series.Id];
            for (var pointIndex = 0; pointIndex < points.Length; pointIndex++)
            {
                var point = points[pointIndex];
                var top = Math.Min(point.Y, overlayBaselineY);
                var columnHeight = Math.Max(1.5, Math.Abs(point.Y - overlayBaselineY));
                var shape = ShapeFrame(new PpjFrameModel(point.X + offset * overlayColumnWidth - overlayColumnWidth * 0.42, top, overlayColumnWidth * 0.84, columnHeight, 0, false, false), "rect");
                ApplyVectorSeriesFill(shape, item.Json, catalog, item.Index, element.Id, defaultOpacity: 0.42);
                group.Children.Add(new PresentationElement
                {
                    Id = CandlestickNativeId(element.Id, $"overlay-column/{item.Series.Id}/{pointIndex}"),
                    Name = $"candlestick column overlay {item.Series.Name} {pointIndex + 1}",
                    Shape = shape,
                });
            }
        }

        var wickStyle = candlestick.GetProperty("wick");
        for (var index = 0; index < categories.Length; index++)
        {
            var centerX = plotX + (index + 0.5) * slotWidth;
            group.Children.Add(VectorChartLineElement(
                CandlestickNativeId(element.Id, $"wick/{index}"),
                $"candlestick wick {index + 1}",
                centerX,
                PlotY(highValues[index]),
                centerX,
                PlotY(lowValues[index]),
                wickStyle,
                catalog));

            if (isOhlc)
            {
                var open = openValues[index];
                var close = closeValues[index];
                var top = PlotY(Math.Max(open, close));
                var bottom = PlotY(Math.Min(open, close));
                var bodyHeight = Math.Max(1.5, bottom - top);
                if (bodyHeight > bottom - top) top = Math.Clamp((top + bottom - bodyHeight) / 2, plotY, plotY + plotHeight - bodyHeight);
                var role = close >= open ? "up" : "down";
                var bodyStyle = candlestick.GetProperty(role);
                var body = ShapeFrame(
                    new PpjFrameModel(centerX - bodyWidth / 2, top, bodyWidth, bodyHeight, 0, false, false),
                    "rect");
                ApplyTextBoxFill(body, bodyStyle.GetProperty("fill"), catalog);
                if (bodyStyle.TryGetProperty("stroke", out var bodyStroke)) ApplyLine(body, bodyStroke, catalog);
                else body.LineStyle = "none";
                group.Children.Add(new PresentationElement
                {
                    Id = CandlestickNativeId(element.Id, $"body/{index}"),
                    Name = $"candlestick {role} body {index + 1}",
                    Shape = body,
                });
            }
            else
            {
                var closeY = PlotY(closeValues[index]);
                group.Children.Add(VectorChartLineElement(
                    CandlestickNativeId(element.Id, $"close/{index}"),
                    $"candlestick close tick {index + 1}",
                    centerX,
                    closeY,
                    centerX + bodyWidth / 2,
                    closeY,
                    wickStyle,
                    catalog));
            }

            if (showXAxis && index % labelInterval == 0)
                group.Children.Add(VectorChartTextElement(
                    CandlestickNativeId(element.Id, $"x-label/{index}"),
                    $"candlestick category label {index + 1}",
                    plotX + index * slotWidth,
                    plotY + plotHeight + 3,
                    Math.Max(slotWidth, 30),
                    15,
                    categories[index],
                    xAxisTextStyle,
                    catalog,
                    Math.Clamp(slotWidth * 0.24, 6, 9),
                    "center"));

            if (showCloseValues)
                group.Children.Add(VectorChartTextElement(
                    CandlestickNativeId(element.Id, $"close-value/{index}"),
                    $"candlestick close value {index + 1}",
                    centerX - slotWidth / 2,
                    Math.Max(plotY, PlotY(highValues[index]) - 16),
                    slotWidth,
                    14,
                    FormatCandlestickValue(closeValues[index], numberFormat),
                    valueStyle,
                    catalog,
                    Math.Clamp(slotWidth * 0.22, 6, 8),
                    "center"));
        }

        foreach (var item in overlays.Where(item => item.Series.ChartType == "line"))
            AddVectorLineSeries(
                group,
                element.Id,
                "candlestick",
                item.Series,
                item.Json,
                item.Index,
                overlayPoints[item.Series.Id],
                catalog,
                drawDefaultMarkers: false);

        if (showXAxis && xAxis is { } xAxisWithTitle && xAxisWithTitle.TryGetProperty("title", out var xTitle))
            group.Children.Add(VectorChartTextElement(
                CandlestickNativeId(element.Id, "x-title"),
                "candlestick category-axis title",
                plotX,
                yPosition + height - 17,
                plotWidth,
                16,
                xTitle.GetString()!,
                xAxisTitleStyle,
                catalog,
                8,
                "center"));

        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static void ValidateCandlestickCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        if (element.Data.Categories.Count is < 1 or > 64 ||
            element.Data.Categories.Any(category => category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString())) ||
            element.Data.Categories.Select(category => category.GetString()).Distinct(StringComparer.Ordinal).Count() != element.Data.Categories.Count)
            throw Unsupported(element.Id, "candlestick categories must be 1..64 unique non-empty strings");
        if (element.Data.Series.Count is < 1 or > 5)
            throw Unsupported(element.Id, "candlestick charts require one OHLC or HLC series and at most four overlays");
        var series = element.Data.Series[0];
        var count = element.Data.Categories.Count;
        if (series.Values.Count != count || series.Values.Any(value => value is null) ||
            series.HighValues.Count != count || series.LowValues.Count != count ||
            (series.OpenValues.Count != 0 && series.OpenValues.Count != count))
            throw Unsupported(element.Id, "candlestick open/high/low/close channels must align with the category count");
        for (var index = 0; index < count; index++)
        {
            var low = series.LowValues[index];
            var high = series.HighValues[index];
            var close = series.Values[index]!.Value;
            if (low > high || close < low || close > high ||
                (series.OpenValues.Count != 0 && (series.OpenValues[index] < low || series.OpenValues[index] > high)))
                throw Unsupported(element.Id, "every open and close must lie inside its low/high interval");
        }
        foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
            if (series.Raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, $"candlestick series do not support {property}");
        for (var index = 1; index < element.Data.Series.Count; index++)
        {
            var overlay = element.Data.Series[index];
            if (overlay.ChartType is not ("line" or "area" or "column"))
                throw Unsupported(element.Id, "candlestick overlays support line, area and column series");
            if (overlay.Values.Count != count || overlay.Values.Any(value => value is null || !double.IsFinite(value.Value)))
                throw Unsupported(element.Id, $"candlestick overlay {overlay.Id} requires one complete finite value per category");
            RejectProperties(overlay.Raw, element.Id, "pointRoles", "xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "parents", "sources", "targets", "axis", "symbol", "trendlines", "errorBars");
            if (overlay.Raw.TryGetProperty("fill", out _) && overlay.Raw.TryGetProperty("color", out _))
                throw Unsupported(element.Id, "candlestick overlay color and fill are aliases and cannot both be present");
            if (overlay.ChartType is "area" or "column" && overlay.Raw.TryGetProperty("marker", out _))
                throw Unsupported(element.Id, $"candlestick {overlay.ChartType} overlays do not render markers");
        }
        foreach (var property in new[] { "secondaryXAxis", "secondaryYAxis" })
            if (raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, "candlestick charts do not support secondary axes");
        RejectTokenizedVectorAxis(Property(raw, "xAxis"), element.Id, "X");
        RejectTokenizedVectorAxis(Property(raw, "yAxis"), element.Id, "Y");
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "startAngle", "holeSize", "bubbleScale", "bubbleSizeMode", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap" })
            if (FirstProperty(inlineStyle, namedStyle, property) is not null)
                throw Unsupported(element.Id, $"candlestick charts do not support chart style field {property}");
        if (FirstProperty(inlineStyle, namedStyle, "candlestick") is not { } candlestick)
            throw Unsupported(element.Id, "candlestick charts require style.candlestick");
        foreach (var role in new[] { "up", "down" })
        {
            var fillType = candlestick.GetProperty(role).GetProperty("fill").GetProperty("type").GetString();
            if (fillType is "none" or "image")
                throw Unsupported(element.Id, $"candlestick {role} fill must be solid or a bounded gradient");
        }
        if (raw.TryGetProperty("yAxis", out var yAxis))
        {
            var overlayValues = element.Data.Series.Skip(1).SelectMany(item => item.Values).Select(value => value!.Value).ToArray();
            var lowest = Math.Min(series.LowValues.Min(), overlayValues.DefaultIfEmpty(double.PositiveInfinity).Min());
            var highest = Math.Max(series.HighValues.Max(), overlayValues.DefaultIfEmpty(double.NegativeInfinity).Max());
            if (element.Data.Series.Skip(1).Any(item => item.ChartType is "area" or "column"))
            {
                lowest = Math.Min(lowest, 0);
                highest = Math.Max(highest, 0);
            }
            if (OptionalDouble(yAxis, "min") is { } minimum && minimum > lowest)
                throw Unsupported(element.Id, "candlestick yAxis.min must not clip the lowest observation");
            if (OptionalDouble(yAxis, "max") is { } maximum && maximum < highest)
                throw Unsupported(element.Id, "candlestick yAxis.max must not clip the highest observation");
        }
    }

    private static PresentationElement VectorChartLineElement(
        string id,
        string name,
        double startX,
        double startY,
        double endX,
        double endY,
        JsonElement stroke,
        Catalog catalog)
    {
        var connector = new PresentationConnector
        {
            ConnectorType = "straight",
            StartXEmu = Emu(startX),
            StartYEmu = Emu(startY),
            EndXEmu = Emu(endX),
            EndYEmu = Emu(endY),
        };
        ApplyLine(connector, stroke, catalog);
        return new PresentationElement { Id = id, Name = name, Connector = connector };
    }

    private static string CandlestickNativeId(string elementId, string suffix) =>
        $"{elementId}/candlestick/{suffix}";

    private static double NiceVectorChartStep(double raw)
    {
        if (!double.IsFinite(raw) || raw <= 0) return 1;
        var magnitude = Math.Pow(10, Math.Floor(Math.Log10(raw)));
        var normalized = raw / magnitude;
        var nice = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
        return nice * magnitude;
    }

    private static string FormatCandlestickValue(double value, string? numberFormat) =>
        value.ToString(numberFormat ?? "0.##", CultureInfo.InvariantCulture);

    private static PresentationGroup BuildTreemap(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ValidateTreemapCompileProfile(element, raw, namedStyle, inlineStyle);
        var treemap = FirstProperty(inlineStyle, namedStyle, "treemap")!.Value;
        var series = element.Data.Series[0];
        var names = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var values = series.Values.Select(value => value!.Value).ToArray();
        var indexes = names.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.Ordinal);
        var nodes = names.Select((name, index) => new TreemapNode(
            index,
            name,
            values[index],
            series.Parents[index] is { } parent ? indexes[parent] : null)).ToArray();
        foreach (var node in nodes)
            if (node.ParentIndex is { } parentIndex) nodes[parentIndex].Children.Add(node.Index);
        var roots = nodes.Where(node => node.ParentIndex is null).Select(node => node.Index).ToArray();
        var visibleLevels = series.Levels ?? 8;

        var rootColors = treemap.GetProperty("rootColors").EnumerateArray()
            .Select(catalog.Color)
            .ToArray();
        var border = Property(treemap, "border");
        var gap = OptionalDouble(treemap, "gap") ?? 2;
        var headerHeight = OptionalDouble(treemap, "headerHeight") ?? 17;
        var depthLighten = OptionalDouble(treemap, "depthLighten") ?? 0.08;
        var showValues = OptionalBoolean(treemap, "showValues") ?? false;
        var labelStyle = Property(treemap, "labelTextStyle");
        var valueStyle = Property(treemap, "valueTextStyle");
        var titleStyle = FirstProperty(inlineStyle, namedStyle, "titleTextStyle");

        var x = element.Frame.X;
        var y = element.Frame.Y;
        var width = element.Frame.Width;
        var height = element.Frame.Height;
        var titleText = element.Title is null ? string.Empty : Flatten(element.Title);
        var titleHeight = titleText.Length == 0 ? 0 : Math.Clamp(height * 0.12, 22, 38);
        var plot = new TreemapRectangle(x, y + titleHeight + 4, width, height - titleHeight - 4);
        if (plot.Width < 140 || plot.Height < 90)
            throw Unsupported(element.Id, "treemap frame is too small for a readable native hierarchy");

        var group = new PresentationGroup
        {
            LeftEmu = Emu(x),
            TopEmu = Emu(y),
            WidthEmu = Emu(width),
            HeightEmu = Emu(height),
            ChildLeftEmu = Emu(x),
            ChildTopEmu = Emu(y),
            ChildWidthEmu = Emu(width),
            ChildHeightEmu = Emu(height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;
        if (titleText.Length > 0)
            group.Children.Add(VectorChartTitleElement(
                TreemapNativeId(element.Id, "title"),
                "treemap title",
                x,
                y,
                width,
                titleHeight,
                raw.GetProperty("title"),
                titleStyle,
                catalog,
                Math.Clamp(titleHeight * 0.48, 11, 18),
                "left"));

        var rootPlacements = SquarifyTreemap(roots, nodes, plot);
        for (var rootOrder = 0; rootOrder < rootPlacements.Count; rootOrder++)
        {
            var placement = rootPlacements[rootOrder];
            RenderTreemapNode(
                group,
                element.Id,
                nodes,
                placement.NodeIndex,
                placement.Rectangle,
                rootColors[rootOrder % rootColors.Length],
                depth: 0,
                visibleLevels,
                gap,
                headerHeight,
                depthLighten,
                showValues,
                border,
                labelStyle,
                valueStyle,
                catalog);
        }

        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static void ValidateTreemapCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        var count = element.Data.Categories.Count;
        if (count is < 1 or > 128 ||
            element.Data.Categories.Any(category => category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString())) ||
            element.Data.Categories.Select(category => category.GetString()).Distinct(StringComparer.Ordinal).Count() != count)
            throw Unsupported(element.Id, "treemap categories must be 1..128 unique non-empty strings");
        if (element.Data.Series.Count != 1)
            throw Unsupported(element.Id, "treemap charts require exactly one hierarchy series");
        var series = element.Data.Series[0];
        if (series.Levels is < 1 or > 8)
            throw Unsupported(element.Id, "treemap display levels must be between one and eight");
        if (series.Values.Count != count || series.Values.Any(value => value is null || value <= 0) || series.Parents.Count != count)
            throw Unsupported(element.Id, "treemap values and parents must be complete, aligned, and strictly positive");
        foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
            if (series.Raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, $"treemap series do not support {property}");
        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, "treemap charts do not use Cartesian axes");
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "startAngle", "holeSize", "bubbleScale", "bubbleSizeMode", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap", "candlestick" })
            if (FirstProperty(inlineStyle, namedStyle, property) is not null)
                throw Unsupported(element.Id, $"treemap charts do not support chart style field {property}");
        if (FirstProperty(inlineStyle, namedStyle, "treemap") is null)
            throw Unsupported(element.Id, "treemap charts require style.treemap");

        var names = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var indexes = names.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.Ordinal);
        var rootCount = 0;
        for (var index = 0; index < count; index++)
        {
            var parent = series.Parents[index];
            if (parent is null)
            {
                rootCount++;
                continue;
            }
            if (!indexes.ContainsKey(parent))
                throw Unsupported(element.Id, $"treemap parent {parent} does not name a declared category");
            if (string.Equals(parent, names[index], StringComparison.Ordinal))
                throw Unsupported(element.Id, "a treemap node cannot parent itself");
        }
        if (rootCount is < 1 or > 16)
            throw Unsupported(element.Id, "treemap charts require between 1 and 16 roots");

        for (var index = 0; index < count; index++)
        {
            var seen = new HashSet<int>();
            var current = index;
            var depth = 0;
            while (true)
            {
                if (!seen.Add(current)) throw Unsupported(element.Id, $"treemap parent chain for {names[index]} contains a cycle");
                var parent = series.Parents[current];
                if (parent is null) break;
                current = indexes[parent];
                depth++;
                if (depth > 8) throw Unsupported(element.Id, $"treemap node {names[index]} exceeds the maximum depth of eight");
            }
        }

        var childSums = new Dictionary<int, double>();
        for (var index = 0; index < count; index++)
            if (series.Parents[index] is { } parent)
            {
                var parentIndex = indexes[parent];
                childSums[parentIndex] = childSums.GetValueOrDefault(parentIndex) + series.Values[index]!.Value;
            }
        foreach (var pair in childSums)
        {
            var declared = series.Values[pair.Key]!.Value;
            var tolerance = Math.Max(1e-9, Math.Abs(declared) * 1e-9);
            if (Math.Abs(declared - pair.Value) > tolerance)
                throw Unsupported(element.Id, $"treemap parent {names[pair.Key]} must equal its direct-child sum");
        }
    }

    private static void RenderTreemapNode(
        PresentationGroup group,
        string elementId,
        IReadOnlyList<TreemapNode> nodes,
        int nodeIndex,
        TreemapRectangle allocated,
        (string Rgb, double Alpha) rootColor,
        int depth,
        int visibleLevels,
        double gap,
        double headerHeight,
        double depthLighten,
        bool showValues,
        JsonElement? border,
        JsonElement? labelStyle,
        JsonElement? valueStyle,
        Catalog catalog)
    {
        var node = nodes[nodeIndex];
        var localGap = Math.Min(gap, Math.Max(0, Math.Min(allocated.Width, allocated.Height) - 0.5) / 2);
        var rectangle = allocated.Inset(localGap / 2);
        if (rectangle.Width < 0.5 || rectangle.Height < 0.5)
            throw Unsupported(elementId, $"treemap node {node.Name} is too small for a stable native rectangle");

        var color = LightenTreemapColor(rootColor, Math.Min(0.72, depth * depthLighten));
        var shape = ShapeFrame(
            new PpjFrameModel(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height, 0, false, false),
            "rect");
        shape.FillRgb = color.Rgb;
        if (color.Alpha < 1) shape.FillOpacityThousandthPercent = Opacity(color.Alpha);
        if (border is { } stroke) ApplyLine(shape, stroke, catalog);
        else shape.LineStyle = "none";
        group.Children.Add(new PresentationElement
        {
            Id = TreemapNativeId(elementId, $"node/{node.Index}"),
            Name = $"treemap node {node.Name}",
            Shape = shape,
        });

        var contrast = TreemapLuminance(color.Rgb) > 0.5 ? "111827" : "FFFFFF";
        if (node.Children.Count != 0 && depth + 1 < visibleLevels)
        {
            var effectiveHeader = Math.Min(headerHeight, Math.Max(0, rectangle.Height * 0.28));
            if (effectiveHeader >= 9 && rectangle.Width >= 28)
                group.Children.Add(VectorChartTextElement(
                    TreemapNativeId(elementId, $"label/{node.Index}"),
                    $"treemap label {node.Name}",
                    rectangle.X + 3,
                    rectangle.Y + 1,
                    rectangle.Width - 6,
                    effectiveHeader - 2,
                    node.Name,
                    labelStyle,
                    catalog,
                    Math.Clamp(Math.Min(effectiveHeader * 0.55, rectangle.Width / Math.Max(4, node.Name.Length) * 1.4), 6, 11),
                    "left",
                    contrast));
            var childArea = new TreemapRectangle(
                rectangle.X,
                rectangle.Y + effectiveHeader,
                rectangle.Width,
                rectangle.Height - effectiveHeader);
            if (childArea.Width < 1 || childArea.Height < 1)
                throw Unsupported(elementId, $"treemap node {node.Name} leaves no room for its children");
            foreach (var placement in SquarifyTreemap(node.Children, nodes, childArea))
                RenderTreemapNode(
                    group,
                    elementId,
                    nodes,
                    placement.NodeIndex,
                    placement.Rectangle,
                    rootColor,
                    depth + 1,
                    visibleLevels,
                    gap,
                    headerHeight,
                    depthLighten,
                    showValues,
                    border,
                    labelStyle,
                    valueStyle,
                    catalog);
            return;
        }

        if (rectangle.Width < 28 || rectangle.Height < 14) return;
        var labelHeight = showValues && rectangle.Height >= 28 ? rectangle.Height * 0.58 : rectangle.Height;
        group.Children.Add(VectorChartTextElement(
            TreemapNativeId(elementId, $"label/{node.Index}"),
            $"treemap label {node.Name}",
            rectangle.X + 3,
            rectangle.Y + 1,
            rectangle.Width - 6,
            Math.Max(10, labelHeight - 2),
            node.Name,
            labelStyle,
            catalog,
            Math.Clamp(Math.Min(labelHeight * 0.34, rectangle.Width / Math.Max(4, node.Name.Length) * 1.5), 6, 12),
            "left",
            contrast));
        if (showValues && rectangle.Height >= 28)
            group.Children.Add(VectorChartTextElement(
                TreemapNativeId(elementId, $"value/{node.Index}"),
                $"treemap value {node.Name}",
                rectangle.X + 3,
                rectangle.Y + labelHeight,
                rectangle.Width - 6,
                rectangle.Height - labelHeight - 2,
                node.Value.ToString("0.##", CultureInfo.InvariantCulture),
                valueStyle,
                catalog,
                Math.Clamp((rectangle.Height - labelHeight) * 0.45, 6, 10),
                "left",
                contrast));
    }

    private static IReadOnlyList<TreemapPlacement> SquarifyTreemap(
        IReadOnlyList<int> nodeIndexes,
        IReadOnlyList<TreemapNode> nodes,
        TreemapRectangle rectangle)
    {
        var total = nodeIndexes.Sum(index => nodes[index].Value);
        var scale = rectangle.Width * rectangle.Height / total;
        var pending = nodeIndexes
            .Select((nodeIndex, order) => new TreemapArea(nodeIndex, nodes[nodeIndex].Value * scale, order))
            .OrderByDescending(item => item.Area)
            .ThenBy(item => item.Order)
            .ToList();
        var placements = new List<TreemapPlacement>(pending.Count);
        var row = new List<TreemapArea>();
        var remaining = rectangle;
        while (pending.Count != 0)
        {
            var candidate = pending[0];
            var side = Math.Min(remaining.Width, remaining.Height);
            if (row.Count == 0 || TreemapWorst(row.Append(candidate), side) <= TreemapWorst(row, side))
            {
                row.Add(candidate);
                pending.RemoveAt(0);
                continue;
            }
            remaining = LayoutTreemapRow(row, remaining, placements);
            row.Clear();
        }
        if (row.Count != 0) LayoutTreemapRow(row, remaining, placements);
        return placements;
    }

    private static double TreemapWorst(IEnumerable<TreemapArea> row, double side)
    {
        var values = row.Select(item => item.Area).ToArray();
        var sum = values.Sum();
        var squaredSide = side * side;
        return Math.Max(squaredSide * values.Max() / (sum * sum), (sum * sum) / (squaredSide * values.Min()));
    }

    private static TreemapRectangle LayoutTreemapRow(
        IReadOnlyList<TreemapArea> row,
        TreemapRectangle rectangle,
        ICollection<TreemapPlacement> placements)
    {
        var area = row.Sum(item => item.Area);
        if (rectangle.Width >= rectangle.Height)
        {
            var rowWidth = area / rectangle.Height;
            var cursorY = rectangle.Y;
            for (var index = 0; index < row.Count; index++)
            {
                var height = index == row.Count - 1
                    ? rectangle.Y + rectangle.Height - cursorY
                    : row[index].Area / rowWidth;
                placements.Add(new TreemapPlacement(row[index].NodeIndex, new TreemapRectangle(rectangle.X, cursorY, rowWidth, height)));
                cursorY += height;
            }
            return new TreemapRectangle(rectangle.X + rowWidth, rectangle.Y, Math.Max(0, rectangle.Width - rowWidth), rectangle.Height);
        }

        var rowHeight = area / rectangle.Width;
        var cursorX = rectangle.X;
        for (var index = 0; index < row.Count; index++)
        {
            var width = index == row.Count - 1
                ? rectangle.X + rectangle.Width - cursorX
                : row[index].Area / rowHeight;
            placements.Add(new TreemapPlacement(row[index].NodeIndex, new TreemapRectangle(cursorX, rectangle.Y, width, rowHeight)));
            cursorX += width;
        }
        return new TreemapRectangle(rectangle.X, rectangle.Y + rowHeight, rectangle.Width, Math.Max(0, rectangle.Height - rowHeight));
    }

    private static (string Rgb, double Alpha) LightenTreemapColor((string Rgb, double Alpha) color, double amount)
    {
        byte Lighten(byte channel) => checked((byte)Math.Round(channel + (255 - channel) * amount));
        var red = byte.Parse(color.Rgb.AsSpan(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var green = byte.Parse(color.Rgb.AsSpan(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var blue = byte.Parse(color.Rgb.AsSpan(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        return ($"{Lighten(red):X2}{Lighten(green):X2}{Lighten(blue):X2}", color.Alpha);
    }

    private static double TreemapLuminance(string rgb)
    {
        static double Channel(byte value)
        {
            var normalized = value / 255d;
            return normalized <= 0.03928 ? normalized / 12.92 : Math.Pow((normalized + 0.055) / 1.055, 2.4);
        }
        var red = byte.Parse(rgb.AsSpan(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var green = byte.Parse(rgb.AsSpan(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var blue = byte.Parse(rgb.AsSpan(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        return 0.2126 * Channel(red) + 0.7152 * Channel(green) + 0.0722 * Channel(blue);
    }

    private static string TreemapNativeId(string elementId, string suffix) =>
        $"{elementId}/treemap/{suffix}";

    private sealed class TreemapNode(int index, string name, double value, int? parentIndex)
    {
        internal int Index { get; } = index;
        internal string Name { get; } = name;
        internal double Value { get; } = value;
        internal int? ParentIndex { get; } = parentIndex;
        internal List<int> Children { get; } = [];
    }

    private readonly record struct TreemapRectangle(double X, double Y, double Width, double Height)
    {
        internal TreemapRectangle Inset(double amount) => new(
            X + amount,
            Y + amount,
            Math.Max(0, Width - amount * 2),
            Math.Max(0, Height - amount * 2));
    }

    private readonly record struct TreemapPlacement(int NodeIndex, TreemapRectangle Rectangle);
    private readonly record struct TreemapArea(int NodeIndex, double Area, int Order);

    private static PresentationGroup BuildSunburst(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ValidateSunburstCompileProfile(element, raw, namedStyle, inlineStyle);
        var sunburst = FirstProperty(inlineStyle, namedStyle, "sunburst")!.Value;
        var series = element.Data.Series[0];
        var names = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var values = series.Values.Select(value => value!.Value).ToArray();
        var indexes = names.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.Ordinal);
        var nodes = names.Select((name, index) => new SunburstNode(
            index,
            name,
            values[index],
            series.Parents[index] is { } parent ? indexes[parent] : null)).ToArray();
        foreach (var node in nodes)
            if (node.ParentIndex is { } parentIndex) nodes[parentIndex].Children.Add(node.Index);
        var roots = nodes.Where(node => node.ParentIndex is null).Select(node => node.Index).ToArray();
        foreach (var root in roots) AssignSunburstDepth(nodes, root, 0);
        var levelCount = nodes.Max(node => node.Depth) + 1;
        var visibleLevels = Math.Min(series.Levels ?? levelCount, levelCount);

        var rootColors = sunburst.GetProperty("rootColors").EnumerateArray().Select(catalog.Color).ToArray();
        var border = Property(sunburst, "border");
        var innerRadiusRatio = OptionalDouble(sunburst, "innerRadiusRatio") ?? 0;
        var ringGap = OptionalDouble(sunburst, "ringGap") ?? 1.5;
        var segmentGapRadians = (OptionalDouble(sunburst, "segmentGapDegrees") ?? 0.8) * Math.PI / 180;
        var startAngle = (OptionalDouble(sunburst, "startAngle") ?? -90) * Math.PI / 180;
        var clockwise = OptionalBoolean(sunburst, "clockwise") ?? true;
        var direction = clockwise ? 1d : -1d;
        var depthLighten = OptionalDouble(sunburst, "depthLighten") ?? 0.08;
        var showValues = OptionalBoolean(sunburst, "showValues") ?? false;
        var labelStyle = Property(sunburst, "labelTextStyle");
        var valueStyle = Property(sunburst, "valueTextStyle");
        var titleStyle = FirstProperty(inlineStyle, namedStyle, "titleTextStyle");

        var x = element.Frame.X;
        var y = element.Frame.Y;
        var width = element.Frame.Width;
        var height = element.Frame.Height;
        var titleText = element.Title is null ? string.Empty : Flatten(element.Title);
        var titleHeight = titleText.Length == 0 ? 0 : Math.Clamp(height * 0.12, 22, 38);
        var availableHeight = height - titleHeight - 4;
        var diameter = Math.Min(width, availableHeight);
        if (diameter < 160)
            throw Unsupported(element.Id, "sunburst frame is too small for a readable native hierarchy");
        var plotX = x + (width - diameter) / 2;
        var plotY = y + titleHeight + 4 + (availableHeight - diameter) / 2;
        var outerRadius = diameter / 2;
        var innerRadius = outerRadius * innerRadiusRatio;
        var ringWidth = (outerRadius - innerRadius) / visibleLevels;
        if (ringWidth - ringGap < 8)
            throw Unsupported(element.Id, "sunburst rings are too narrow after applying the configured gap");

        var group = new PresentationGroup
        {
            LeftEmu = Emu(x),
            TopEmu = Emu(y),
            WidthEmu = Emu(width),
            HeightEmu = Emu(height),
            ChildLeftEmu = Emu(x),
            ChildTopEmu = Emu(y),
            ChildWidthEmu = Emu(width),
            ChildHeightEmu = Emu(height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;
        if (titleText.Length > 0)
            group.Children.Add(VectorChartTitleElement(
                SunburstNativeId(element.Id, "title"),
                "sunburst title",
                x,
                y,
                width,
                titleHeight,
                raw.GetProperty("title"),
                titleStyle,
                catalog,
                Math.Clamp(titleHeight * 0.48, 11, 18),
                "left"));

        var total = roots.Sum(root => nodes[root].Value);
        var cursor = startAngle;
        for (var rootOrder = 0; rootOrder < roots.Length; rootOrder++)
        {
            var root = roots[rootOrder];
            var span = Math.PI * 2 * nodes[root].Value / total * direction;
            RenderSunburstNode(
                group,
                element.Id,
                nodes,
                root,
                cursor,
                cursor + span,
                rootColors[rootOrder % rootColors.Length],
                plotX,
                plotY,
                diameter,
                innerRadius,
                ringWidth,
                ringGap,
                segmentGapRadians,
                depthLighten,
                visibleLevels,
                showValues,
                border,
                labelStyle,
                valueStyle,
                catalog);
            cursor += span;
        }

        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static void ValidateSunburstCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        var count = element.Data.Categories.Count;
        if (count is < 1 or > 96 ||
            element.Data.Categories.Any(category => category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString())) ||
            element.Data.Categories.Select(category => category.GetString()).Distinct(StringComparer.Ordinal).Count() != count)
            throw Unsupported(element.Id, "sunburst categories must be 1..96 unique non-empty strings");
        if (element.Data.Series.Count != 1)
            throw Unsupported(element.Id, "sunburst charts require exactly one hierarchy series");
        var series = element.Data.Series[0];
        if (series.Levels is < 1 or > 6)
            throw Unsupported(element.Id, "sunburst display levels must be between one and six");
        if (series.Values.Count != count || series.Values.Any(value => value is null || value <= 0) || series.Parents.Count != count)
            throw Unsupported(element.Id, "sunburst values and parents must be complete, aligned, and strictly positive");
        foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
            if (series.Raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, $"sunburst series do not support {property}");
        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, "sunburst charts do not use Cartesian axes");
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "startAngle", "holeSize", "bubbleScale", "bubbleSizeMode", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap", "candlestick", "treemap" })
            if (FirstProperty(inlineStyle, namedStyle, property) is not null)
                throw Unsupported(element.Id, $"sunburst charts do not support chart style field {property}");
        if (FirstProperty(inlineStyle, namedStyle, "sunburst") is null)
            throw Unsupported(element.Id, "sunburst charts require style.sunburst");

        var names = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var indexes = names.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.Ordinal);
        var rootCount = 0;
        for (var index = 0; index < count; index++)
        {
            var parent = series.Parents[index];
            if (parent is null)
            {
                rootCount++;
                continue;
            }
            if (!indexes.ContainsKey(parent))
                throw Unsupported(element.Id, $"sunburst parent {parent} does not name a declared category");
            if (string.Equals(parent, names[index], StringComparison.Ordinal))
                throw Unsupported(element.Id, "a sunburst node cannot parent itself");
        }
        if (rootCount is < 1 or > 16)
            throw Unsupported(element.Id, "sunburst charts require between 1 and 16 roots");

        for (var index = 0; index < count; index++)
        {
            var seen = new HashSet<int>();
            var current = index;
            var depth = 0;
            while (true)
            {
                if (!seen.Add(current)) throw Unsupported(element.Id, $"sunburst parent chain for {names[index]} contains a cycle");
                var parent = series.Parents[current];
                if (parent is null) break;
                current = indexes[parent];
                depth++;
                if (depth > 6) throw Unsupported(element.Id, $"sunburst node {names[index]} exceeds the maximum depth of six");
            }
        }

        var childSums = new Dictionary<int, double>();
        for (var index = 0; index < count; index++)
            if (series.Parents[index] is { } parent)
            {
                var parentIndex = indexes[parent];
                childSums[parentIndex] = childSums.GetValueOrDefault(parentIndex) + series.Values[index]!.Value;
            }
        foreach (var pair in childSums)
        {
            var declared = series.Values[pair.Key]!.Value;
            var tolerance = Math.Max(1e-9, Math.Abs(declared) * 1e-9);
            if (Math.Abs(declared - pair.Value) > tolerance)
                throw Unsupported(element.Id, $"sunburst parent {names[pair.Key]} must equal its direct-child sum");
        }
    }

    private static void AssignSunburstDepth(IReadOnlyList<SunburstNode> nodes, int nodeIndex, int depth)
    {
        nodes[nodeIndex].Depth = depth;
        foreach (var child in nodes[nodeIndex].Children) AssignSunburstDepth(nodes, child, depth + 1);
    }

    private static void RenderSunburstNode(
        PresentationGroup group,
        string elementId,
        IReadOnlyList<SunburstNode> nodes,
        int nodeIndex,
        double startAngle,
        double endAngle,
        (string Rgb, double Alpha) rootColor,
        double plotX,
        double plotY,
        double diameter,
        double baseInnerRadius,
        double ringWidth,
        double ringGap,
        double segmentGapRadians,
        double depthLighten,
        int visibleLevels,
        bool showValues,
        JsonElement? border,
        JsonElement? labelStyle,
        JsonElement? valueStyle,
        Catalog catalog)
    {
        var node = nodes[nodeIndex];
        var sign = Math.Sign(endAngle - startAngle);
        var span = Math.Abs(endAngle - startAngle);
        var effectiveSpan = span - segmentGapRadians;
        if (effectiveSpan <= Math.PI / 720)
            throw Unsupported(elementId, $"sunburst node {node.Name} is too narrow after applying the configured segment gap");
        var visibleStart = startAngle + sign * segmentGapRadians / 2;
        var visibleEnd = endAngle - sign * segmentGapRadians / 2;
        var innerRadius = baseInnerRadius + node.Depth * ringWidth + ringGap / 2;
        var outerRadius = baseInnerRadius + (node.Depth + 1) * ringWidth - ringGap / 2;
        if (outerRadius <= innerRadius)
            throw Unsupported(elementId, $"sunburst node {node.Name} has no visible ring thickness");

        var color = LightenTreemapColor(rootColor, Math.Min(0.72, node.Depth * depthLighten));
        var shape = ShapeFrame(new PpjFrameModel(plotX, plotY, diameter, diameter, 0, false, false), "custom");
        shape.FillRgb = color.Rgb;
        if (color.Alpha < 1) shape.FillOpacityThousandthPercent = Opacity(color.Alpha);
        if (border is { } stroke) ApplyLine(shape, stroke, catalog);
        else shape.LineStyle = "none";
        shape.CustomPaths.Add(BuildSunburstSectorPath(
            diameter,
            innerRadius,
            outerRadius,
            visibleStart,
            visibleEnd,
            border is not null));
        group.Children.Add(new PresentationElement
        {
            Id = SunburstNativeId(elementId, $"sector/{node.Index}"),
            Name = $"sunburst sector {node.Name}",
            Shape = shape,
        });

        var midRadius = (innerRadius + outerRadius) / 2;
        var midAngle = (visibleStart + visibleEnd) / 2;
        var arcLength = midRadius * effectiveSpan;
        var thickness = outerRadius - innerRadius;
        var estimatedLabelWidth = Math.Max(28, node.Name.Length * 5.5);
        if (thickness >= (showValues ? 25 : 16) && arcLength >= estimatedLabelWidth * 1.35 && effectiveSpan >= Math.PI / 12)
        {
            var centerX = plotX + diameter / 2 + Math.Cos(midAngle) * midRadius;
            var centerY = plotY + diameter / 2 + Math.Sin(midAngle) * midRadius;
            var textWidth = Math.Min(Math.Max(estimatedLabelWidth, 36), Math.Min(110, arcLength * 0.72));
            var labelHeight = showValues ? Math.Min(14, thickness * 0.42) : Math.Min(16, thickness * 0.64);
            var contrast = TreemapLuminance(color.Rgb) > 0.5 ? "111827" : "FFFFFF";
            group.Children.Add(VectorChartTextElement(
                SunburstNativeId(elementId, $"label/{node.Index}"),
                $"sunburst label {node.Name}",
                centerX - textWidth / 2,
                centerY - (showValues ? labelHeight : labelHeight / 2),
                textWidth,
                labelHeight,
                node.Name,
                labelStyle,
                catalog,
                Math.Clamp(Math.Min(labelHeight * 0.62, textWidth / Math.Max(4, node.Name.Length) * 1.5), 6, 11),
                "center",
                contrast));
            if (showValues)
                group.Children.Add(VectorChartTextElement(
                    SunburstNativeId(elementId, $"value/{node.Index}"),
                    $"sunburst value {node.Name}",
                    centerX - textWidth / 2,
                    centerY,
                    textWidth,
                    Math.Min(12, thickness * 0.34),
                    node.Value.ToString("0.##", CultureInfo.InvariantCulture),
                    valueStyle,
                    catalog,
                    Math.Clamp(thickness * 0.18, 6, 9),
                    "center",
                    contrast));
        }

        if (node.Children.Count == 0 || node.Depth + 1 >= visibleLevels) return;
        var cursor = startAngle;
        var parentSpan = endAngle - startAngle;
        foreach (var childIndex in node.Children)
        {
            var childSpan = parentSpan * nodes[childIndex].Value / node.Value;
            RenderSunburstNode(
                group,
                elementId,
                nodes,
                childIndex,
                cursor,
                cursor + childSpan,
                rootColor,
                plotX,
                plotY,
                diameter,
                baseInnerRadius,
                ringWidth,
                ringGap,
                segmentGapRadians,
                depthLighten,
                visibleLevels,
                showValues,
                border,
                labelStyle,
                valueStyle,
                catalog);
            cursor += childSpan;
        }
    }

    private static PresentationCustomGeometryPath BuildSunburstSectorPath(
        double diameter,
        double innerRadius,
        double outerRadius,
        double startAngle,
        double endAngle,
        bool stroke)
    {
        const long viewport = 100_000;
        const double center = viewport / 2d;
        var scale = viewport / diameter;
        var inner = innerRadius * scale;
        var outer = outerRadius * scale;
        var path = new PresentationCustomGeometryPath
        {
            Width = viewport,
            Height = viewport,
            FillMode = PresentationCustomGeometryPath.Types.FillMode.Normal,
            Stroke = stroke,
        };
        path.Commands.Add(MoveTo(SunburstPoint(center, outer, startAngle)));
        AddSunburstArc(path, center, outer, startAngle, endAngle);
        if (inner <= 0.5)
        {
            path.Commands.Add(LineTo(new PresentationCustomGeometryPoint { X = viewport / 2, Y = viewport / 2 }));
        }
        else
        {
            path.Commands.Add(LineTo(SunburstPoint(center, inner, endAngle)));
            AddSunburstArc(path, center, inner, endAngle, startAngle);
        }
        path.Commands.Add(new PresentationCustomGeometryCommand { Close = true });
        return path;
    }

    private static void AddSunburstArc(
        PresentationCustomGeometryPath path,
        double center,
        double radius,
        double startAngle,
        double endAngle)
    {
        var span = endAngle - startAngle;
        var segments = Math.Max(1, (int)Math.Ceiling(Math.Abs(span) / (Math.PI / 2)));
        var step = span / segments;
        for (var segment = 0; segment < segments; segment++)
        {
            var start = startAngle + step * segment;
            var end = start + step;
            var kappa = 4d / 3d * Math.Tan(step / 4d);
            var startPoint = SunburstPoint(center, radius, start);
            var endPoint = SunburstPoint(center, radius, end);
            var control1 = new PresentationCustomGeometryPoint
            {
                X = checked((long)Math.Round(startPoint.X - kappa * radius * Math.Sin(start), MidpointRounding.AwayFromZero)),
                Y = checked((long)Math.Round(startPoint.Y + kappa * radius * Math.Cos(start), MidpointRounding.AwayFromZero)),
            };
            var control2 = new PresentationCustomGeometryPoint
            {
                X = checked((long)Math.Round(endPoint.X + kappa * radius * Math.Sin(end), MidpointRounding.AwayFromZero)),
                Y = checked((long)Math.Round(endPoint.Y - kappa * radius * Math.Cos(end), MidpointRounding.AwayFromZero)),
            };
            path.Commands.Add(new PresentationCustomGeometryCommand
            {
                CubicBezierTo = new PresentationCustomGeometryCubicBezier
                {
                    Control1 = control1,
                    Control2 = control2,
                    End = endPoint,
                },
            });
        }
    }

    private static PresentationCustomGeometryPoint SunburstPoint(double center, double radius, double angle) => new()
    {
        X = checked((long)Math.Round(center + radius * Math.Cos(angle), MidpointRounding.AwayFromZero)),
        Y = checked((long)Math.Round(center + radius * Math.Sin(angle), MidpointRounding.AwayFromZero)),
    };

    private static PresentationCustomGeometryCommand MoveTo(PresentationCustomGeometryPoint point) => new() { MoveTo = point };
    private static PresentationCustomGeometryCommand LineTo(PresentationCustomGeometryPoint point) => new() { LineTo = point };

    private static string SunburstNativeId(string elementId, string suffix) =>
        $"{elementId}/sunburst/{suffix}";

    private sealed class SunburstNode(int index, string name, double value, int? parentIndex)
    {
        internal int Index { get; } = index;
        internal string Name { get; } = name;
        internal double Value { get; } = value;
        internal int? ParentIndex { get; } = parentIndex;
        internal int Depth { get; set; }
        internal List<int> Children { get; } = [];
    }

    private static PresentationGroup BuildSankey(PpjChartElementModel element, JsonElement raw, Catalog catalog)
    {
        var namedStyle = catalog.ChartStyle(element.StyleRef);
        var inlineStyle = Property(raw, "style");
        ValidateSankeyCompileProfile(element, raw, namedStyle, inlineStyle);
        var sankey = FirstProperty(inlineStyle, namedStyle, "sankey")!.Value;
        var series = element.Data.Series[0];
        var names = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var indexes = names.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.Ordinal);
        var nodes = names.Select((name, index) => new SankeyNode(index, name)).ToArray();
        var edges = Enumerable.Range(0, series.Values.Count).Select(index => new SankeyEdge(
            index,
            indexes[series.Sources[index]],
            indexes[series.Targets[index]],
            series.Values[index]!.Value)).ToArray();
        foreach (var edge in edges)
        {
            nodes[edge.SourceIndex].Outgoing.Add(edge.Index);
            nodes[edge.TargetIndex].Incoming.Add(edge.Index);
            nodes[edge.SourceIndex].OutgoingValue += edge.Value;
            nodes[edge.TargetIndex].IncomingValue += edge.Value;
        }

        var indegree = nodes.Select(node => node.Incoming.Count).ToArray();
        var queue = new Queue<int>(nodes.Where(node => indegree[node.Index] == 0).Select(node => node.Index));
        var topologicalOrder = new List<int>(nodes.Length);
        while (queue.Count != 0)
        {
            var nodeIndex = queue.Dequeue();
            topologicalOrder.Add(nodeIndex);
            foreach (var edgeIndex in nodes[nodeIndex].Outgoing)
            {
                var edge = edges[edgeIndex];
                nodes[edge.TargetIndex].Column = Math.Max(nodes[edge.TargetIndex].Column, nodes[nodeIndex].Column + 1);
                if (--indegree[edge.TargetIndex] == 0) queue.Enqueue(edge.TargetIndex);
            }
        }
        var maximumColumn = nodes.Max(node => node.Column);

        var nodeColors = sankey.GetProperty("nodeColors").EnumerateArray().Select(catalog.Color).ToArray();
        var nodeStroke = Property(sankey, "nodeStroke");
        var nodeWidth = OptionalDouble(sankey, "nodeWidth") ?? 12;
        var nodeGap = OptionalDouble(sankey, "nodeGap") ?? 9;
        var nodeAlign = OptionalString(sankey, "nodeAlign") ?? "justify";
        if (nodeAlign == "justify")
            foreach (var node in nodes.Where(node => node.Outgoing.Count == 0)) node.Column = maximumColumn;
        else if (nodeAlign == "right")
        {
            var distanceToSink = new int[nodes.Length];
            for (var orderIndex = topologicalOrder.Count - 1; orderIndex >= 0; orderIndex--)
            {
                var nodeIndex = topologicalOrder[orderIndex];
                foreach (var edgeIndex in nodes[nodeIndex].Outgoing)
                    distanceToSink[nodeIndex] = Math.Max(
                        distanceToSink[nodeIndex],
                        distanceToSink[edges[edgeIndex].TargetIndex] + 1);
            }
            foreach (var node in nodes) node.Column = maximumColumn - distanceToSink[node.Index];
        }
        var flowOpacity = OptionalDouble(sankey, "flowOpacity") ?? 0.45;
        var flowCurvature = OptionalDouble(sankey, "flowCurvature") ?? 0.7;
        var flowColorMode = OptionalString(sankey, "flowColorMode") ?? "source";
        var showValues = OptionalBoolean(sankey, "showValues") ?? false;
        var labelStyle = Property(sankey, "labelTextStyle");
        var valueStyle = Property(sankey, "valueTextStyle");
        var titleStyle = FirstProperty(inlineStyle, namedStyle, "titleTextStyle");
        var nodeColorMap = Property(sankey, "nodeColorMap")?.EnumerateObject()
            .ToDictionary(property => property.Name, property => catalog.Color(property.Value), StringComparer.Ordinal);
        foreach (var node in nodes)
            node.Color = nodeColorMap is not null && nodeColorMap.TryGetValue(node.Name, out var nodeColor)
                ? nodeColor
                : nodeColors[node.Index % nodeColors.Length];

        var x = element.Frame.X;
        var y = element.Frame.Y;
        var width = element.Frame.Width;
        var height = element.Frame.Height;
        var titleText = element.Title is null ? string.Empty : Flatten(element.Title);
        var titleHeight = titleText.Length == 0 ? 0 : Math.Clamp(height * 0.12, 22, 38);
        var plotX = x;
        var plotY = y + titleHeight + 4;
        var plotWidth = width;
        var plotHeight = height - titleHeight - 4;
        if (plotWidth < 260 || plotHeight < 130 || maximumColumn < 1)
            throw Unsupported(element.Id, "sankey frame or graph depth is too small for a readable native flow layout");

        var columns = nodes.GroupBy(node => node.Column).OrderBy(group => group.Key).ToArray();
        var scale = columns.Min(column =>
        {
            var available = plotHeight - nodeGap * (column.Count() - 1);
            var magnitude = column.Sum(node => node.Value);
            return available / magnitude;
        });
        if (!double.IsFinite(scale) || scale <= 0)
            throw Unsupported(element.Id, "sankey node gaps leave no room for positive flow thickness");
        var columnStep = (plotWidth - nodeWidth) / maximumColumn;
        if (columnStep - nodeWidth < 48)
            throw Unsupported(element.Id, "sankey columns leave insufficient horizontal room for native ribbons and labels");
        foreach (var column in columns)
        {
            var ordered = column.OrderBy(node => node.Index).ToArray();
            var occupied = ordered.Sum(node => node.Value * scale) + nodeGap * (ordered.Length - 1);
            var cursorY = plotY + (plotHeight - occupied) / 2;
            foreach (var node in ordered)
            {
                node.X = plotX + node.Column * columnStep;
                node.Y = cursorY;
                node.Height = node.Value * scale;
                cursorY += node.Height + nodeGap;
            }
        }

        var group = new PresentationGroup
        {
            LeftEmu = Emu(x),
            TopEmu = Emu(y),
            WidthEmu = Emu(width),
            HeightEmu = Emu(height),
            ChildLeftEmu = Emu(x),
            ChildTopEmu = Emu(y),
            ChildWidthEmu = Emu(width),
            ChildHeightEmu = Emu(height),
        };
        if (BuildFrameTransform(element.Frame) is { } frameTransform) group.FrameTransform = frameTransform;
        if (titleText.Length > 0)
            group.Children.Add(VectorChartTitleElement(
                SankeyNativeId(element.Id, "title"),
                "sankey title",
                x,
                y,
                width,
                titleHeight,
                raw.GetProperty("title"),
                titleStyle,
                catalog,
                Math.Clamp(titleHeight * 0.48, 11, 18),
                "left"));

        var outgoingOffsets = new double[nodes.Length];
        var incomingOffsets = new double[nodes.Length];
        foreach (var edge in edges)
        {
            var source = nodes[edge.SourceIndex];
            var target = nodes[edge.TargetIndex];
            var thickness = edge.Value * scale;
            var sourceTop = source.Y + outgoingOffsets[source.Index];
            var targetTop = target.Y + incomingOffsets[target.Index];
            var color = flowColorMode == "target" ? target.Color : source.Color;
            var ribbon = ShapeFrame(new PpjFrameModel(plotX, plotY, plotWidth, plotHeight, 0, false, false), "custom");
            ribbon.FillRgb = color.Rgb;
            ribbon.FillOpacityThousandthPercent = Opacity(flowOpacity * color.Alpha);
            ribbon.LineStyle = "none";
            ribbon.CustomPaths.Add(BuildSankeyRibbonPath(
                plotX,
                plotY,
                plotWidth,
                plotHeight,
                source.X + nodeWidth,
                sourceTop,
                sourceTop + thickness,
                target.X,
                targetTop,
                targetTop + thickness,
                flowCurvature));
            group.Children.Add(new PresentationElement
            {
                Id = SankeyNativeId(element.Id, $"flow/{edge.Index}"),
                Name = $"sankey flow {source.Name} to {target.Name}",
                Shape = ribbon,
            });
            outgoingOffsets[source.Index] += thickness;
            incomingOffsets[target.Index] += thickness;
        }

        var labelWidth = Math.Clamp(columnStep - nodeWidth - 8, 42, 120);
        foreach (var node in nodes)
        {
            var shape = ShapeFrame(new PpjFrameModel(node.X, node.Y, nodeWidth, node.Height, 0, false, false), "rect");
            shape.FillRgb = node.Color.Rgb;
            if (node.Color.Alpha < 1) shape.FillOpacityThousandthPercent = Opacity(node.Color.Alpha);
            if (nodeStroke is { } stroke) ApplyLine(shape, stroke, catalog);
            else shape.LineStyle = "none";
            group.Children.Add(new PresentationElement
            {
                Id = SankeyNativeId(element.Id, $"node/{node.Index}"),
                Name = $"sankey node {node.Name}",
                Shape = shape,
            });

            if (node.Height < 10) continue;
            var placeLeft = node.Column == maximumColumn;
            var textX = placeLeft ? node.X - labelWidth - 4 : node.X + nodeWidth + 4;
            var labelHeight = showValues && node.Height >= 24 ? Math.Min(14, node.Height * 0.48) : Math.Min(16, node.Height);
            group.Children.Add(VectorChartTextElement(
                SankeyNativeId(element.Id, $"label/{node.Index}"),
                $"sankey label {node.Name}",
                textX,
                node.Y,
                labelWidth,
                labelHeight,
                node.Name,
                labelStyle,
                catalog,
                Math.Clamp(Math.Min(labelHeight * 0.62, labelWidth / Math.Max(4, node.Name.Length) * 1.5), 6, 10),
                placeLeft ? "right" : "left",
                "16324F"));
            if (showValues && node.Height >= 24)
                group.Children.Add(VectorChartTextElement(
                    SankeyNativeId(element.Id, $"value/{node.Index}"),
                    $"sankey value {node.Name}",
                    textX,
                    node.Y + labelHeight,
                    labelWidth,
                    Math.Min(12, node.Height - labelHeight),
                    node.Value.ToString("0.##", CultureInfo.InvariantCulture),
                    valueStyle,
                    catalog,
                    Math.Clamp(node.Height * 0.18, 6, 9),
                    placeLeft ? "right" : "left",
                    "52606D"));
        }

        ApplyAccessibility(group, element.Accessibility);
        return group;
    }

    private static void ValidateSankeyCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        var nodeCount = element.Data.Categories.Count;
        if (nodeCount is < 2 or > 64 ||
            element.Data.Categories.Any(category => category.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(category.GetString())) ||
            element.Data.Categories.Select(category => category.GetString()).Distinct(StringComparer.Ordinal).Count() != nodeCount)
            throw Unsupported(element.Id, "sankey categories must be 2..64 unique non-empty node names");
        if (element.Data.Series.Count != 1)
            throw Unsupported(element.Id, "sankey charts require exactly one directed-flow series");
        var series = element.Data.Series[0];
        var edgeCount = series.Values.Count;
        if (edgeCount is < 1 or > 256 || series.Values.Any(value => value is null || value <= 0) ||
            series.Sources.Count != edgeCount || series.Targets.Count != edgeCount)
            throw Unsupported(element.Id, "sankey sources, targets, and positive flow values must align across 1..256 edges");
        foreach (var property in new[] { "pointRoles", "xValues", "bubbleSizes", "openValues", "highValues", "lowValues", "parents", "chartType", "axis", "color", "fill", "stroke", "marker", "trendlines", "errorBars" })
            if (series.Raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, $"sankey series do not support {property}");
        foreach (var property in new[] { "xAxis", "yAxis", "secondaryXAxis", "secondaryYAxis" })
            if (raw.TryGetProperty(property, out _))
                throw Unsupported(element.Id, "sankey charts do not use Cartesian axes");
        foreach (var property in new[] { "legend", "stacking", "gapWidth", "startAngle", "holeSize", "bubbleScale", "bubbleSizeMode", "showCategoryAxis", "showValueAxis", "showGridlines", "showDataLabels", "dataLabelPosition", "dataLabels", "chartAreaFill", "plotAreaFill", "legendTextStyle", "smooth", "varyColors", "waterfall", "heatmap", "candlestick", "treemap", "sunburst" })
            if (FirstProperty(inlineStyle, namedStyle, property) is not null)
                throw Unsupported(element.Id, $"sankey charts do not support chart style field {property}");
        if (FirstProperty(inlineStyle, namedStyle, "sankey") is null)
            throw Unsupported(element.Id, "sankey charts require style.sankey");

        var names = element.Data.Categories.Select(category => category.GetString()!).ToArray();
        var indexes = names.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.Ordinal);
        if (FirstProperty(inlineStyle, namedStyle, "sankey") is { } sankeyStyle &&
            Property(sankeyStyle, "nodeColorMap") is { } nodeColorMap)
            foreach (var property in nodeColorMap.EnumerateObject())
                if (!indexes.ContainsKey(property.Name))
                    throw Unsupported(element.Id, $"sankey nodeColorMap names undeclared node {property.Name}");
        var indegree = new int[nodeCount];
        var outgoing = Enumerable.Range(0, nodeCount).Select(_ => new List<int>()).ToArray();
        var incomingFlow = new double[nodeCount];
        var outgoingFlow = new double[nodeCount];
        var used = new bool[nodeCount];
        var edgeKeys = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < edgeCount; index++)
        {
            if (!indexes.TryGetValue(series.Sources[index], out var source) || !indexes.TryGetValue(series.Targets[index], out var target))
                throw Unsupported(element.Id, "every sankey source and target must name a declared category");
            if (source == target) throw Unsupported(element.Id, "a sankey edge cannot target its source node");
            if (!edgeKeys.Add(series.Sources[index] + "\u0000" + series.Targets[index]))
                throw Unsupported(element.Id, "duplicate sankey endpoints must be combined into one explicit flow value");
            outgoing[source].Add(target);
            indegree[target]++;
            used[source] = true;
            used[target] = true;
            outgoingFlow[source] += series.Values[index]!.Value;
            incomingFlow[target] += series.Values[index]!.Value;
        }
        if (used.Any(value => !value)) throw Unsupported(element.Id, "every declared sankey node must be incident to an edge");

        var queue = new Queue<int>(Enumerable.Range(0, nodeCount).Where(index => indegree[index] == 0));
        var visited = 0;
        while (queue.Count != 0)
        {
            var source = queue.Dequeue();
            visited++;
            foreach (var target in outgoing[source]) if (--indegree[target] == 0) queue.Enqueue(target);
        }
        if (visited != nodeCount) throw Unsupported(element.Id, "sankey edges must form a directed acyclic graph");
        for (var index = 0; index < nodeCount; index++)
            if (incomingFlow[index] > 0 && outgoingFlow[index] > 0)
            {
                var tolerance = Math.Max(1e-9, Math.Max(incomingFlow[index], outgoingFlow[index]) * 1e-9);
                if (Math.Abs(incomingFlow[index] - outgoingFlow[index]) > tolerance)
                    throw Unsupported(element.Id, $"sankey internal node {names[index]} must conserve flow");
            }
    }

    private static PresentationCustomGeometryPath BuildSankeyRibbonPath(
        double plotX,
        double plotY,
        double plotWidth,
        double plotHeight,
        double sourceX,
        double sourceTop,
        double sourceBottom,
        double targetX,
        double targetTop,
        double targetBottom,
        double curvature)
    {
        const long viewport = 100_000;
        PresentationCustomGeometryPoint Point(double x, double y) => new()
        {
            X = checked((long)Math.Round((x - plotX) / plotWidth * viewport, MidpointRounding.AwayFromZero)),
            Y = checked((long)Math.Round((y - plotY) / plotHeight * viewport, MidpointRounding.AwayFromZero)),
        };
        var controlFraction = curvature * 0.5;
        var control1X = sourceX + (targetX - sourceX) * controlFraction;
        var control2X = targetX - (targetX - sourceX) * controlFraction;
        var path = new PresentationCustomGeometryPath
        {
            Width = viewport,
            Height = viewport,
            FillMode = PresentationCustomGeometryPath.Types.FillMode.Normal,
            Stroke = false,
        };
        path.Commands.Add(MoveTo(Point(sourceX, sourceTop)));
        path.Commands.Add(new PresentationCustomGeometryCommand
        {
            CubicBezierTo = new PresentationCustomGeometryCubicBezier
            {
                Control1 = Point(control1X, sourceTop),
                Control2 = Point(control2X, targetTop),
                End = Point(targetX, targetTop),
            },
        });
        path.Commands.Add(LineTo(Point(targetX, targetBottom)));
        path.Commands.Add(new PresentationCustomGeometryCommand
        {
            CubicBezierTo = new PresentationCustomGeometryCubicBezier
            {
                Control1 = Point(control2X, targetBottom),
                Control2 = Point(control1X, sourceBottom),
                End = Point(sourceX, sourceBottom),
            },
        });
        path.Commands.Add(new PresentationCustomGeometryCommand { Close = true });
        return path;
    }

    private static string SankeyNativeId(string elementId, string suffix) =>
        $"{elementId}/sankey/{suffix}";

    private sealed class SankeyNode(int index, string name)
    {
        internal int Index { get; } = index;
        internal string Name { get; } = name;
        internal List<int> Incoming { get; } = [];
        internal List<int> Outgoing { get; } = [];
        internal double IncomingValue { get; set; }
        internal double OutgoingValue { get; set; }
        internal double Value => Math.Max(IncomingValue, OutgoingValue);
        internal int Column { get; set; }
        internal double X { get; set; }
        internal double Y { get; set; }
        internal double Height { get; set; }
        internal (string Rgb, double Alpha) Color { get; set; }
    }

    private sealed record SankeyEdge(int Index, int SourceIndex, int TargetIndex, double Value);

    private static void ValidateWaterfallCompileProfile(
        PpjChartElementModel element,
        JsonElement raw,
        JsonElement? namedStyle,
        JsonElement? inlineStyle)
    {
        if (element.Data.Series.Count != 1)
            throw Unsupported(element.Id, "waterfall charts require exactly one semantic series");
        var series = raw.GetProperty("data").GetProperty("series")[0];
        RejectProperties(
            series,
            element.Id,
            "chartType",
            "axis",
            "color",
            "fill",
            "stroke",
            "marker",
            "trendlines",
            "errorBars",
            "xValues",
            "bubbleSizes");
        foreach (var name in new[] { "stacking", "showDataLabels", "dataLabelPosition", "dataLabels", "smooth", "varyColors" })
            if (FirstProperty(inlineStyle, namedStyle, name) is not null)
                throw Unsupported(element.Id, $"{name} is outside the bounded waterfall style profile");
        if (FirstProperty(inlineStyle, namedStyle, "legend") is { } legend && legend.GetString() != "none")
            throw Unsupported(element.Id, "waterfall charts do not expose their internal lowering series through a legend");
        if (FirstProperty(inlineStyle, namedStyle, "waterfall") is not { } waterfall)
            throw Unsupported(element.Id, "waterfall charts require style.waterfall increase, decrease, and total role styles");
        foreach (var role in new[] { "increase", "decrease", "total" })
        {
            var fillType = waterfall.GetProperty(role).GetProperty("fill").GetProperty("type").GetString();
            if (fillType is "none" or "image")
                throw Unsupported(element.Id, $"waterfall {role} fill must be solid or a bounded gradient");
        }
    }

    private static void BuildWaterfallSeries(
        PresentationChart chart,
        PpjChartElementModel element,
        JsonElement? namedStyle,
        JsonElement? inlineStyle,
        Catalog catalog)
    {
        var semantic = element.Data.Series[0];
        if (semantic.PointRoles.Count != semantic.Values.Count || semantic.Values.Any(value => value is null))
            throw Unsupported(element.Id, "waterfall values and pointRoles must be complete and aligned");

        var count = semantic.Values.Count;
        var offset = new double[count];
        var increase = new double[count];
        var decrease = new double[count];
        var total = new double[count];
        var increaseMissing = new List<uint>(count);
        var decreaseMissing = new List<uint>(count);
        var totalMissing = new List<uint>(count);
        var running = 0d;
        for (var index = 0; index < count; index++)
        {
            var value = semantic.Values[index]!.Value;
            if (semantic.PointRoles[index] == "total")
            {
                offset[index] = 0;
                total[index] = value;
                increaseMissing.Add(checked((uint)index));
                decreaseMissing.Add(checked((uint)index));
                running = value;
                continue;
            }

            var next = running + value;
            offset[index] = Math.Min(running, next);
            if (value >= 0)
            {
                increase[index] = value;
                decreaseMissing.Add(checked((uint)index));
            }
            else
            {
                decrease[index] = -value;
                increaseMissing.Add(checked((uint)index));
            }
            totalMissing.Add(checked((uint)index));
            running = next;
        }

        var waterfall = FirstProperty(inlineStyle, namedStyle, "waterfall")!.Value;
        chart.Grouping = "stacked";
        if (!chart.HasGapWidth) chart.GapWidth = 60;
        chart.HasLegend = false;
        chart.LegendPosition = string.Empty;
        chart.DataLabels = null;
        chart.XAxis ??= new SpreadsheetChartAxisArtifact();
        chart.YAxis ??= new SpreadsheetChartAxisArtifact();
        if (!chart.YAxis.HasMinimum) chart.YAxis.Minimum = 0;

        var offsetSeries = new SpreadsheetChartSeriesArtifact
        {
            Name = "__offset__",
            SeriesFill = new SpreadsheetChartSurfaceFill { NoFill = true },
            Line = new SpreadsheetChartLineStyleArtifact
            {
                Color = new SpreadsheetColor { Rgb = "000000" },
                WidthPoints = 0,
                OpacityThousandthPercent = 0,
            },
        };
        offsetSeries.Values.Add(offset);
        chart.Series.Add(offsetSeries);
        chart.Series.Add(WaterfallRoleSeries(
            waterfall.GetProperty("increase"), increase, increaseMissing, catalog, element.Id, "increase"));
        chart.Series.Add(WaterfallRoleSeries(
            waterfall.GetProperty("decrease"), decrease, decreaseMissing, catalog, element.Id, "decrease"));
        chart.Series.Add(WaterfallRoleSeries(
            waterfall.GetProperty("total"), total, totalMissing, catalog, element.Id, "total"));
    }

    private static SpreadsheetChartSeriesArtifact WaterfallRoleSeries(
        JsonElement role,
        IReadOnlyList<double> values,
        IEnumerable<uint> missingIndexes,
        Catalog catalog,
        string elementId,
        string roleName)
    {
        var output = new SpreadsheetChartSeriesArtifact
        {
            Name = role.GetProperty("label").GetString()!,
            SeriesFill = BuildChartFill(role.GetProperty("fill"), catalog, $"{elementId} waterfall {roleName} fill"),
        };
        output.Values.Add(values);
        output.MissingValueIndexes.Add(missingIndexes);
        if (role.TryGetProperty("stroke", out var stroke)) output.Line = BuildChartLine(stroke, catalog);
        return output;
    }

    private static SpreadsheetChartSeriesArtifact BuildSeries(
        PpjChartSeriesModel source,
        JsonElement raw,
        Catalog catalog,
        SpreadsheetChartType chartType)
    {
        if (raw.TryGetProperty("fill", out _) && raw.TryGetProperty("color", out _))
            throw Unsupported(source.Id, "chart-series color and fill are aliases and cannot both be present");
        var series = new SpreadsheetChartSeriesArtifact { Name = source.Name };
        for (var index = 0; index < source.Values.Count; index++)
        {
            var value = source.Values[index];
            if (value is null)
            {
                series.Values.Add(0d);
                series.MissingValueIndexes.Add(checked((uint)index));
            }
            else series.Values.Add(value.Value);
        }
        series.XValues.Add(source.XValues);
        series.BubbleSizes.Add(source.BubbleSizes);
        if (raw.TryGetProperty("fill", out var fill))
        {
            var chartFill = BuildChartFill(fill, catalog, $"{source.Id} series fill");
            if (chartFill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.SolidRgb && !chartFill.HasOpacityThousandthPercent)
                series.Fill = new SpreadsheetColor { Rgb = chartFill.SolidRgb };
            else
                series.SeriesFill = chartFill;
        }
        else if (raw.TryGetProperty("color", out var color))
        {
            var resolved = catalog.Color(color);
            if (resolved.Alpha == 1)
                series.Fill = new SpreadsheetColor { Rgb = resolved.Rgb };
            else
                series.SeriesFill = new SpreadsheetChartSurfaceFill
                {
                    SolidRgb = resolved.Rgb,
                    OpacityThousandthPercent = Opacity(resolved.Alpha),
                };
        }
        if (raw.TryGetProperty("stroke", out var stroke))
            series.Line = BuildChartLine(stroke, catalog);
        if (raw.TryGetProperty("pointStyles", out var pointStyles))
        {
            foreach (var point in pointStyles.EnumerateArray())
            {
                var output = new SpreadsheetChartPointStyleArtifact
                {
                    Index = checked((uint)point.GetProperty("index").GetInt32()),
                };
                if (point.TryGetProperty("fill", out var pointFill))
                    output.Fill = BuildChartFill(pointFill, catalog, $"{source.Id} point {output.Index} fill");
                if (point.TryGetProperty("stroke", out var pointStroke))
                    output.Line = BuildChartLine(pointStroke, catalog);
                if (point.TryGetProperty("explosion", out var explosion))
                    output.Explosion = checked((uint)explosion.GetInt32());
                series.PointStyles.Add(output);
            }
        }
        if (raw.TryGetProperty("marker", out var marker))
            series.Marker = BuildChartMarker(marker, catalog);
        if (raw.TryGetProperty("trendlines", out var trendlines))
        {
            if (chartType is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line))
                throw Unsupported(source.Id, "trendlines require a bar, column, or line series");
            series.Trendlines.Add(trendlines.EnumerateArray().Select(item => BuildChartTrendline(item, catalog)));
        }
        if (raw.TryGetProperty("errorBars", out var errorBars))
        {
            if (chartType is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line))
                throw Unsupported(source.Id, "error bars require a bar, column, or line series");
            series.ErrorBars = BuildChartErrorBars(errorBars, catalog);
        }
        if (raw.TryGetProperty("dataLabels", out var dataLabels))
            series.DataLabels = BuildSeriesDataLabels(dataLabels, catalog);
        return series;
    }

    private static SpreadsheetChartSeriesDataLabelsArtifact BuildSeriesDataLabels(JsonElement source, Catalog catalog)
    {
        var output = new SpreadsheetChartSeriesDataLabelsArtifact();
        if (HasChartLabelFields(source)) output.Defaults = BuildChartLabelOverride(source, catalog);
        if (source.TryGetProperty("points", out var points))
        {
            foreach (var point in points.EnumerateArray())
                output.Points.Add(new SpreadsheetChartPointDataLabelArtifact
                {
                    Index = checked((uint)point.GetProperty("index").GetInt32()),
                    Override = BuildChartLabelOverride(point, catalog),
                });
        }
        return output;
    }

    private static SpreadsheetChartDataLabelOverrideArtifact BuildChartLabelOverride(JsonElement source, Catalog catalog)
    {
        var output = new SpreadsheetChartDataLabelOverrideArtifact();
        if (source.TryGetProperty("showValue", out var showValue))
            output.ShowValue = catalog.BooleanToken(showValue, "boolean", "chart data-label showValue");
        if (source.TryGetProperty("showCategory", out var showCategory))
            output.ShowCategoryName = catalog.BooleanToken(showCategory, "boolean", "chart data-label showCategory");
        if (source.TryGetProperty("showSeries", out var showSeries))
            output.ShowSeriesName = catalog.BooleanToken(showSeries, "boolean", "chart data-label showSeries");
        if (source.TryGetProperty("showPercent", out var showPercent))
            output.ShowPercent = catalog.BooleanToken(showPercent, "boolean", "chart data-label showPercent");
        if (source.TryGetProperty("position", out var position))
            output.Position = LabelPosition(ChartEnumToken(position, catalog, "chart data-label position",
                "best-fit", "bottom", "center", "inside-base", "inside-end", "left", "outside-end", "right", "top"));
        if (source.TryGetProperty("numberFormat", out var numberFormat))
            output.NumberFormatCode = catalog.StringToken(numberFormat, "string", "chart data-label numberFormat");
        if (source.TryGetProperty("textStyle", out var textStyle)) output.TextStyle = BuildChartTextStyle(textStyle, catalog);
        return output;
    }

    private static bool HasChartLabelFields(JsonElement source) =>
        source.TryGetProperty("showValue", out _) || source.TryGetProperty("showCategory", out _) ||
        source.TryGetProperty("showSeries", out _) || source.TryGetProperty("showPercent", out _) ||
        source.TryGetProperty("position", out _) || source.TryGetProperty("numberFormat", out _) ||
        source.TryGetProperty("textStyle", out _);

    private static SpreadsheetChartAxisArtifact BuildChartAxis(JsonElement source, Catalog catalog)
    {
        var axis = new SpreadsheetChartAxisArtifact
        {
            Title = source.TryGetProperty("title", out var title)
                ? catalog.StringToken(title, "string", "chart axis title")
                : string.Empty,
            NumberFormatCode = source.TryGetProperty("numberFormat", out var numberFormat)
                ? catalog.StringToken(numberFormat, "string", "chart axis numberFormat")
                : string.Empty,
        };
        if (source.TryGetProperty("tickLabelInterval", out var tickLabelInterval))
            axis.TickLabelInterval = AxisInteger(tickLabelInterval, catalog, "chart axis tickLabelInterval", 1, 10_000);
        if (source.TryGetProperty("min", out var minimum))
            axis.Minimum = catalog.NumberToken(minimum, "size", "chart axis min");
        if (source.TryGetProperty("max", out var maximum))
            axis.Maximum = catalog.NumberToken(maximum, "size", "chart axis max");
        if (source.TryGetProperty("majorUnit", out var majorUnit))
            axis.MajorUnit = catalog.PositiveNumberToken(majorUnit, "size", "chart axis majorUnit");
        if (source.TryGetProperty("visible", out var visible))
            axis.Visible = catalog.BooleanToken(visible, "boolean", "chart axis visible");
        if (source.TryGetProperty("tickLabelsVisible", out var tickLabelsVisible))
            axis.TickLabelsVisible = catalog.BooleanToken(tickLabelsVisible, "boolean", "chart axis tickLabelsVisible");
        if (source.TryGetProperty("reverse", out var reverse))
            axis.Reverse = catalog.BooleanToken(reverse, "boolean", "chart axis reverse");
        if (source.TryGetProperty("axisLine", out var axisLine))
        {
            if (axisLine.ValueKind is JsonValueKind.True or JsonValueKind.False || IsTokenReference(axisLine))
                axis.AxisLineVisible = catalog.BooleanToken(axisLine, "boolean", "chart axis axisLine");
            else
            {
                axis.AxisLineVisible = true;
                axis.AxisLine = BuildChartLine(axisLine, catalog);
            }
        }
        if (source.TryGetProperty("axisLineArrow", out var axisLineArrow))
        {
            if (source.TryGetProperty("axisLine", out axisLine) && axisLine.ValueKind == JsonValueKind.False)
                throw Unsupported("chart axis", "axisLineArrow requires a visible axis line");
            axis.AxisLineVisible = true;
            axis.AxisLine ??= new SpreadsheetChartLineStyleArtifact();
            ApplyChartAxisArrows(axis.AxisLine, axisLineArrow);
        }
        if (source.TryGetProperty("gridLine", out var gridLine))
        {
            if (gridLine.ValueKind is JsonValueKind.True or JsonValueKind.False || IsTokenReference(gridLine))
            {
                axis.ShowMajorGridlines = true;
                if (!catalog.BooleanToken(gridLine, "boolean", "chart axis gridLine"))
                    axis.MajorGridlineVisible = false;
            }
            else
            {
                axis.ShowMajorGridlines = true;
                axis.MajorGridlineStyle = BuildChartLine(gridLine, catalog);
            }
        }
        if (source.TryGetProperty("textStyle", out var textStyle))
            axis.TextStyle = BuildChartTextStyle(textStyle, catalog);
        if (source.TryGetProperty("titleTextStyle", out var titleTextStyle))
        {
            if (axis.Title.Length == 0) throw Unsupported("chart axis", "titleTextStyle requires a non-empty axis title");
            axis.TitleTextStyle = BuildChartTextStyle(titleTextStyle, catalog);
        }
        return axis;
    }

    private static void ApplyChartAxisArrows(SpreadsheetChartLineStyleArtifact line, JsonElement source)
    {
        line.StartArrow = OptionalString(source, "start") ?? string.Empty;
        line.EndArrow = OptionalString(source, "end") ?? string.Empty;
    }

    private static (SpreadsheetChartAxisArtifact XAxis, SpreadsheetChartAxisArtifact YAxis) BuildRadarSpokeAxes(
        JsonElement source,
        Catalog catalog)
    {
        var xAxis = new SpreadsheetChartAxisArtifact();
        var yAxis = new SpreadsheetChartAxisArtifact();
        var show = !source.TryGetProperty("show", out var showValue) ||
            catalog.BooleanToken(showValue, "boolean", "radar spokeAxis show");
        xAxis.Visible = show;
        yAxis.Visible = show;

        if (source.TryGetProperty("min", out var minimum))
            yAxis.Minimum = catalog.NumberToken(minimum, "size", "radar spokeAxis min");
        if (source.TryGetProperty("max", out var maximum))
            yAxis.Maximum = catalog.NumberToken(maximum, "size", "radar spokeAxis max");
        if (source.TryGetProperty("majorUnit", out var majorUnit))
            yAxis.MajorUnit = catalog.PositiveNumberToken(majorUnit, "size", "radar spokeAxis majorUnit");

        if (!show)
        {
            yAxis.TickLabelsVisible = false;
            return (xAxis, yAxis);
        }

        if (!source.TryGetProperty("label", out var label))
            yAxis.TickLabelsVisible = true;
        else if (label.ValueKind is JsonValueKind.True or JsonValueKind.False || IsTokenReference(label))
            yAxis.TickLabelsVisible = catalog.BooleanToken(label, "boolean", "radar spokeAxis label");
        else
        {
            yAxis.TickLabelsVisible = true;
            if (label.TryGetProperty("numberFormat", out var numberFormat))
                yAxis.NumberFormatCode = catalog.StringToken(numberFormat, "string", "radar spoke label numberFormat");
            yAxis.TextStyle = BuildChartTextStyle(label, catalog);
        }

        ApplyRadarGuideLine(xAxis, source, "axisLine", catalog);
        ApplyRadarGuideLine(yAxis, source, "gridLine", catalog);
        return (xAxis, yAxis);
    }

    private static void ApplyRadarGuideLine(
        SpreadsheetChartAxisArtifact axis,
        JsonElement source,
        string propertyName,
        Catalog catalog)
    {
        axis.ShowMajorGridlines = true;
        if (!source.TryGetProperty(propertyName, out var line)) return;
        if (line.ValueKind is JsonValueKind.True or JsonValueKind.False || IsTokenReference(line))
        {
            if (!catalog.BooleanToken(line, "boolean", $"radar spokeAxis {propertyName}"))
                axis.MajorGridlineVisible = false;
            return;
        }
        axis.MajorGridlineStyle = BuildChartLine(line, catalog);
    }

    private static uint AxisInteger(JsonElement value, Catalog catalog, string owner, uint minimum, uint maximum)
    {
        var number = catalog.NumberToken(value, "size", owner);
        if (!double.IsFinite(number) || number < minimum || number > maximum || Math.Truncate(number) != number)
            throw Unsupported(owner, $"must resolve to an integer between {minimum} and {maximum}");
        return checked((uint)number);
    }

    private static bool IsTokenReference(JsonElement value) =>
        value.ValueKind == JsonValueKind.Object && value.TryGetProperty("token", out var token) &&
        token.ValueKind == JsonValueKind.String;

    private static void RejectTokenizedVectorAxis(JsonElement? axis, string elementId, string name)
    {
        if (axis is not { ValueKind: JsonValueKind.Object } value) return;
        foreach (var property in value.EnumerateObject())
            if (IsTokenReference(property.Value))
                throw Unsupported(elementId, $"generated {name} axis cannot resolve tokenized {property.Name}; use a native ChartPart chart for typed axis tokens");
    }

    private static SpreadsheetChartMarkerArtifact BuildChartMarker(JsonElement source, Catalog catalog)
    {
        if (source.ValueKind == JsonValueKind.String)
            return new SpreadsheetChartMarkerArtifact { Symbol = Marker(source.GetString()!) };
        var marker = new SpreadsheetChartMarkerArtifact
        {
            Symbol = Marker(OptionalString(source, "symbol") ?? "circle"),
        };
        if (source.TryGetProperty("size", out var size)) marker.Size = checked((uint)size.GetInt32());
        if (source.TryGetProperty("fill", out var fill))
        {
            var color = catalog.Color(fill);
            marker.Fill = new SpreadsheetColor { Rgb = color.Rgb };
            if (color.Alpha < 1) marker.FillOpacityThousandthPercent = Opacity(color.Alpha);
        }
        if (source.TryGetProperty("stroke", out var stroke)) marker.Line = BuildChartLine(stroke, catalog);
        return marker;
    }

    private static SpreadsheetChartTrendlineArtifact BuildChartTrendline(JsonElement source, Catalog catalog)
    {
        var trendline = new SpreadsheetChartTrendlineArtifact
        {
            Type = TrendlineType(source.GetProperty("type").GetString()!),
            Name = OptionalString(source, "name") ?? string.Empty,
            DisplayEquation = source.TryGetProperty("displayEquation", out var equation) && equation.GetBoolean(),
            DisplayRSquared = source.TryGetProperty("displayRSquared", out var rSquared) && rSquared.GetBoolean(),
        };
        if (source.TryGetProperty("order", out var order)) trendline.PolynomialOrder = checked((uint)order.GetInt32());
        if (source.TryGetProperty("period", out var period)) trendline.Period = checked((uint)period.GetInt32());
        if (source.TryGetProperty("forward", out var forward)) trendline.Forward = forward.GetDouble();
        if (source.TryGetProperty("backward", out var backward)) trendline.Backward = backward.GetDouble();
        if (source.TryGetProperty("intercept", out var intercept)) trendline.Intercept = intercept.GetDouble();
        if (source.TryGetProperty("stroke", out var stroke)) trendline.Line = BuildChartLine(stroke, catalog);
        return trendline;
    }

    private static SpreadsheetChartErrorBarsArtifact BuildChartErrorBars(JsonElement source, Catalog catalog)
    {
        var errorBars = new SpreadsheetChartErrorBarsArtifact
        {
            Direction = ErrorBarDirection(OptionalString(source, "direction") ?? "y"),
            Type = ErrorBarType(OptionalString(source, "type") ?? "both"),
            ValueType = ErrorBarValueType(source.GetProperty("valueType").GetString()!),
            NoEndCap = source.TryGetProperty("noEndCap", out var noEndCap) && noEndCap.GetBoolean(),
        };
        if (source.TryGetProperty("value", out var value)) errorBars.Value = value.GetDouble();
        if (source.TryGetProperty("stroke", out var stroke)) errorBars.Line = BuildChartLine(stroke, catalog);
        return errorBars;
    }

    internal static SpreadsheetChartLineStyleArtifact BuildChartLine(JsonElement source, Catalog catalog)
    {
        var color = catalog.Color(source.GetProperty("color"));
        var line = new SpreadsheetChartLineStyleArtifact
        {
            Color = new SpreadsheetColor { Rgb = color.Rgb },
            DashStyle = ChartDash(OptionalString(source, "dash")),
            Cap = OptionalString(source, "cap") ?? string.Empty,
            Join = OptionalString(source, "join") ?? string.Empty,
        };
        if (source.TryGetProperty("width", out _)) line.WidthPoints = StrokeWidth(source, catalog, "chart line width");
        var opacity = OptionalOpacity(source, "opacity", catalog, "chart line opacity") ?? color.Alpha;
        if (opacity < 1) line.OpacityThousandthPercent = Opacity(opacity);
        return line;
    }

    }
}
