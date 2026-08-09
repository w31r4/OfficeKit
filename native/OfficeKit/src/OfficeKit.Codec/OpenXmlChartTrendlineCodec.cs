using System.Globalization;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns the bounded c:trendline projection shared by XLSX and PPTX ChartSpace.
// Package adapters decide source identity and topology; this leaf owns the
// semantic series children and rejects labels/extensions/complex line graphs.
internal static class OpenXmlChartTrendlineCodec
{
    private const int MaxTrendlines = 16;
    private const double MaxForecast = 1_000_000;
    private const double MaxSafeInteger = 9_007_199_254_740_991;
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly IReadOnlyDictionary<XName, int> ChildOrder = new Dictionary<XName, int>
    {
        [ChartNs + "name"] = 0,
        [ChartNs + "spPr"] = 1,
        [ChartNs + "trendlineType"] = 2,
        [ChartNs + "order"] = 3,
        [ChartNs + "period"] = 4,
        [ChartNs + "forward"] = 5,
        [ChartNs + "backward"] = 6,
        [ChartNs + "intercept"] = 7,
        [ChartNs + "dispRSqr"] = 8,
        [ChartNs + "dispEq"] = 9,
    };

    internal static void Validate(SpreadsheetChartSeriesArtifact series, SpreadsheetChartType chartType, string worksheetId, string chartId)
    {
        if (series.Trendlines.Count == 0) return;
        if (chartType is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line))
            throw Invalid(worksheetId, chartId, series.Name, "trendlines require a bar or line series");
        if (series.Trendlines.Count > MaxTrendlines)
            throw Invalid(worksheetId, chartId, series.Name, $"supports at most {MaxTrendlines} trendlines");
        for (var index = 0; index < series.Trendlines.Count; index++)
        {
            var item = series.Trendlines[index];
            if (!Supported(item.Type)) throw Invalid(worksheetId, chartId, series.Name, $"trendline {index + 1} has an unsupported type");
            if (item.Name.Length > 255 || HasControls(item.Name)) throw Invalid(worksheetId, chartId, series.Name, $"trendline {index + 1} name must contain at most 255 characters without controls");
            if (item.Type == SpreadsheetChartTrendlineType.Polynomial)
            {
                if (!item.HasPolynomialOrder || item.PolynomialOrder is < 2 or > 6) throw Invalid(worksheetId, chartId, series.Name, $"trendline {index + 1} polynomial order must be from 2 through 6");
            }
            else if (item.HasPolynomialOrder) throw Invalid(worksheetId, chartId, series.Name, $"trendline {index + 1} order is valid only for polynomial trendlines");
            if (item.Type == SpreadsheetChartTrendlineType.MovingAverage)
            {
                var maximum = Math.Min(255, series.Values.Count - 1);
                if (!item.HasPeriod || maximum < 2 || item.Period < 2 || item.Period > maximum) throw Invalid(worksheetId, chartId, series.Name, $"trendline {index + 1} moving-average period must be from 2 through {Math.Max(1, maximum)}");
            }
            else if (item.HasPeriod) throw Invalid(worksheetId, chartId, series.Name, $"trendline {index + 1} period is valid only for moving-average trendlines");
            ValidateForecast(item.HasForward, item.Forward, "forward", worksheetId, chartId, series.Name, index);
            ValidateForecast(item.HasBackward, item.Backward, "backward", worksheetId, chartId, series.Name, index);
            if (item.HasIntercept && (!double.IsFinite(item.Intercept) || Math.Abs(item.Intercept) > MaxSafeInteger))
                throw Invalid(worksheetId, chartId, series.Name, $"trendline {index + 1} intercept must be finite and within the JavaScript safe-integer magnitude");
            XlsxChartSeriesLineStyleCodec.ValidateLine(item.Line, worksheetId, chartId, series.Name, $"trendline {index + 1} line");
        }
    }

    internal static bool TryRead(XElement nativeSeries, SpreadsheetChartSeriesArtifact series, SpreadsheetChartType chartType)
    {
        var native = nativeSeries.Elements(ChartNs + "trendline").ToArray();
        if (native.Length == 0) return true;
        if (chartType is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line) || native.Length > MaxTrendlines) return false;
        var parsed = new List<SpreadsheetChartTrendlineArtifact>();
        foreach (var item in native)
        {
            if (!TryRead(item, series.Values.Count, out var trendline)) return false;
            parsed.Add(trendline);
        }
        series.Trendlines.Add(parsed);
        return true;
    }

    internal static IEnumerable<XElement> Elements(IEnumerable<SpreadsheetChartTrendlineArtifact> trendlines) => trendlines.Select(Element);

    internal static void Patch(XElement nativeSeries, SpreadsheetChartSeriesArtifact target, string errorCode, string subject)
    {
        var native = nativeSeries.Elements(ChartNs + "trendline").ToArray();
        if (native.Length != target.Trendlines.Count) throw Topology(errorCode, subject, "trendline topology changed unexpectedly");
        for (var index = 0; index < native.Length; index++)
        {
            if (!TryRead(native[index], target.Values.Count, out var original)) throw Topology(errorCode, subject, $"trendline {index + 1} no longer matches the editable profile");
            if (Semantics(original).Equals(Semantics(target.Trendlines[index]), StringComparison.Ordinal)) continue;
            native[index].ReplaceWith(Element(target.Trendlines[index]));
        }
    }

    internal static string Semantics(IEnumerable<SpreadsheetChartTrendlineArtifact> trendlines) =>
        string.Join('\u001d', trendlines.Select(Semantics));

    private static bool TryRead(XElement source, int valueCount, out SpreadsheetChartTrendlineArtifact trendline)
    {
        trendline = new SpreadsheetChartTrendlineArtifact();
        if (source.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || HasUnexpectedNodes(source)) return false;
        var children = source.Elements().ToArray();
        var seen = new HashSet<XName>();
        var order = -1;
        foreach (var child in children)
        {
            if (!ChildOrder.TryGetValue(child.Name, out var current) || current < order || !seen.Add(child.Name)) return false;
            order = current;
        }
        var typeElement = source.Element(ChartNs + "trendlineType");
        if (!TryScalar(typeElement, out var typeValue) || !TryType(typeValue, out var type)) return false;
        trendline.Type = type;

        var name = source.Element(ChartNs + "name");
        if (name is not null)
        {
            if (!TryText(name, out var value) || value.Length is < 1 or > 255 || HasControls(value)) return false;
            trendline.Name = value;
        }
        var shape = source.Element(ChartNs + "spPr");
        if (shape is not null)
        {
            if (shape.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || shape.Elements().Any(element => element.Name != DrawingNs + "ln") ||
                HasUnexpectedNodes(shape) ||
                !XlsxChartSeriesLineStyleCodec.TryReadLine(shape, out var line)) return false;
            if (line is not null) trendline.Line = line;
        }

        var orderElement = source.Element(ChartNs + "order");
        if (type == SpreadsheetChartTrendlineType.Polynomial)
        {
            if (orderElement is null) trendline.PolynomialOrder = 2;
            else if (!TryScalar(orderElement, out var value) || !uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed) || parsed is < 2 or > 6) return false;
            else trendline.PolynomialOrder = parsed;
        }
        else if (orderElement is not null) return false;

        var periodElement = source.Element(ChartNs + "period");
        if (type == SpreadsheetChartTrendlineType.MovingAverage)
        {
            var maximum = Math.Min(255, valueCount - 1);
            if (maximum < 2) return false;
            if (periodElement is null) trendline.Period = 2;
            else if (!TryScalar(periodElement, out var value) || !uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed) || parsed < 2 || parsed > maximum) return false;
            else trendline.Period = parsed;
        }
        else if (periodElement is not null) return false;

        if (!TryOptionalDouble(source, "forward", 0, MaxForecast, halfIncrement: true, out var forward)) return false;
        if (forward.HasValue) trendline.Forward = forward.Value;
        if (!TryOptionalDouble(source, "backward", 0, MaxForecast, halfIncrement: true, out var backward)) return false;
        if (backward.HasValue) trendline.Backward = backward.Value;
        if (!TryOptionalDouble(source, "intercept", -MaxSafeInteger, MaxSafeInteger, halfIncrement: false, out var intercept)) return false;
        if (intercept.HasValue) trendline.Intercept = intercept.Value;
        if (!TryOptionalBoolean(source, "dispRSqr", out var displayRSquared)) return false;
        trendline.DisplayRSquared = displayRSquared ?? false;
        if (!TryOptionalBoolean(source, "dispEq", out var displayEquation)) return false;
        trendline.DisplayEquation = displayEquation ?? false;
        return true;
    }

    private static XElement Element(SpreadsheetChartTrendlineArtifact trendline)
    {
        var output = new XElement(ChartNs + "trendline");
        if (trendline.Name.Length > 0) output.Add(new XElement(ChartNs + "name", trendline.Name));
        var line = XlsxChartSeriesLineStyleCodec.Element(trendline.Line);
        if (line is not null) output.Add(new XElement(ChartNs + "spPr", line));
        output.Add(Scalar("trendlineType", TypeValue(trendline.Type)));
        if (trendline.HasPolynomialOrder) output.Add(Scalar("order", trendline.PolynomialOrder.ToString(CultureInfo.InvariantCulture)));
        if (trendline.HasPeriod) output.Add(Scalar("period", trendline.Period.ToString(CultureInfo.InvariantCulture)));
        if (trendline.HasForward) output.Add(Scalar("forward", trendline.Forward.ToString("R", CultureInfo.InvariantCulture)));
        if (trendline.HasBackward) output.Add(Scalar("backward", trendline.Backward.ToString("R", CultureInfo.InvariantCulture)));
        if (trendline.HasIntercept) output.Add(Scalar("intercept", trendline.Intercept.ToString("R", CultureInfo.InvariantCulture)));
        if (trendline.DisplayRSquared) output.Add(Scalar("dispRSqr", "1"));
        if (trendline.DisplayEquation) output.Add(Scalar("dispEq", "1"));
        return output;
    }

    private static string Semantics(SpreadsheetChartTrendlineArtifact trendline) => string.Join(':',
        (int)trendline.Type,
        trendline.Name,
        trendline.HasPolynomialOrder ? trendline.PolynomialOrder.ToString(CultureInfo.InvariantCulture) : "no-order",
        trendline.HasPeriod ? trendline.Period.ToString(CultureInfo.InvariantCulture) : "no-period",
        trendline.HasForward ? trendline.Forward.ToString("R", CultureInfo.InvariantCulture) : "no-forward",
        trendline.HasBackward ? trendline.Backward.ToString("R", CultureInfo.InvariantCulture) : "no-backward",
        trendline.HasIntercept ? trendline.Intercept.ToString("R", CultureInfo.InvariantCulture) : "no-intercept",
        trendline.DisplayEquation ? "equation" : "no-equation",
        trendline.DisplayRSquared ? "r-squared" : "no-r-squared",
        XlsxChartSeriesLineStyleCodec.Semantics(trendline.Line));

    private static void ValidateForecast(bool present, double value, string name, string worksheetId, string chartId, string seriesName, int index)
    {
        if (!present) return;
        if (!double.IsFinite(value) || value < 0 || value > MaxForecast || Math.Abs(value * 2 - Math.Round(value * 2)) > 1e-9)
            throw Invalid(worksheetId, chartId, seriesName, $"trendline {index + 1} {name} must be from 0 through {MaxForecast} in 0.5 increments");
    }

    private static bool TryOptionalDouble(XElement owner, string name, double minimum, double maximum, bool halfIncrement, out double? value)
    {
        value = null;
        var element = owner.Element(ChartNs + name);
        if (element is null) return true;
        if (!TryScalar(element, out var text) || !double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) || !double.IsFinite(parsed) || parsed < minimum || parsed > maximum || halfIncrement && Math.Abs(parsed * 2 - Math.Round(parsed * 2)) > 1e-9) return false;
        value = parsed;
        return true;
    }

    private static bool TryOptionalBoolean(XElement owner, string name, out bool? value)
    {
        value = null;
        var element = owner.Element(ChartNs + name);
        if (element is null) return true;
        if (!TryScalar(element, out var text)) return false;
        value = text switch { "1" or "true" => true, "0" or "false" => false, _ => null };
        return value.HasValue;
    }

    private static bool TryText(XElement element, out string value)
    {
        value = string.Empty;
        if (element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || element.Elements().Any() || element.Nodes().Any(node => node is not XText)) return false;
        value = element.Value;
        return true;
    }

    private static bool TryScalar(XElement? element, out string value)
    {
        value = string.Empty;
        if (element is null || element.Elements().Any() || element.Nodes().Any(node => node is not XText text || !string.IsNullOrWhiteSpace(text.Value))) return false;
        var attributes = element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        if (attributes.Length != 1 || attributes[0].Name != "val") return false;
        value = attributes[0].Value;
        return value.Length > 0;
    }

    private static XElement Scalar(string name, string value) => new(ChartNs + name, new XAttribute("val", value));

    private static bool Supported(SpreadsheetChartTrendlineType type) => type is
        SpreadsheetChartTrendlineType.Exponential or SpreadsheetChartTrendlineType.Linear or
        SpreadsheetChartTrendlineType.Logarithmic or SpreadsheetChartTrendlineType.MovingAverage or
        SpreadsheetChartTrendlineType.Polynomial or SpreadsheetChartTrendlineType.Power;

    private static bool TryType(string value, out SpreadsheetChartTrendlineType type)
    {
        type = value switch
        {
            "exp" => SpreadsheetChartTrendlineType.Exponential,
            "linear" => SpreadsheetChartTrendlineType.Linear,
            "log" => SpreadsheetChartTrendlineType.Logarithmic,
            "movingAvg" => SpreadsheetChartTrendlineType.MovingAverage,
            "poly" => SpreadsheetChartTrendlineType.Polynomial,
            "power" => SpreadsheetChartTrendlineType.Power,
            _ => SpreadsheetChartTrendlineType.Unspecified,
        };
        return type != SpreadsheetChartTrendlineType.Unspecified;
    }

    private static string TypeValue(SpreadsheetChartTrendlineType type) => type switch
    {
        SpreadsheetChartTrendlineType.Exponential => "exp",
        SpreadsheetChartTrendlineType.Linear => "linear",
        SpreadsheetChartTrendlineType.Logarithmic => "log",
        SpreadsheetChartTrendlineType.MovingAverage => "movingAvg",
        SpreadsheetChartTrendlineType.Polynomial => "poly",
        SpreadsheetChartTrendlineType.Power => "power",
        _ => throw new InvalidOperationException("Validated chart trendline type changed unexpectedly."),
    };

    private static bool HasControls(string value) => value.Any(char.IsControl);
    private static bool HasUnexpectedNodes(XElement owner) => owner.Nodes().Any(node => node switch
    {
        XElement => false,
        XText text => !string.IsNullOrWhiteSpace(text.Value),
        _ => true,
    });
    private static CodecException Invalid(string worksheetId, string chartId, string seriesName, string message) =>
        new("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chartId} series {seriesName} {message}.");
    private static CodecException Topology(string code, string subject, string message) => new(code, $"{subject} {message}.");
}
