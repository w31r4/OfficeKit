using System.Globalization;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns the one bounded c:errBars projection shared by XLSX and PPTX
// ChartSpace series. Package adapters retain source identity and decide whether
// formula-backed custom data is legal in their host format.
internal static class OpenXmlChartErrorBarsCodec
{
    private const int MaxPoints = 1_048_576;
    private const double MaxSafeInteger = 9_007_199_254_740_991;
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly IReadOnlyDictionary<XName, int> ChildOrder = new Dictionary<XName, int>
    {
        [ChartNs + "errDir"] = 0,
        [ChartNs + "errBarType"] = 1,
        [ChartNs + "errValType"] = 2,
        [ChartNs + "noEndCap"] = 3,
        [ChartNs + "plus"] = 4,
        [ChartNs + "minus"] = 5,
        [ChartNs + "val"] = 6,
        [ChartNs + "spPr"] = 7,
    };

    internal static void Validate(SpreadsheetChartSeriesArtifact series, SpreadsheetChartType chartType, string worksheetId, string chartId)
    {
        var errorBars = series.ErrorBars;
        if (errorBars is null) return;
        if (chartType is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line))
            throw Invalid(worksheetId, chartId, series.Name, "error bars require a bar or line series");
        if (!Supported(errorBars.Direction) || !Supported(errorBars.Type) || !Supported(errorBars.ValueType))
            throw Invalid(worksheetId, chartId, series.Name, "error bars contain an unsupported direction, type, or value type");

        if (errorBars.ValueType == SpreadsheetChartErrorBarValueType.Custom)
        {
            if (errorBars.HasValue) throw Invalid(worksheetId, chartId, series.Name, "custom error bars cannot carry a scalar value");
            ValidateData(errorBars.Plus, required: errorBars.Type != SpreadsheetChartErrorBarType.Minus, series.Values.Count, "plus", worksheetId, chartId, series.Name);
            ValidateData(errorBars.Minus, required: errorBars.Type != SpreadsheetChartErrorBarType.Plus, series.Values.Count, "minus", worksheetId, chartId, series.Name);
        }
        else
        {
            if (errorBars.Plus is not null || errorBars.Minus is not null)
                throw Invalid(worksheetId, chartId, series.Name, "plus/minus numeric sources are valid only for custom error bars");
            if (errorBars.ValueType == SpreadsheetChartErrorBarValueType.StandardError)
            {
                if (errorBars.HasValue) throw Invalid(worksheetId, chartId, series.Name, "standard-error error bars cannot carry a scalar value");
            }
            else if (!errorBars.HasValue || !double.IsFinite(errorBars.Value) || errorBars.Value < 0 || errorBars.Value > MaxSafeInteger)
                throw Invalid(worksheetId, chartId, series.Name, "error-bar value must be finite and within the JavaScript safe-integer range");
        }
        XlsxChartSeriesLineStyleCodec.ValidateLine(errorBars.Line, worksheetId, chartId, series.Name, "error-bar line");
    }

    internal static bool TryRead(XElement nativeSeries, SpreadsheetChartSeriesArtifact series, SpreadsheetChartType chartType)
    {
        var native = nativeSeries.Elements(ChartNs + "errBars").Take(2).ToArray();
        if (native.Length == 0) return true;
        if (native.Length != 1 || chartType is not (SpreadsheetChartType.Bar or SpreadsheetChartType.Line) ||
            !TryRead(native[0], series.Values.Count, out var errorBars)) return false;
        series.ErrorBars = errorBars;
        return true;
    }

    internal static XElement? Element(SpreadsheetChartErrorBarsArtifact? errorBars)
    {
        if (errorBars is null) return null;
        var output = new XElement(ChartNs + "errBars",
            Scalar("errDir", DirectionValue(errorBars.Direction)),
            Scalar("errBarType", TypeValue(errorBars.Type)),
            Scalar("errValType", ValueTypeValue(errorBars.ValueType)));
        if (errorBars.NoEndCap) output.Add(Scalar("noEndCap", "1"));
        if (errorBars.Plus is not null) output.Add(DataElement("plus", errorBars.Plus));
        if (errorBars.Minus is not null) output.Add(DataElement("minus", errorBars.Minus));
        if (errorBars.HasValue) output.Add(Scalar("val", errorBars.Value.ToString("R", CultureInfo.InvariantCulture)));
        var line = XlsxChartSeriesLineStyleCodec.Element(errorBars.Line);
        if (line is not null) output.Add(new XElement(ChartNs + "spPr", line));
        return output;
    }

    internal static void Patch(XElement nativeSeries, SpreadsheetChartSeriesArtifact target, string errorCode, string subject)
    {
        var native = nativeSeries.Elements(ChartNs + "errBars").Take(2).ToArray();
        var expected = target.ErrorBars is null ? 0 : 1;
        if (native.Length != expected) throw Topology(errorCode, subject, "error-bar topology changed unexpectedly");
        if (expected == 0) return;
        if (!TryRead(native[0], target.Values.Count, out var original)) throw Topology(errorCode, subject, "error bars no longer match the editable profile");
        if (!Semantics(original).Equals(Semantics(target.ErrorBars), StringComparison.Ordinal)) native[0].ReplaceWith(Element(target.ErrorBars)!);
    }

    internal static string Semantics(SpreadsheetChartErrorBarsArtifact? errorBars)
    {
        if (errorBars is null) return "no-error-bars";
        return string.Join(':',
            (int)errorBars.Direction,
            (int)errorBars.Type,
            (int)errorBars.ValueType,
            errorBars.HasValue ? errorBars.Value.ToString("R", CultureInfo.InvariantCulture) : "no-value",
            errorBars.NoEndCap ? "no-cap" : "cap",
            DataSemantics(errorBars.Plus),
            DataSemantics(errorBars.Minus),
            XlsxChartSeriesLineStyleCodec.Semantics(errorBars.Line));
    }

    private static bool TryRead(XElement source, int valueCount, out SpreadsheetChartErrorBarsArtifact errorBars)
    {
        errorBars = new SpreadsheetChartErrorBarsArtifact();
        if (source.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || HasUnexpectedNodes(source)) return false;
        var children = source.Elements().ToArray();
        var seen = new HashSet<XName>();
        var order = -1;
        foreach (var child in children)
        {
            if (!ChildOrder.TryGetValue(child.Name, out var current) || current < order || !seen.Add(child.Name)) return false;
            order = current;
        }
        if (!TryScalar(source.Element(ChartNs + "errDir"), out var directionValue) || !TryDirection(directionValue, out var direction) ||
            !TryScalar(source.Element(ChartNs + "errBarType"), out var typeValue) || !TryType(typeValue, out var type) ||
            !TryScalar(source.Element(ChartNs + "errValType"), out var valueTypeValue) || !TryValueType(valueTypeValue, out var valueType)) return false;
        errorBars.Direction = direction;
        errorBars.Type = type;
        errorBars.ValueType = valueType;
        if (!TryOptionalBoolean(source.Element(ChartNs + "noEndCap"), out var noEndCap)) return false;
        errorBars.NoEndCap = noEndCap ?? false;

        var plus = source.Element(ChartNs + "plus");
        var minus = source.Element(ChartNs + "minus");
        var scalar = source.Element(ChartNs + "val");
        if (valueType == SpreadsheetChartErrorBarValueType.Custom)
        {
            if (scalar is not null || (type == SpreadsheetChartErrorBarType.Minus ? plus is not null : plus is null) ||
                (type == SpreadsheetChartErrorBarType.Plus ? minus is not null : minus is null)) return false;
            if (plus is not null)
            {
                if (!TryData(plus, valueCount, out var data)) return false;
                errorBars.Plus = data;
            }
            if (minus is not null)
            {
                if (!TryData(minus, valueCount, out var data)) return false;
                errorBars.Minus = data;
            }
        }
        else
        {
            if (plus is not null || minus is not null) return false;
            if (valueType == SpreadsheetChartErrorBarValueType.StandardError)
            {
                if (scalar is not null) return false;
            }
            else
            {
                if (!TryScalar(scalar, out var value) || !double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) ||
                    !double.IsFinite(parsed) || parsed < 0 || parsed > MaxSafeInteger) return false;
                errorBars.Value = parsed;
            }
        }

        var shape = source.Element(ChartNs + "spPr");
        if (shape is not null)
        {
            if (shape.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || shape.Elements().Any(element => element.Name != DrawingNs + "ln") ||
                HasUnexpectedNodes(shape) || !XlsxChartSeriesLineStyleCodec.TryReadLine(shape, out var line)) return false;
            if (line is not null) errorBars.Line = line;
        }
        return true;
    }

    private static void ValidateData(SpreadsheetChartErrorBarDataArtifact? data, bool required, int valueCount, string side, string worksheetId, string chartId, string seriesName)
    {
        if (!required)
        {
            if (data is not null) throw Invalid(worksheetId, chartId, seriesName, $"{side} data is invalid when the error-bar type excludes {side}");
            return;
        }
        if (data is null) throw Invalid(worksheetId, chartId, seriesName, $"custom error bars require {side} data");
        if (data.Formula.Length > 8_192 || data.Formula.StartsWith('=') || HasControls(data.Formula))
            throw Invalid(worksheetId, chartId, seriesName, $"custom {side} formula must contain at most 8192 characters without a leading equals sign or controls");
        if (data.Values.Count == 0 && data.Formula.Length == 0)
            throw Invalid(worksheetId, chartId, seriesName, $"custom {side} data requires literal values or a formula");
        if (data.Values.Count != 0 && data.Values.Count != valueCount)
            throw Invalid(worksheetId, chartId, seriesName, $"custom {side} cache must contain exactly {valueCount} values");
        if (data.Values.Any(value => !double.IsFinite(value) || value < 0 || value > MaxSafeInteger))
            throw Invalid(worksheetId, chartId, seriesName, $"custom {side} values must be non-negative finite JavaScript-safe numbers");
        if (data.FormatCode.Length > 255 || HasControls(data.FormatCode) || data.FormatCode.Length > 0 && data.Values.Count == 0)
            throw Invalid(worksheetId, chartId, seriesName, $"custom {side} format code requires cached values and at most 255 characters without controls");
    }

    private static bool TryData(XElement source, int valueCount, out SpreadsheetChartErrorBarDataArtifact data)
    {
        data = new SpreadsheetChartErrorBarDataArtifact();
        if (source.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || HasUnexpectedNodes(source)) return false;
        var children = source.Elements().ToArray();
        if (children.Length != 1 || children[0].Name is not { LocalName: "numLit" } and not { LocalName: "numRef" } || children[0].Name.Namespace != ChartNs) return false;
        var branch = children[0];
        if (branch.Name == ChartNs + "numLit")
        {
            if (!TryCache(branch, valueCount, allowEmpty: false, out var values, out var formatCode)) return false;
            data.Values.Add(values);
            data.FormatCode = formatCode;
            return true;
        }
        if (branch.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || HasUnexpectedNodes(branch)) return false;
        var referenceChildren = branch.Elements().ToArray();
        if (referenceChildren.Length is < 1 or > 2 || referenceChildren[0].Name != ChartNs + "f" ||
            referenceChildren.Length == 2 && referenceChildren[1].Name != ChartNs + "numCache" ||
            !TryText(referenceChildren[0], out var formula) || formula.Length is < 1 or > 8_192 || formula.StartsWith('=') || HasControls(formula)) return false;
        data.Formula = formula;
        if (referenceChildren.Length == 2)
        {
            if (!TryCache(referenceChildren[1], valueCount, allowEmpty: true, out var values, out var formatCode)) return false;
            data.Values.Add(values);
            data.FormatCode = formatCode;
        }
        return true;
    }

    private static bool TryCache(XElement cache, int valueCount, bool allowEmpty, out double[] values, out string formatCode)
    {
        values = [];
        formatCode = string.Empty;
        if (cache.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || HasUnexpectedNodes(cache)) return false;
        var children = cache.Elements().ToArray();
        var offset = 0;
        if (children.FirstOrDefault()?.Name == ChartNs + "formatCode")
        {
            if (!TryText(children[0], out formatCode) || formatCode.Length is < 1 or > 255 || HasControls(formatCode)) return false;
            offset = 1;
        }
        if (children.Length <= offset || children[offset].Name != ChartNs + "ptCount" || !TryScalar(children[offset], out var countText) ||
            !int.TryParse(countText, NumberStyles.None, CultureInfo.InvariantCulture, out var count) || count < 0 || count > MaxPoints) return false;
        var points = children.Skip(offset + 1).ToArray();
        if (points.Length != count || (!allowEmpty && count == 0) || count != 0 && count != valueCount || formatCode.Length > 0 && count == 0) return false;
        var output = new double[count];
        for (var index = 0; index < points.Length; index++)
        {
            var point = points[index];
            var attributes = point.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
            var valueElements = point.Elements().ToArray();
            if (point.Name != ChartNs + "pt" || attributes.Length != 1 || attributes[0].Name != "idx" || attributes[0].Value != index.ToString(CultureInfo.InvariantCulture) ||
                valueElements.Length != 1 || valueElements[0].Name != ChartNs + "v" || HasUnexpectedNodes(point) || !TryText(valueElements[0], out var text) ||
                !double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) || !double.IsFinite(parsed) || parsed < 0 || parsed > MaxSafeInteger) return false;
            output[index] = parsed;
        }
        values = output;
        return true;
    }

    private static XElement DataElement(string name, SpreadsheetChartErrorBarDataArtifact data)
    {
        XElement Cache(string cacheName)
        {
            var cache = new XElement(ChartNs + cacheName);
            if (data.FormatCode.Length > 0) cache.Add(new XElement(ChartNs + "formatCode", data.FormatCode));
            cache.Add(new XElement(ChartNs + "ptCount", new XAttribute("val", data.Values.Count)));
            for (var index = 0; index < data.Values.Count; index++)
                cache.Add(new XElement(ChartNs + "pt", new XAttribute("idx", index), new XElement(ChartNs + "v", data.Values[index].ToString("R", CultureInfo.InvariantCulture))));
            return cache;
        }

        if (data.Formula.Length == 0) return new XElement(ChartNs + name, Cache("numLit"));
        var reference = new XElement(ChartNs + "numRef", new XElement(ChartNs + "f", data.Formula));
        if (data.Values.Count > 0) reference.Add(Cache("numCache"));
        return new XElement(ChartNs + name, reference);
    }

    private static string DataSemantics(SpreadsheetChartErrorBarDataArtifact? data) => data is null
        ? "no-data"
        : string.Join(',', data.Formula, data.FormatCode, string.Join(';', data.Values.Select(value => value.ToString("R", CultureInfo.InvariantCulture))));

    private static bool TryOptionalBoolean(XElement? element, out bool? value)
    {
        value = null;
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
    private static bool TryDirection(string value, out SpreadsheetChartErrorBarDirection direction)
    {
        direction = value switch { "x" => SpreadsheetChartErrorBarDirection.X, "y" => SpreadsheetChartErrorBarDirection.Y, _ => SpreadsheetChartErrorBarDirection.Unspecified };
        return Supported(direction);
    }
    private static bool TryType(string value, out SpreadsheetChartErrorBarType type)
    {
        type = value switch { "both" => SpreadsheetChartErrorBarType.Both, "minus" => SpreadsheetChartErrorBarType.Minus, "plus" => SpreadsheetChartErrorBarType.Plus, _ => SpreadsheetChartErrorBarType.Unspecified };
        return Supported(type);
    }
    private static bool TryValueType(string value, out SpreadsheetChartErrorBarValueType type)
    {
        type = value switch
        {
            "cust" => SpreadsheetChartErrorBarValueType.Custom,
            "fixedVal" => SpreadsheetChartErrorBarValueType.FixedValue,
            "percentage" => SpreadsheetChartErrorBarValueType.Percentage,
            "stdDev" => SpreadsheetChartErrorBarValueType.StandardDeviation,
            "stdErr" => SpreadsheetChartErrorBarValueType.StandardError,
            _ => SpreadsheetChartErrorBarValueType.Unspecified,
        };
        return Supported(type);
    }
    private static bool Supported(SpreadsheetChartErrorBarDirection value) => value is SpreadsheetChartErrorBarDirection.X or SpreadsheetChartErrorBarDirection.Y;
    private static bool Supported(SpreadsheetChartErrorBarType value) => value is SpreadsheetChartErrorBarType.Both or SpreadsheetChartErrorBarType.Minus or SpreadsheetChartErrorBarType.Plus;
    private static bool Supported(SpreadsheetChartErrorBarValueType value) => value is SpreadsheetChartErrorBarValueType.Custom or SpreadsheetChartErrorBarValueType.FixedValue or SpreadsheetChartErrorBarValueType.Percentage or SpreadsheetChartErrorBarValueType.StandardDeviation or SpreadsheetChartErrorBarValueType.StandardError;
    private static string DirectionValue(SpreadsheetChartErrorBarDirection value) => value switch { SpreadsheetChartErrorBarDirection.X => "x", SpreadsheetChartErrorBarDirection.Y => "y", _ => throw new InvalidOperationException("Validated error-bar direction changed unexpectedly.") };
    private static string TypeValue(SpreadsheetChartErrorBarType value) => value switch { SpreadsheetChartErrorBarType.Both => "both", SpreadsheetChartErrorBarType.Minus => "minus", SpreadsheetChartErrorBarType.Plus => "plus", _ => throw new InvalidOperationException("Validated error-bar type changed unexpectedly.") };
    private static string ValueTypeValue(SpreadsheetChartErrorBarValueType value) => value switch
    {
        SpreadsheetChartErrorBarValueType.Custom => "cust",
        SpreadsheetChartErrorBarValueType.FixedValue => "fixedVal",
        SpreadsheetChartErrorBarValueType.Percentage => "percentage",
        SpreadsheetChartErrorBarValueType.StandardDeviation => "stdDev",
        SpreadsheetChartErrorBarValueType.StandardError => "stdErr",
        _ => throw new InvalidOperationException("Validated error-bar value type changed unexpectedly."),
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
