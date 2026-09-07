using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns the bounded c:numCache/c:numLit/c:formatCode branch used by a
// series' numeric Y values. "General" is the native default and therefore
// stays absent from the PPJ surface; X and bubble-size format codes remain
// source-owned until their own public fields are introduced.
internal static class XlsxChartSeriesValueFormatCodec
{
    internal const string NativeDefault = "General";

    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    internal static bool TryRead(XElement cache, out string formatCode)
    {
        formatCode = string.Empty;
        var elements = cache.Elements(ChartNs + "formatCode").ToArray();
        if (elements.Length == 0) return true;
        if (elements.Length != 1) return false;
        var element = elements[0];
        if (element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || element.HasElements ||
            element.Nodes().Any(node => node is not XText)) return false;
        formatCode = element.Value;
        return formatCode.Length is > 0 and <= 255 && !formatCode.Any(char.IsControl);
    }

    internal static string PpjValue(string formatCode) =>
        string.Equals(formatCode, NativeDefault, StringComparison.Ordinal) ? string.Empty : formatCode;

    internal static void Validate(SpreadsheetChartSeriesArtifact series, string worksheetId, string chartId)
    {
        if (!series.HasValuesFormatCode) return;
        if (series.ValuesFormatCode.Length is < 1 or > 255 || series.ValuesFormatCode.Any(char.IsControl))
            throw new CodecException(
                "invalid_spreadsheet_chart",
                $"Worksheet {worksheetId} chart {chartId} series {series.Name} has an invalid values format code.");
    }

    internal static void Patch(XElement cache, string formatCode, string errorCode, string subject)
    {
        if (formatCode.Length is < 1 or > 255 || formatCode.Any(char.IsControl))
            throw new CodecException(errorCode, $"{subject} has an invalid series values format code.");
        if (!TryRead(cache, out _))
            throw new CodecException(errorCode, $"{subject} has a series values format code outside the bounded profile.");
        var element = cache.Element(ChartNs + "formatCode");
        if (element is null)
        {
            element = new XElement(ChartNs + "formatCode");
            cache.AddFirst(element);
        }
        element.Value = formatCode;
    }

    internal static string Semantics(SpreadsheetChartSeriesArtifact series) =>
        series.HasValuesFormatCode ? series.ValuesFormatCode : NativeDefault;
}
