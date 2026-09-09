using System.Xml.Linq;
using Google.Protobuf;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

internal static class OpenXmlChartTrendlineLabelCodec
{
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XName[] Children = [C + "layout", C + "tx", C + "numFmt", C + "spPr", C + "txPr"];

    internal static void Validate(SpreadsheetChartTrendlineLabelArtifact? label, string worksheet, string chart, string series)
    {
        if (label is null) return;
        if (label.HasText && !ValidText(label.Text) || label.HasNumberFormatCode && !ValidText(label.NumberFormatCode))
            throw new CodecException("invalid_spreadsheet_chart", $"Chart {chart} trendline label text/number format must contain 1 through 255 characters without controls.");
        XlsxChartTextStyleCodec.ValidateStyle(label.TextStyle, worksheet, chart, "trendline label text style");
        XlsxChartSurfaceFillCodec.Validate(label.Fill, "trendline label fill");
        XlsxChartSeriesLineStyleCodec.ValidateLine(label.Line, worksheet, chart, series, "trendline label line");
        OpenXmlChartLayoutCodec.Validate(label.Layout);
    }

    internal static bool TryRead(XElement source, out SpreadsheetChartTrendlineLabelArtifact label)
    {
        label = new SpreadsheetChartTrendlineLabelArtifact();
        if (source.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || UnexpectedNodes(source)) return false;
        var previous = -1;
        foreach (var child in source.Elements())
        {
            var index = Array.IndexOf(Children, child.Name);
            if (index <= previous) return false;
            previous = index;
        }
        if (source.Element(C + "layout") is { } layout)
        {
            if (!OpenXmlChartLayoutCodec.TryRead(layout, out var value)) return false;
            label.Layout = value;
        }
        if (source.Element(C + "tx") is { } text)
        {
            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            if (text.DescendantNodes().Any(node => node is not (XElement or XText) ||
                node is XText literalText && literalText.Parent?.Name != drawing + "t" && !string.IsNullOrWhiteSpace(literalText.Value))) return false;
            if (!XlsxChartSeriesDataLabelsCodec.TryReadPointText(text, out var literal)) return false;
            label.Text = literal;
        }
        if (source.Element(C + "numFmt") is { } format)
        {
            var attributes = format.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
            var code = (string?)format.Attribute("formatCode");
            if (format.HasElements || UnexpectedNodes(format) || attributes.Length != 2 ||
                attributes.Any(attribute => attribute.Name != "formatCode" && attribute.Name != "sourceLinked") ||
                code is null || !ValidText(code) || (string?)format.Attribute("sourceLinked") is not ("0" or "false")) return false;
            label.NumberFormatCode = code;
        }
        if (!XlsxChartTextStyleCodec.TryReadTextProperties(source, out var textStyle)) return false;
        label.TextStyle = textStyle;
        if (source.Element(C + "spPr") is { } properties)
        {
            if (!XlsxChartSeriesDataLabelsCodec.TryReadPointProperties(properties, out var fill, out var line)) return false;
            label.Fill = fill;
            label.Line = line;
        }
        return true;
    }

    internal static XElement? Element(SpreadsheetChartTrendlineLabelArtifact? label) => label is null ? null :
        new XElement(C + "trendlineLbl",
            OpenXmlChartLayoutCodec.Element(label.Layout),
            label.HasText ? XlsxChartSeriesDataLabelsCodec.PointTextElement(label.Text) : null,
            label.HasNumberFormatCode ? new XElement(C + "numFmt", new XAttribute("formatCode", label.NumberFormatCode), new XAttribute("sourceLinked", "0")) : null,
            label.Fill is not null || label.Line is not null ? XlsxChartSeriesDataLabelsCodec.PointPropertiesElement(label.Fill, label.Line) : null,
            label.TextStyle is not null ? XlsxChartTextStyleCodec.TextPropertiesElement(label.TextStyle) : null);

    internal static string Semantics(SpreadsheetChartTrendlineLabelArtifact? label) =>
        label is null ? "no-label" : Convert.ToBase64String(label.ToByteArray());

    private static bool ValidText(string text) => text.Length is > 0 and <= 255 && !text.Any(char.IsControl);
    private static bool UnexpectedNodes(XElement element) => element.Nodes().Any(node => node switch
    {
        XElement => false,
        XText text => !string.IsNullOrWhiteSpace(text.Value),
        _ => true,
    });
}
