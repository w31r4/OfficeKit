using System.Text.Json;
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
        if (label.HasText && label.RichText is not null)
            throw new CodecException("invalid_spreadsheet_chart", "A trendline label cannot combine literal and structured text.");
        OpenXmlChartRichTextCodec.Validate(label.RichText);
        if (label.HasText && !ValidText(label.Text) || label.HasNumberFormatCode && !ValidText(label.NumberFormatCode))
            throw new CodecException("invalid_spreadsheet_chart", $"Chart {chart} trendline label text/number format must contain 1 through 255 characters without controls.");
        if (label.NumberFormatLink is not (SpreadsheetChartNumberFormatLink.Unspecified or SpreadsheetChartNumberFormatLink.Source or SpreadsheetChartNumberFormatLink.Omitted) ||
            label.NumberFormatLink != SpreadsheetChartNumberFormatLink.Unspecified && !label.HasNumberFormatCode)
            throw new CodecException("invalid_spreadsheet_chart", "Trendline label format linkage requires a number format and a supported link state.");
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
            if (XlsxChartSeriesDataLabelsCodec.TryReadPointText(text, out var literal)) label.Text = literal;
            else
            {
                if (!OpenXmlChartRichTextCodec.TryRead(text, out var rich)) return false;
                label.RichText = rich;
                OpenXmlChartRichTextCodec.NormalizeLabel(label);
            }
        }
        if (source.Element(C + "numFmt") is { } format)
        {
            var attributes = format.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
            var code = (string?)format.Attribute("formatCode");
            var linked = (string?)format.Attribute("sourceLinked");
            if (format.HasElements || UnexpectedNodes(format) || attributes.Length is < 1 or > 2 ||
                attributes.Any(attribute => attribute.Name != "formatCode" && attribute.Name != "sourceLinked") ||
                code is null || !ValidText(code) || linked is not (null or "0" or "false" or "1" or "true")) return false;
            label.NumberFormatCode = code;
            label.NumberFormatLink = linked switch
            {
                null => SpreadsheetChartNumberFormatLink.Omitted,
                "1" or "true" => SpreadsheetChartNumberFormatLink.Source,
                _ => SpreadsheetChartNumberFormatLink.Unspecified,
            };
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
            label.RichText is null ? null : OpenXmlChartRichTextCodec.Element(label.RichText),
            label.HasNumberFormatCode ? new XElement(C + "numFmt", new XAttribute("formatCode", label.NumberFormatCode),
                label.NumberFormatLink == SpreadsheetChartNumberFormatLink.Omitted ? null :
                    new XAttribute("sourceLinked", label.NumberFormatLink == SpreadsheetChartNumberFormatLink.Source ? "1" : "0")) : null,
            label.Fill is not null || label.Line is not null ? XlsxChartSeriesDataLabelsCodec.PointPropertiesElement(label.Fill, label.Line) : null,
            label.TextStyle is not null ? XlsxChartTextStyleCodec.TextPropertiesElement(label.TextStyle) : null);

    internal static string Semantics(SpreadsheetChartTrendlineLabelArtifact? label) =>
        label is null ? "no-label" : Convert.ToBase64String(label.ToByteArray());

    internal static SpreadsheetChartNumberFormatLink NumberFormatLinkFromPpj(JsonElement label)
    {
        if (!label.TryGetProperty("numberFormatSourceLinked", out var linked)) return SpreadsheetChartNumberFormatLink.Unspecified;
        if (!label.TryGetProperty("numberFormat", out _))
            throw new CodecException("invalid_spreadsheet_chart", "Trendline label numberFormatSourceLinked requires numberFormat.");
        return linked.ValueKind switch
        {
            JsonValueKind.Null => SpreadsheetChartNumberFormatLink.Omitted,
            JsonValueKind.True => SpreadsheetChartNumberFormatLink.Source,
            JsonValueKind.False => SpreadsheetChartNumberFormatLink.Unspecified,
            _ => throw new CodecException("invalid_spreadsheet_chart", "Trendline label numberFormatSourceLinked must be boolean or null."),
        };
    }

    private static bool ValidText(string text) => text.Length is > 0 and <= 255 && !text.Any(char.IsControl);
    private static bool UnexpectedNodes(XElement element) => element.Nodes().Any(node => node switch
    {
        XElement => false,
        XText text => !string.IsNullOrWhiteSpace(text.Value),
        _ => true,
    });
}
