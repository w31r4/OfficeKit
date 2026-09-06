using System.Globalization;
using System.Xml.Linq;
using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

// Owns one exact chart-title / axis-tick DrawingML text profile. Font identity,
// emphasis, paragraph alignment, direct text paint, and point size are
// editable; unrecognized rich text graphs keep the containing chart
// source-owned instead of being normalized.
internal static class XlsxChartTextStyleCodec
{
    private const double MinimumFontSizePoints = 1;
    private const double MaximumFontSizePoints = 4_000;
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly HashSet<string> UnderlineValues = new(StringComparer.Ordinal)
    {
        "none", "words", "sng", "dbl", "heavy", "dotted", "dottedHeavy", "dash", "dashHeavy", "dashLong", "dashLongHeavy",
        "dotDash", "dotDashHeavy", "dotDotDash", "dotDotDashHeavy", "wavy", "wavyHeavy", "wavyDbl",
    };
    private static readonly HashSet<string> AlignmentValues = new(StringComparer.Ordinal)
    {
        "l", "ctr", "r", "just",
    };

    internal static void Validate(SpreadsheetChartArtifact chart, string worksheetId)
    {
        ValidateStyle(chart.TitleTextStyle, worksheetId, chart.Id, "title_text_style");
        if (chart.TitleTextStyle is not null && chart.Title.Length == 0)
            throw Invalid(worksheetId, chart.Id, "title_text_style requires a non-empty title.");
        ValidateStyle(chart.LegendTextStyle, worksheetId, chart.Id, "legend_text_style");
        if (chart.LegendTextStyle is not null && !chart.HasLegend)
            throw Invalid(worksheetId, chart.Id, "legend_text_style requires an enabled legend.");
        ValidateStyle(chart.DataLabels?.TextStyle, worksheetId, chart.Id, "data_labels.text_style");
        ValidateStyle(chart.XAxis?.TextStyle, worksheetId, chart.Id, "x_axis.text_style");
        ValidateStyle(chart.YAxis?.TextStyle, worksheetId, chart.Id, "y_axis.text_style");
        ValidateAxisTitleStyle(chart.XAxis, worksheetId, chart.Id, "x_axis.title_text_style");
        ValidateAxisTitleStyle(chart.YAxis, worksheetId, chart.Id, "y_axis.title_text_style");
    }

    internal static bool TryReadTitle(XElement title, SpreadsheetChartArtifact chart)
    {
        if (!TryReadTitleStyle(title, out var style)) return false;
        if (style is not null) chart.TitleTextStyle = style;
        return true;
    }

    internal static bool TryReadAxis(XElement axis, SpreadsheetChartAxisArtifact semantic)
    {
        if (!TryReadTextProperties(axis, out var style)) return false;
        if (style is not null) semantic.TextStyle = style;
        var title = axis.Element(ChartNs + "title");
        if (title is not null)
        {
            if (!TryReadTitleStyle(title, out var titleStyle)) return false;
            if (titleStyle is not null) semantic.TitleTextStyle = titleStyle;
        }
        return true;
    }

    internal static bool TryReadTextProperties(XElement owner, out SpreadsheetChartTextStyleArtifact? style)
    {
        style = null;
        var properties = owner.Elements(ChartNs + "txPr").Take(2).ToArray();
        if (properties.Length == 0) return true;
        if (properties.Length != 1 || !TryExactAxisTextProperties(properties[0], out var parsed)) return false;
        style = parsed;
        return true;
    }

    internal static XElement TitleElement(
        string title,
        SpreadsheetChartTextStyleArtifact? style,
        string titlePlacement = "")
    {
        var overlay = titlePlacement switch
        {
            "" => null,
            "aboveChart" => new XElement(ChartNs + "overlay", new XAttribute("val", "0")),
            "centeredOverlay" => new XElement(ChartNs + "overlay", new XAttribute("val", "1")),
            _ => throw new CodecException("invalid_chart_title", $"Unsupported chart title placement {titlePlacement}."),
        };
        var run = new XElement(DrawingNs + "r");
        if (style is not null && HasCharacterStyle(style)) run.Add(StyleProperties("rPr", style));
        run.Add(new XElement(DrawingNs + "t", title));
        var paragraph = new XElement(DrawingNs + "p");
        if (style is not null && style.Alignment.Length > 0)
            paragraph.Add(new XElement(DrawingNs + "pPr", new XAttribute("algn", style.Alignment)));
        paragraph.Add(run);
        return new XElement(ChartNs + "title",
            new XElement(ChartNs + "tx", new XElement(ChartNs + "rich",
                new XElement(DrawingNs + "bodyPr"), new XElement(DrawingNs + "lstStyle"),
                paragraph)),
            new XElement(ChartNs + "layout"),
            overlay);
    }

    internal static void AppendAuthoredAxis(XElement axis, SpreadsheetChartTextStyleArtifact? style)
    {
        if (style is not null) axis.Add(AxisTextProperties(style));
    }

    internal static void PatchTitle(XElement title, SpreadsheetChartTextStyleArtifact? style)
    {
        if (!TryExactTitleRun(title, out var run, out var paragraphProperties)) throw ReadOnly("title");
        var existing = run.Element(DrawingNs + "rPr");
        if (style is null)
        {
            existing?.Remove();
            paragraphProperties?.Remove();
            return;
        }
        if (existing is not null && !TryExactStyleProperties(existing, out _)) throw ReadOnly("title");
        if (HasCharacterStyle(style))
        {
            var replacement = StyleProperties("rPr", style);
            if (existing is null) run.AddFirst(replacement);
            else existing.ReplaceWith(replacement);
        }
        else
            existing?.Remove();
        if (style.Alignment.Length > 0)
        {
            var replacement = new XElement(DrawingNs + "pPr", new XAttribute("algn", style.Alignment));
            if (paragraphProperties is null) run.AddBeforeSelf(replacement);
            else paragraphProperties.ReplaceWith(replacement);
        }
        else
            paragraphProperties?.Remove();
    }

    internal static void PatchAxis(XElement axis, SpreadsheetChartTextStyleArtifact? style)
    {
        var existing = axis.Element(ChartNs + "txPr");
        if (style is null) { existing?.Remove(); return; }
        if (existing is not null && !TryExactAxisTextProperties(existing, out _)) throw ReadOnly("axis");
        var replacement = AxisTextProperties(style);
        if (existing is null)
        {
            var crossAxis = axis.Element(ChartNs + "crossAx");
            if (crossAxis is null) axis.Add(replacement);
            else crossAxis.AddBeforeSelf(replacement);
        }
        else existing.ReplaceWith(replacement);
    }

    internal static XElement TextPropertiesElement(SpreadsheetChartTextStyleArtifact style) => AxisTextProperties(style);

    internal static void PatchTextProperties(
        XElement owner,
        SpreadsheetChartTextStyleArtifact? style,
        IReadOnlySet<string> laterNames)
    {
        var existing = owner.Element(ChartNs + "txPr");
        if (style is null) { existing?.Remove(); return; }
        if (existing is not null && !TryExactAxisTextProperties(existing, out _)) throw ReadOnly(owner.Name.LocalName);
        var replacement = AxisTextProperties(style);
        if (existing is not null) { existing.ReplaceWith(replacement); return; }
        var next = owner.Elements().FirstOrDefault(element => laterNames.Contains(element.Name.LocalName));
        if (next is null) owner.Add(replacement);
        else next.AddBeforeSelf(replacement);
    }

    internal static string Semantics(SpreadsheetChartTextStyleArtifact? style)
    {
        if (style is null) return "-";
        return string.Join(':',
            style.HasFontSizePoints ? style.FontSizePoints.ToString("R", CultureInfo.InvariantCulture) : "default-size",
            style.FontFamily.Length > 0 ? style.FontFamily : "default-latin",
            style.FontFamilyEastAsia.Length > 0 ? style.FontFamilyEastAsia : "default-eastAsia",
            style.FontFamilyComplexScript.Length > 0 ? style.FontFamilyComplexScript : "default-complexScript",
            style.HasBold ? style.Bold.ToString(CultureInfo.InvariantCulture) : "default-bold",
            style.HasItalic ? style.Italic.ToString(CultureInfo.InvariantCulture) : "default-italic",
            style.Alignment.Length > 0 ? style.Alignment : "default-alignment",
            style.Fill is not null ? XlsxChartSurfaceFillCodec.Semantics(style.Fill) : "default-fill",
            style.Underline.Length > 0 ? style.Underline : "default-underline",
            style.ColorRgb.Length > 0 ? style.ColorRgb.ToUpperInvariant() : "default-color",
            style.HasOpacityThousandthPercent ? style.OpacityThousandthPercent.ToString(CultureInfo.InvariantCulture) : "default-alpha");
    }

    private static bool TryReadTitleStyle(XElement title, out SpreadsheetChartTextStyleArtifact? style)
    {
        style = null;
        if (!TryExactTitleRun(title, out var run, out var paragraphProperties)) return false;
        var properties = run.Element(DrawingNs + "rPr");
        if (properties is not null)
        {
            if (!TryExactStyleProperties(properties, out var parsed)) return false;
            style = parsed;
        }
        if (paragraphProperties is not null)
        {
            style ??= new SpreadsheetChartTextStyleArtifact();
            style.Alignment = paragraphProperties.Attribute("algn")!.Value;
        }
        if (style is not null && !HasAnyStyle(style)) return false;
        return true;
    }

    private static void ValidateAxisTitleStyle(SpreadsheetChartAxisArtifact? axis, string worksheetId, string chartId, string field)
    {
        ValidateStyle(axis?.TitleTextStyle, worksheetId, chartId, field);
        if (axis?.TitleTextStyle is not null && axis.Title.Length == 0)
            throw Invalid(worksheetId, chartId, $"{field} requires a non-empty axis title.");
    }

    internal static void ValidateStyle(SpreadsheetChartTextStyleArtifact? style, string worksheetId, string chartId, string field)
    {
        if (style is null) return;
        if (!style.HasFontSizePoints && style.FontFamily.Length == 0 && style.FontFamilyEastAsia.Length == 0 && style.FontFamilyComplexScript.Length == 0 &&
            !style.HasBold && !style.HasItalic && style.Alignment.Length == 0 && style.Underline.Length == 0 && style.ColorRgb.Length == 0 && !style.HasOpacityThousandthPercent && style.Fill is null)
            throw Invalid(worksheetId, chartId, $"{field} must declare at least one bounded property.");
        if (style.HasFontSizePoints && (!double.IsFinite(style.FontSizePoints) || style.FontSizePoints < MinimumFontSizePoints || style.FontSizePoints > MaximumFontSizePoints))
            throw Invalid(worksheetId, chartId, $"{field}.font_size_points must be from 1 through 4000.");
        ValidateTypeface(style.FontFamily, worksheetId, chartId, field + ".font_family");
        ValidateTypeface(style.FontFamilyEastAsia, worksheetId, chartId, field + ".font_family_east_asia");
        ValidateTypeface(style.FontFamilyComplexScript, worksheetId, chartId, field + ".font_family_complex_script");
        if (style.Fill is not null)
        {
            XlsxChartSurfaceFillCodec.Validate(style.Fill, field + ".fill");
            if (style.ColorRgb.Length > 0 || style.HasOpacityThousandthPercent)
                throw Invalid(worksheetId, chartId, $"{field}.fill cannot be combined with color.");
        }
        if (style.Alignment.Length > 0 && !AlignmentValues.Contains(style.Alignment))
            throw Invalid(worksheetId, chartId, $"{field}.alignment must be a bounded DrawingML paragraph-alignment token.");
        if (style.Underline.Length > 0 && !UnderlineValues.Contains(style.Underline))
            throw Invalid(worksheetId, chartId, $"{field}.underline must be a bounded DrawingML underline token.");
        if (style.ColorRgb.Length > 0 && (style.ColorRgb.Length != 6 || !style.ColorRgb.All(Uri.IsHexDigit)))
            throw Invalid(worksheetId, chartId, $"{field}.color_rgb must be a six-digit RGB color.");
        if (style.HasOpacityThousandthPercent && (style.ColorRgb.Length == 0 || style.OpacityThousandthPercent > 100_000))
            throw Invalid(worksheetId, chartId, $"{field}.opacity requires a color and must be between 0% and 100%.");
    }

    private static void ValidateTypeface(string value, string worksheetId, string chartId, string field)
    {
        if (value.Length > 255 || value.Any(char.IsControl))
            throw Invalid(worksheetId, chartId, $"{field} must contain at most 255 characters without controls.");
    }

    private static bool TryExactTitleRun(XElement title, out XElement run, out XElement? paragraphProperties)
    {
        run = null!;
        paragraphProperties = null;
        var titleChildren = title.Elements().ToArray();
        if (titleChildren.Any(item => item.Name != ChartNs + "tx" && item.Name != ChartNs + "layout" && item.Name != ChartNs + "overlay") || titleChildren.Count(item => item.Name == ChartNs + "tx") != 1) return false;
        var tx = title.Element(ChartNs + "tx")!;
        var rich = tx.Element(ChartNs + "rich");
        if (rich is null || tx.Elements().Count() != 1) return false;
        var richChildren = rich.Elements().ToArray();
        if (richChildren.Length != 3 || richChildren[0].Name != DrawingNs + "bodyPr" || richChildren[1].Name != DrawingNs + "lstStyle" || richChildren[2].Name != DrawingNs + "p" ||
            richChildren[0].HasAttributes || richChildren[0].HasElements || richChildren[1].HasAttributes || richChildren[1].HasElements) return false;
        var paragraphChildren = richChildren[2].Elements().ToArray();
        if (paragraphChildren.Length is < 1 or > 2 || paragraphChildren[^1].Name != DrawingNs + "r") return false;
        if (paragraphChildren.Length == 2)
        {
            paragraphProperties = paragraphChildren[0];
            var attributes = paragraphProperties.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
            if (paragraphProperties.HasElements || attributes.Length != 1 || attributes[0].Name != "algn" || !AlignmentValues.Contains(attributes[0].Value))
                return false;
        }
        run = paragraphChildren[^1];
        var runChildren = run.Elements().ToArray();
        return runChildren.Length is >= 1 and <= 2 && runChildren[^1].Name == DrawingNs + "t" &&
            (runChildren.Length != 2 || runChildren[0].Name == DrawingNs + "rPr");
    }

    private static bool TryExactAxisTextProperties(XElement properties, out SpreadsheetChartTextStyleArtifact style)
    {
        style = new SpreadsheetChartTextStyleArtifact();
        var children = properties.Elements().ToArray();
        if (children.Length != 3 || children[0].Name != DrawingNs + "bodyPr" || children[1].Name != DrawingNs + "lstStyle" || children[2].Name != DrawingNs + "p" ||
            children[0].HasAttributes || children[0].HasElements || children[1].HasAttributes || children[1].HasElements) return false;
        var paragraphChildren = children[2].Elements().ToArray();
        if (paragraphChildren.Length != 2 || paragraphChildren[0].Name != DrawingNs + "pPr" || paragraphChildren[1].Name != DrawingNs + "endParaRPr" ||
            paragraphChildren[1].HasAttributes || paragraphChildren[1].HasElements) return false;
        var paragraphProperties = paragraphChildren[0];
        var defaults = paragraphProperties.Elements().ToArray();
        var attributes = paragraphProperties.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        if (attributes.Any(attribute => attribute.Name != "algn") ||
            attributes.Any(attribute => attribute.Name == "algn" && !AlignmentValues.Contains(attribute.Value)) ||
            defaults.Length > 1 || defaults.Any(defaultsChild => defaultsChild.Name != DrawingNs + "defRPr"))
            return false;
        if (defaults.Length == 1 && !TryExactStyleProperties(defaults[0], out style)) return false;
        if (paragraphProperties.Attribute("algn") is { } alignment) style.Alignment = alignment.Value;
        return defaults.Length == 1 || style.Alignment.Length > 0;
    }

    private static bool TryExactStyleProperties(XElement properties, out SpreadsheetChartTextStyleArtifact style)
    {
        style = new SpreadsheetChartTextStyleArtifact();
        var allowedAttributes = new HashSet<XName> { "sz", "b", "i", "u" };
        var attributes = properties.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        if (attributes.Any(attribute => !allowedAttributes.Contains(attribute.Name))) return false;
        if (properties.Attribute("sz") is { } size)
        {
            if (!uint.TryParse(size.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var raw) || raw is < 100 or > 400_000) return false;
            style.FontSizePoints = raw / 100d;
        }
        if (properties.Attribute("b") is { } bold)
        {
            if (!TryBoolean(bold.Value, out var value)) return false;
            style.Bold = value;
        }
        if (properties.Attribute("i") is { } italic)
        {
            if (!TryBoolean(italic.Value, out var value)) return false;
            style.Italic = value;
        }
        if (properties.Attribute("u") is { } underline)
        {
            if (!UnderlineValues.Contains(underline.Value)) return false;
            style.Underline = underline.Value;
        }
        var children = properties.Elements().ToArray();
        var index = 0;
        if (index < children.Length && IsPaint(children[index]))
        {
            if (!XlsxChartSurfaceFillCodec.TryReadPaint(children[index++], out var fill) || fill is null) return false;
            ApplyParsedFill(style, fill);
        }
        if (index < children.Length && children[index].Name == DrawingNs + "latin")
        {
            if (!TryTypeface(children[index++], out var value)) return false;
            style.FontFamily = value;
        }
        if (index < children.Length && children[index].Name == DrawingNs + "ea")
        {
            if (!TryTypeface(children[index++], out var value)) return false;
            style.FontFamilyEastAsia = value;
        }
        if (index < children.Length && children[index].Name == DrawingNs + "cs")
        {
            if (!TryTypeface(children[index++], out var value)) return false;
            style.FontFamilyComplexScript = value;
        }
        return index == children.Length &&
            (style.HasFontSizePoints || style.FontFamily.Length > 0 || style.FontFamilyEastAsia.Length > 0 || style.FontFamilyComplexScript.Length > 0 ||
             style.HasBold || style.HasItalic || style.Underline.Length > 0 || style.ColorRgb.Length > 0 || style.Fill is not null);
    }

    private static void ApplyParsedFill(SpreadsheetChartTextStyleArtifact style, SpreadsheetChartSurfaceFill fill)
    {
        if (fill.FillCase == SpreadsheetChartSurfaceFill.FillOneofCase.SolidRgb)
        {
            style.ColorRgb = fill.SolidRgb;
            if (fill.HasOpacityThousandthPercent)
                style.OpacityThousandthPercent = fill.OpacityThousandthPercent;
            return;
        }
        style.Fill = fill;
    }

    private static bool TryTypeface(XElement source, out string value)
    {
        value = string.Empty;
        var attributes = source.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToArray();
        if (source.HasElements || attributes.Length != 1 || attributes[0].Name != "typeface" || attributes[0].Value.Length is 0 or > 255 || attributes[0].Value.Any(char.IsControl)) return false;
        value = attributes[0].Value;
        return true;
    }

    private static XElement AxisTextProperties(SpreadsheetChartTextStyleArtifact style) => new(ChartNs + "txPr",
        new XElement(DrawingNs + "bodyPr"), new XElement(DrawingNs + "lstStyle"),
        new XElement(DrawingNs + "p",
            ParagraphProperties(style),
            new XElement(DrawingNs + "endParaRPr")));

    private static XElement ParagraphProperties(SpreadsheetChartTextStyleArtifact style)
    {
        var output = new XElement(DrawingNs + "pPr");
        if (style.Alignment.Length > 0) output.SetAttributeValue("algn", style.Alignment);
        if (HasCharacterStyle(style)) output.Add(StyleProperties("defRPr", style));
        return output;
    }

    private static bool HasCharacterStyle(SpreadsheetChartTextStyleArtifact style) =>
        style.HasFontSizePoints || style.FontFamily.Length > 0 || style.FontFamilyEastAsia.Length > 0 || style.FontFamilyComplexScript.Length > 0 ||
        style.HasBold || style.HasItalic || style.Underline.Length > 0 || style.ColorRgb.Length > 0 || style.HasOpacityThousandthPercent || style.Fill is not null;

    private static bool HasAnyStyle(SpreadsheetChartTextStyleArtifact style) => HasCharacterStyle(style) || style.Alignment.Length > 0;

    private static XElement StyleProperties(string name, SpreadsheetChartTextStyleArtifact style)
    {
        var output = new XElement(DrawingNs + name);
        if (style.HasFontSizePoints) output.SetAttributeValue("sz", Size(style.FontSizePoints));
        if (style.HasBold) output.SetAttributeValue("b", style.Bold ? "1" : "0");
        if (style.HasItalic) output.SetAttributeValue("i", style.Italic ? "1" : "0");
        if (style.Underline.Length > 0) output.SetAttributeValue("u", style.Underline);
        if (style.Fill is not null)
            output.Add(XlsxChartSurfaceFillCodec.PaintElement(style.Fill, "chart text fill"));
        else if (style.ColorRgb.Length > 0)
        {
            var color = new XElement(DrawingNs + "srgbClr", new XAttribute("val", style.ColorRgb.ToUpperInvariant()));
            if (style.HasOpacityThousandthPercent)
                color.Add(new XElement(DrawingNs + "alpha", new XAttribute("val", style.OpacityThousandthPercent)));
            output.Add(new XElement(DrawingNs + "solidFill", color));
        }
        if (style.FontFamily.Length > 0)
            output.Add(new XElement(DrawingNs + "latin", new XAttribute("typeface", style.FontFamily)));
        if (style.FontFamilyEastAsia.Length > 0)
            output.Add(new XElement(DrawingNs + "ea", new XAttribute("typeface", style.FontFamilyEastAsia)));
        if (style.FontFamilyComplexScript.Length > 0)
            output.Add(new XElement(DrawingNs + "cs", new XAttribute("typeface", style.FontFamilyComplexScript)));
        return output;
    }

    private static bool IsPaint(XElement element) => element.Name == DrawingNs + "noFill" ||
        element.Name == DrawingNs + "solidFill" || element.Name == DrawingNs + "gradFill";

    private static bool TryBoolean(string source, out bool value)
    {
        if (source is "1" or "true") { value = true; return true; }
        if (source is "0" or "false") { value = false; return true; }
        value = false;
        return false;
    }

    private static uint Size(double points) => checked((uint)Math.Round(points * 100, MidpointRounding.AwayFromZero));
    private static CodecException ReadOnly(string subject) =>
        new("unsupported_spreadsheet_chart_edit", $"Referenced worksheet-chart {subject} text styling is read-only.");
    private static CodecException Invalid(string worksheetId, string chartId, string message) =>
        new("invalid_spreadsheet_chart", $"Worksheet {worksheetId} chart {chartId} {message}");
}
